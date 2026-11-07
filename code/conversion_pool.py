"""Schedule bounded conversion jobs with one ledger connection per worker."""

from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from contextlib import ExitStack
import json
from pathlib import Path
import time

import psutil

from ledger import Ledger, StorageFull, utc, write_json
from workers import convert_claimed


def owned_private_bytes():
    """Count the converter's owned process tree, including PowerShell and Excel."""
    private = 0
    for process in psutil.Process().children(recursive=True):
        try:
            private += process.memory_info().private
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return private


def run_job(root, obj):
    """Use a worker-local SQLite connection through the normal conversion path."""
    config = json.loads((root / 'config.json').read_bytes())
    database = Ledger(root, config['reserve_bytes'])
    try:
        return convert_claimed(database, obj, config)
    finally:
        database.db.close()


def run(root):
    """Finish active files on pause/stop and prevent concurrent duplicate hashes.

    The global scheduler and per-corpus locks also exclude legacy serial workers
    before interrupted claims are reset. A host failure falls back to one worker
    and leaves the affected input pending for a bounded retry.
    """
    from excel_shrink_research.enron.common import RunLock
    config = json.loads((root / 'config.json').read_bytes())
    roots = {name: Path(path) for name, path in config['corpus_roots'].items()}
    configured = max(1, int(config.get('conversion_workers', 1)))
    effective = configured
    fallback_file = root / 'conversion-fallback.json'
    if fallback_file.exists():
        effective = min(effective, int(json.loads(fallback_file.read_bytes())['workers']))
    reserve = config.get('conversion_memory_reserve_bytes', 1536 * 1024**2)
    budget = config.get('conversion_job_memory_bytes', 256 * 1024**2)
    active = {}
    databases = {}
    cooldown_until = 0
    last_error = None
    last_status = 0
    cursor = 0
    names = list(roots)
    with ExitStack() as cleanup:
        cleanup.enter_context(RunLock(root / 'locks/convert-all'))
        for name, child in roots.items():
            cleanup.enter_context(RunLock(child / 'locks/convert'))
            database = Ledger(child, config['reserve_bytes'])
            cleanup.callback(database.db.close)
            database.db.execute("UPDATE objects SET conversion='pending' WHERE candidate=1 AND conversion='working'")
            databases[name] = database
        pool = cleanup.enter_context(ThreadPoolExecutor(max_workers=configured))
        while not (root / 'stop.request').exists() or active:
            paused = (root / 'conversion.pause.request').exists() or (root / 'stop.request').exists()
            available = psutil.virtual_memory().available
            private = owned_private_bytes()
            reserved = max(0, len(active) * budget - private)
            state = 'pausing' if paused and active else 'paused' if paused else 'running'
            while not paused and time.monotonic() >= cooldown_until and len(active) < effective:
                if available < reserve + reserved + budget:
                    state = 'waiting_for_memory'
                    break
                selected = None
                for offset in range(len(names)):
                    index = (cursor + offset) % len(names)
                    name = names[index]
                    child = roots[name]
                    if (child / 'pause.request').exists():
                        continue
                    database = databases[name]
                    try:
                        database.room((len(active) + 1) * 1024**3)
                    except StorageFull:
                        state = 'waiting_for_space'
                        continue
                    obj = database.claim('convert', exclude=tuple(job['sha'] for job in active.values()))
                    if obj is not None:
                        selected = (name, child, obj)
                        cursor = (index + 1) % len(names)
                        break
                if selected is None:
                    if state == 'running' and not active:
                        state = 'waiting_for_inputs'
                    break
                name, child, obj = selected
                future = pool.submit(run_job, child, obj)
                active[future] = {'corpus': name, 'sha': obj['sha'], 'source': obj['path']}
                reserved += budget
            done, _ = wait(active, timeout=.2, return_when=FIRST_COMPLETED)
            changed = bool(done)
            for future in done:
                job = active.pop(future)
                try:
                    future.result()
                except Exception as error:
                    database = databases[job['corpus']]
                    database.db.execute("UPDATE objects SET conversion='pending' WHERE sha=? AND conversion='working'", (job['sha'],))
                    effective = 1
                    cooldown_until = time.monotonic() + 60
                    last_error = type(error).__name__ + ': ' + str(error)
                    write_json(fallback_file, {'workers': 1, 'error': last_error, 'source_sha256': job['sha'], 'updated_utc': utc()})
            if changed or time.monotonic() - last_status >= 2:
                snapshot = {'state': state, 'configured_workers': configured, 'effective_workers': effective,
                            'active_jobs': list(active.values()), 'available_memory_bytes': available,
                            'owned_private_bytes': private, 'memory_reserve_bytes': reserve,
                            'job_memory_budget_bytes': budget, 'last_error': last_error, 'updated_utc': utc()}
                write_json(root / 'conversion-pool.json', snapshot)
                for name, child in roots.items():
                    jobs = [job for job in active.values() if job['corpus'] == name]
                    local_state = 'paused' if (child / 'pause.request').exists() else 'running' if jobs else state if state != 'running' else 'waiting_for_inputs'
                    write_json(child / 'conversion-status.json', {'state': local_state, 'active_jobs': jobs, 'updated_utc': utc()})
                last_status = time.monotonic()
            if not active:
                time.sleep(1)
