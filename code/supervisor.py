"""Run extraction, exact-source reuse, conversion and FUSE cleaning independently."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

from bootstrap import setup
from ledger import Ledger,digest,write_json,utc


class StatusSampler:
    """Keep the worker heartbeat frequent without rescanning the ledger."""
    def __init__(self, interval=120, initial=None):
        self.interval = interval
        self.summary = dict(initial) if initial else None
        self.last = time.monotonic() if initial else None
        if self.summary:
            self.summary['updated_utc'] = initial.get('statistics_updated_utc', initial['updated_utc'])

    def collect(self, db, now):
        if self.last is None or now-self.last>=self.interval:
            self.summary=db.status();self.last=now
        status=dict(self.summary)
        status['statistics_updated_utc']=self.summary['updated_utc']
        status['updated_utc']=utc()
        return status


def queue_finished(db, reader_finished, required, feeds):
    # Before the source readers finish, completion is impossible. Do not scan
    # hundreds of thousands of objects merely to establish that every 10s.
    if not reader_finished or not all(feeds.get(n)==1 for n in required):return False
    unfinished=db.db.execute("SELECT 1 FROM objects WHERE state IN ('pending','working') OR (candidate=1 AND conversion IN ('pending','working')) LIMIT 1").fetchone()
    if unfinished:return False
    maximum=db.db.execute("""SELECT 1 FROM objects o WHERE candidate=1
        AND conversion IN ('accepted','reused') AND maximum IN ('pending','working')
        AND EXISTS(SELECT 1 FROM memberships m WHERE m.sha=o.sha AND m.corpus='fuse') LIMIT 1""").fetchone()
    return maximum is None


def run(root, adoption=None):
    setup(root)
    from excel_shrink_research.enron.common import RunLock
    cfg=json.loads((root/'config.json').read_bytes())
    if cfg.get('corpus_roots'):
        from partitioned import run_supervisor
        return run_supervisor(root, cfg)
    receipt=json.loads((root/'code-files.json').read_bytes())
    for file in receipt:
        if digest(root/'code'/file['path'])!=file['sha256']:raise ValueError('Frozen pipeline code changed: '+file['path'])
    db=Ledger(root,cfg['reserve_bytes']);children={};last_export=time.monotonic();reader_reported=0
    db.db.execute('PRAGMA cache_size=-65536')
    previous_status = root / 'status.json'
    initial = json.loads(previous_status.read_bytes()) if previous_status.is_file() else None
    sampler=StatusSampler(cfg.get('statistics_interval_seconds',120), initial)
    def start(name,command):
        logs=root/'logs';logs.mkdir(exist_ok=True)
        stdout=(logs/(name+'.stdout.log')).open('ab');stderr=(logs/(name+'.stderr.log')).open('ab')
        try:
            process=subprocess.Popen(command,stdout=stdout,stderr=stderr,cwd=root,
                                     creationflags=0x08000000 if os.name=='nt' else 0)
        finally:stdout.close();stderr.close()
        children[name]=(process,time.monotonic())
        write_json(root/(name+'-launch.json'),{'pid':process.pid,'command':command,'started_utc':utc()})
    commands={name:[sys.executable,str(root/'code/workers.py'),name,'--root',str(root)] for name in ['inspect','convert','maximum']}
    commands['seed']=[sys.executable,str(root/'code/seeding.py'),'--root',str(root)]
    commands['downloads']=[sys.executable,str(root/'code/fetch_sources.py'),'--root',str(root)]
    node_command=[cfg['node'],'--max-old-space-size=2048',str(root/'code/read_pst_attachments.cjs'),str(root/'pst-job.json')]
    with RunLock(root/'locks/supervisor'):
        if adoption:
            from processes import ExistingProcess
            snapshot=json.loads(Path(adoption).read_bytes())
            for name,identity in snapshot['workers'].items():
                process=ExistingProcess(identity['pid'],identity['created_filetime'])
                children[name]=(process,time.monotonic()-61)
            write_json(root/'supervisor-adoption.json',{'workers':snapshot['workers'],'adopted_utc':utc()})
        while not (root/'stop.request').exists():
            enough_space=shutil.disk_usage(root).free>=cfg['reserve_bytes']+1024**3
            reader=root/'pst-all/reader-status.json'
            reader_finished=reader.is_file() and json.loads(reader.read_bytes()).get('state') in ('complete','complete_with_issues')
            if not enough_space:
                current=children.get('pst-reader')
                if current and current[0].poll() is None:
                    current[0].terminate();current[0].wait(timeout=10)
            for name,command in dict(commands,**({} if reader_finished or not enough_space else {'pst-reader':node_command})).items():
                if name=='downloads' and (root/'download-status.json').is_file():
                    if json.loads((root/'download-status.json').read_bytes()).get('state')=='complete':continue
                previous=children.get(name)
                if previous is None or (previous[0].poll() is not None and time.monotonic()-previous[1]>60):
                    start(name,command)
            issues=root/'pst-all/reader-issues.jsonl'
            if issues.is_file() and issues.stat().st_size!=reader_reported:
                with issues.open(encoding='utf8') as f:
                    for line in f:
                        if not line.endswith('\n'):break
                        r=json.loads(line);db.issue(str(r.get('pst','pst-reader')),'pst_reader',json.dumps(r,sort_keys=True))
                reader_reported=issues.stat().st_size;db.db.commit()
            status=sampler.collect(db,time.monotonic());status['state']='running' if enough_space else 'waiting_for_space'
            status['workers']={name:{'pid':p.pid,'running':p.poll() is None,'exit_code':p.poll()} for name,(p,_) in children.items()}
            status['pst_full_retention_complete']=reader_finished
            status['fuse_profile']='maximum-cleaning'
            status['temporary_directory']=str(root/'temp')
            if 'needle_candidates' not in sampler.summary or time.monotonic()-sampler.last<1:
                sampler.summary['needle_candidates']=sum(1 for _ in (root/'needles').glob('*.json')) if (root/'needles').exists() else 0
            status['needle_candidates']=sampler.summary['needle_candidates']
            write_json(root/'status.json',status)
            if time.monotonic()-last_export>1200:
                db.export();last_export=time.monotonic()
            feeds={r['name']:r['complete'] for r in db.db.execute('SELECT * FROM feeds')}
            required={s['name'] for s in cfg['sources']}|{'previous-pst-journal','full-retention-pst-journal','fuse-binaries'}
            if queue_finished(db,reader_finished,required,feeds):
                db.export();status.update(db.status());status['statistics_updated_utc']=status['updated_utc']
                status['state']='complete_with_open_cases' if status['open_issues'] else 'complete'
                write_json(root/'status.json',status)
                (root/'stop.request').write_text('All queued stages finished; inspect open cases before claiming corpus completeness.\n')
                break
            time.sleep(10)
        for name,(process,_) in children.items():
            if process.poll() is None and name=='pst-reader':process.terminate()
        # Conversion/inspection workers finish their currently bounded file and
        # observe stop.request themselves. Their owned Excel jobs clean up safely.


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--adopt-workers',type=Path);a=p.parse_args()
    try:run(a.root.resolve(),a.adopt_workers)
    except Exception as error:
        write_json(a.root/'supervisor-error.json',{'state':'failed','error':type(error).__name__+': '+str(error),'updated_utc':utc()});raise
