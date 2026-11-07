import argparse
import csv
import hashlib
import json
import os
from pathlib import Path
import time

from bootstrap import setup
from ledger import Ledger,write_json,utc


def prepare_reuse(db,cfg):
    targets={}
    with open(cfg['old_missing'],newline='',encoding='utf-8-sig') as f:
        for r in csv.DictReader(f):
            targets[r['sha256']]=json.loads(r['member_chain'])[-1]
    assert len(targets)==146,'Unexpected missing-Enron search list'
    for sha,member in targets.items():db.db.execute('INSERT OR IGNORE INTO targets(sha,member) VALUES(?,?)',(sha,member))
    with open(cfg['old_stream_index'],encoding='utf8') as f:
        for line in f:
            r=json.loads(line)
            if r.get('file_sha256') in targets and r.get('workbook_sha256'):
                db.db.execute('UPDATE targets SET stream_sha=? WHERE sha=?',(r['workbook_sha256'],r['file_sha256']))
    db.db.commit()
    imported=0
    for directory in cfg['reuse_runs']:
        base=Path(directory)
        with os.scandir(base/'records') as entries:
            for entry in entries:
                if not entry.name.endswith('.json'):continue
                try:
                    r=json.loads(Path(entry.path).read_bytes())
                    sha=r.get('source_sha256');baseline=r.get('baseline_sha256') or r.get('converted_sha256')
                    bundle=base/'prepared'/(Path(entry.name).stem+'.zip')
                    if r.get('prepared') and sha and baseline and len(sha)==64 and bundle.is_file():
                        db.db.execute('INSERT OR IGNORE INTO reuse VALUES(?,?,?,?)',(sha,r.get('source_bytes',-1),str(bundle),baseline));imported+=1
                except (OSError,ValueError):continue
                if imported%500==0:
                    db.db.commit();write_json(db.root/'seed-status.json',{'stage':'indexing_previous_conversions','records':imported,'updated_utc':utc()})
    for manifest in cfg['conversion_manifests']:
        if not Path(manifest).is_file():continue
        with open(manifest,encoding='utf-8-sig') as f:
            for line in f:
                r=json.loads(line)
                if r.get('status')=='accepted' and r.get('output') and r.get('source_sha256') and Path(r['output']).is_file():
                    db.db.execute('INSERT OR IGNORE INTO reuse VALUES(?,?,?,?)',(r['source_sha256'],-1,r['output'],r['output_sha256']));imported+=1
    db.db.commit();write_json(db.root/'reuse-index.json',{'entries':imported,'updated_utc':utc()})


def directory_inputs(root):
    for folder,dirs,files in os.walk(root):
        dirs[:]=sorted(d for d in dirs if not (Path(folder)/d).is_symlink())
        for name in sorted(files):
            p=Path(folder)/name
            if not p.is_symlink():yield p


def seed_static(db,cfg):
    for source in cfg['sources']:
        if source.get('wait_for_fuse'):continue
        name=source['name'];done=db.db.execute('SELECT complete FROM feeds WHERE name=?',(name,)).fetchone()
        if done and done['complete']:continue
        root=Path(source['path'])
        if not root.exists():
            db.issue(name,'source_unavailable',str(root));db.db.commit();continue
        db.db.execute("UPDATE issues SET resolved=1 WHERE sha=? AND stage='source_unavailable'",(name,));db.db.commit()
        paths=directory_inputs(root) if root.is_dir() else [root]
        count=0
        for p in paths:
            if db.stop_file.exists():return
            if source.get('archives_only') and p.suffix.lower() not in {'.zip','.7z','.rar','.gz','.tar'}:continue
            db.room();db.add(p,corpus=source['corpus'],logical=str(p),metadata={'source_group':name})
            count+=1
            if count%25==0:
                db.db.commit();write_json(db.root/'seed-status.json',{'stage':'registering_sources','group':name,'registered':count,'updated_utc':utc()})
        db.db.execute('INSERT OR REPLACE INTO feeds(name,complete) VALUES(?,1)',(name,));db.db.commit()


def feed_journal(db,name,path,base,finished,committed_bytes=None):
    path=Path(path)
    if not path.exists():return
    record=db.db.execute('SELECT position FROM feeds WHERE name=?',(name,)).fetchone();position=record['position'] if record else 0
    if path.stat().st_size<position:raise ValueError('PST journal was truncated after it was ingested')
    count=0
    with path.open('rb') as f:
        f.seek(position)
        while True:
            if committed_bytes is not None and f.tell()>=committed_bytes:break
            before=f.tell();line=f.readline()
            if not line or not line.endswith(b'\n'):break
            if committed_bytes is not None and f.tell()>committed_bytes:break
            r=json.loads(line)
            if r.get('type')=='payload' and r.get('path'):
                p=(Path(base)/r['path']).resolve();p.relative_to(Path(base).resolve())
                if p.is_file():
                    logical=f"{r.get('pst')}/{r.get('message_node_id')}/{json.dumps(r.get('embedded_message_chain',[]),sort_keys=True)}/{r.get('attachment_index')}/{r.get('original_filename')}"
                    db.add(p,corpus='enron-pst',logical=logical,expected=r['sha256'],size=r['bytes'],metadata=r)
                else:db.issue(r.get('sha256','pst'),'retained_attachment_missing',str(p))
            position=f.tell();count+=1
            if count%100==0:
                db.db.execute('INSERT OR REPLACE INTO feeds VALUES(?,?,0)',(name,position));db.db.commit()
    db.db.execute('INSERT OR REPLACE INTO feeds VALUES(?,?,?)',(name,position,int(finished and position==path.stat().st_size)));db.db.commit()


def pid_alive(pid,creation=None):
    import ctypes
    from ctypes import wintypes
    k=ctypes.WinDLL('kernel32',use_last_error=True)
    k.OpenProcess.restype=wintypes.HANDLE
    k.OpenProcess.argtypes=[wintypes.DWORD,wintypes.BOOL,wintypes.DWORD]
    k.GetExitCodeProcess.argtypes=[wintypes.HANDLE,ctypes.POINTER(wintypes.DWORD)]
    k.GetProcessTimes.argtypes=[wintypes.HANDLE]+[ctypes.POINTER(wintypes.FILETIME)]*4
    k.CloseHandle.argtypes=[wintypes.HANDLE]
    handle=k.OpenProcess(0x1000,False,int(pid))
    if not handle:return ctypes.get_last_error()==5
    try:
        code=wintypes.DWORD();k.GetExitCodeProcess(handle,ctypes.byref(code))
        if code.value!=259:return False
        if creation:
            times=[wintypes.FILETIME() for _ in range(4)]
            if not k.GetProcessTimes(handle,*[ctypes.byref(t) for t in times]):return True
            actual=(times[0].dwHighDateTime<<32)|times[0].dwLowDateTime
            if str(actual)!=str(creation):return False
        return True
    finally:k.CloseHandle(handle)


def run(root,once=False):
    cfg=json.loads((root/'config.json').read_bytes());db=Ledger(root,cfg['reserve_bytes'])
    from excel_shrink_research.enron.common import RunLock
    with RunLock(root/'locks/seed'):
        if not cfg.get('reuse_index_migrated'):
            prepare_reuse(db,cfg)
        # Import retained PST payloads before the much larger full-retention pass.
        if cfg.get('corpus') in (None, 'enron-pst'):
            feed_journal(db,'previous-pst-journal',cfg['previous_pst_journal'],cfg['previous_pst_root'],True)
        seed_static(db,cfg)
        while not db.stop_file.exists():
            # Newly downloaded or reconnected corpora join without a restart.
            seed_static(db,cfg)
            reader_root = Path(cfg.get('pst_reader_root', root/'pst-all'))
            marker=reader_root/'reader-status.json'
            finished=marker.is_file() and json.loads(marker.read_bytes()).get('state') in ('complete','complete_with_issues')
            checkpoint=reader_root/'reader-checkpoints.json'
            committed=None if finished else json.loads(checkpoint.read_bytes())['journal_bytes'] if checkpoint.is_file() else 0
            # The reader may truncate an unfinished PST on restart. Only ingest
            # its committed journal prefix, so resume cannot orphan ledger rows.
            if cfg.get('corpus') in (None, 'enron-pst'):
                feed_journal(db,'full-retention-pst-journal',reader_root/'attachments.jsonl',reader_root,finished,committed)
            wait=cfg.get('fuse_wait',{})
            ready=not wait.get('pid') or not pid_alive(wait['pid'],wait.get('created_filetime'))
            done=db.db.execute("SELECT complete FROM feeds WHERE name='fuse-binaries'").fetchone()
            if cfg.get('corpus') in (None, 'fuse') and ready and not (done and done['complete']):
                tar=Path(cfg['fuse_tar'])
                if tar.is_file():
                    db.add(tar,corpus='fuse',logical=str(tar),metadata={'user_extraction_root':cfg['fuse_existing_output']})
                    db.db.execute("INSERT OR REPLACE INTO feeds(name,complete) VALUES('fuse-binaries',1)");db.db.commit()
                else:db.issue('fuse-binaries','source_unavailable',str(tar));db.db.commit()
            write_json(root/'seed-status.json',{'state':'watching_inputs','fuse_extractor_finished':ready,'full_pst_reader_finished':finished,'updated_utc':utc()})
            if once:
                db.db.close()
                return
            time.sleep(10)


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--once',action='store_true');p.add_argument('--all',action='store_true');a=p.parse_args();setup(a.root)
    if a.all:
        from partitioned import run_seeding
        run_seeding(a.root.resolve())
    else:
        run(a.root.resolve(),a.once)
