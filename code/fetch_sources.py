"""Restore missing configured corpora directly to the F: workspace."""
import argparse
import hashlib
import json
from pathlib import Path
import time
from urllib.request import Request,urlopen
import zipfile

from bootstrap import setup
from ledger import write_json,utc


def get_json(url):
    with urlopen(Request(url,headers={'User-Agent':'corpus-preparation-research'}),timeout=60) as r:return json.load(r)


def run(root):
    cfg=json.loads((root/'config.json').read_bytes())
    while not (root/'stop.request').exists():
        unfinished=0
        for item in cfg.get('downloads',[]):
            target=Path(item['path']);receipt=target.with_suffix('.source.json')
            if target.is_file() and receipt.is_file():continue
            unfinished+=1
            try:
                target.parent.mkdir(parents=True,exist_ok=True)
                # The corpus may be reconnected before downloading starts.
                if item.get('local_fallback') and Path(item['local_fallback']).is_file():
                    import shutil
                    shutil.copyfile(item['local_fallback'],target)
                    write_json(receipt,dict(item,method='local_copy',finished_utc=utc()));continue
                source=target.with_suffix('.pinned.json')
                if source.exists():pin=json.loads(source.read_bytes())
                else:
                    repo=get_json('https://api.github.com/repos/'+item['repository'])
                    commit=get_json('https://api.github.com/repos/'+item['repository']+'/commits/'+repo['default_branch'])['sha']
                    pin={'repository':item['repository'],'commit':commit,'url':'https://codeload.github.com/'+item['repository']+'/zip/'+commit,'pinned_utc':utc()};write_json(source,pin)
                pending=target.with_suffix('.part');offset=pending.stat().st_size if pending.exists() else 0
                headers={'User-Agent':'corpus-preparation-research'}
                if offset:headers['Range']='bytes='+str(offset)+'-'
                with urlopen(Request(pin['url'],headers=headers),timeout=120) as response:
                    resume=response.status==206 and offset>0
                    if response.status not in (200,206):raise ValueError('Unexpected HTTP status')
                    if response.status==206 and not response.headers.get('Content-Range','').startswith('bytes '+str(offset)+'-'):
                        raise ValueError('Download resume range does not match the saved prefix')
                    received=0;declared=response.headers.get('Content-Length')
                    total=offset if resume else 0;last=0
                    with pending.open('ab' if resume else 'wb') as f:
                        while True:
                            if (root/'stop.request').exists():return
                            block=response.read(1024*1024)
                            if not block:break
                            import shutil
                            if shutil.disk_usage(root).free<cfg['reserve_bytes']+1024**3:raise RuntimeError('Data drive reserve reached')
                            f.write(block);total+=len(block)
                            received+=len(block)
                            if time.monotonic()-last>10:
                                write_json(root/'download-status.json',{'state':'downloading','corpus':item['corpus'],'bytes':total,'updated_utc':utc()});last=time.monotonic()
                    if declared is not None and received!=int(declared):raise ValueError('HTTP download ended before its declared length')
                with zipfile.ZipFile(pending) as z:members=len(z.infolist())
                h=hashlib.sha256()
                with pending.open('rb') as f:
                    for b in iter(lambda:f.read(1024*1024),b''):h.update(b)
                pending.replace(target);write_json(receipt,dict(pin,bytes=target.stat().st_size,sha256=h.hexdigest(),members=members,finished_utc=utc()))
            except Exception as error:
                write_json(root/'download-status.json',{'state':'retry_pending','corpus':item['corpus'],'error':str(error),'updated_utc':utc()})
        if not unfinished:
            write_json(root/'download-status.json',{'state':'complete','updated_utc':utc()});return
        time.sleep(120)


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);a=p.parse_args();setup(a.root);run(a.root.resolve())
