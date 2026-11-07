import argparse
import hashlib
import itertools
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
import uuid
import zipfile

from bootstrap import setup
from ledger import InspectionPaused,Ledger,StorageFull,digest,write_json,utc


def inspect_one(root,sha):
    from containers import INSPECTION_REVISION, inspect_object
    cfg=json.loads((root/'config.json').read_bytes());db=Ledger(root,cfg['reserve_bytes'])
    obj=dict(db.db.execute('SELECT * FROM objects WHERE sha=?',(sha,)).fetchone())
    before=db.db.execute('SELECT COUNT(*) FROM issues WHERE sha=?',(sha,)).fetchone()[0]
    try:
        inspect_object(db,obj,cfg)
        issues=db.db.execute('SELECT COUNT(*) FROM issues WHERE sha=?',(sha,)).fetchone()[0]
        db.db.execute('UPDATE objects SET state=?,inspection_revision=? WHERE sha=?',('review' if issues>before else 'complete',INSPECTION_REVISION,sha))
        db.db.commit()
    except InspectionPaused:
        db.db.execute("UPDATE objects SET state='pending' WHERE sha=?",(sha,))
        db.db.close()
        return 76
    except StorageFull:
        db.db.execute("UPDATE objects SET state='pending' WHERE sha=?",(sha,))
        db.db.commit()
        db.db.close()
        return 75
    except Exception as error:
        from containers import EXCEL_SUFFIXES
        # A damaged legacy workbook still gets an Excel attempt; failed container
        # inspection remains a separate open case even if Excel can convert it.
        if Path(obj['path']).suffix.lower() in EXCEL_SUFFIXES:db.workbook(sha)
        db.issue(sha,'container_inspection',type(error).__name__+': '+str(error))
        db.db.execute("UPDATE objects SET state='review',inspection_revision=? WHERE sha=?",(INSPECTION_REVISION,sha));db.db.commit()
    db.db.close()
    return 0


def inspect_loop(root,limit=0,lane=0,lanes=1):
    from containers import INSPECTION_REVISION,classify,inspect_object
    from inspection_process import InspectorProcess
    from ledger import inspection_lane_bounds
    from inspection_reuse import try_reuse
    from excel_shrink_research.enron.common import RunLock
    cfg=json.loads((root/'config.json').read_bytes());db=Ledger(root,cfg['reserve_bytes'])
    lower, upper = inspection_lane_bounds(lane, lanes)
    if lanes != cfg.get('inspection_workers', 1):
        db.db.close()
        raise ValueError('Inspection lanes must match the configured worker count')
    status_path = root / ('inspect-status.json' if lanes == 1 else 'inspect-status-' + str(lane) + '.json')
    cfg['inspection_status_path'] = str(status_path)
    lock_name = 'inspect' if lanes == 1 else 'inspect-' + str(lane)
    preferences=itertools.cycle(['enron-pst','fuse']);count=0;last_status=0
    with RunLock(root/'locks'/lock_name), InspectorProcess(root) as parser:
        db.db.execute("UPDATE objects SET state='pending' WHERE state='working' AND sha>=? AND sha<?", (lower, upper));db.db.commit()
        while not db.stop_file.exists():
            if (root / 'pause.request').exists():
                write_json(status_path, {'state': 'paused', 'updated_utc': utc()})
                if limit:
                    db.db.close()
                    return count
                time.sleep(10)
                continue
            try:db.room()
            except StorageFull:
                write_json(status_path,{'state':'waiting_for_space','updated_utc':utc()});time.sleep(30);continue
            obj=db.claim('inspect',next(preferences),lane=lane,lanes=lanes)
            if not obj:
                write_json(status_path,{'state':'waiting_for_inputs','updated_utc':utc()})
                if limit:
                    db.db.close()
                    return count
                time.sleep(3);continue
            sha=obj['sha']
            if time.monotonic() - last_status >= 1:
                write_json(status_path,{'state':'running','sha256':sha,'source':obj['path'],'updated_utc':utc()})
                last_status = time.monotonic()
            # Isolate complex Office/PDF parsers so one corrupt document cannot
            # stall the entire corpus. Simple archive streaming has no whole-job
            # deadline and never uses the system drive for intermediate files.
            try:
                try:
                    reused = try_reuse(db, obj)
                except (OSError, ValueError) as error:
                    db.issue(sha, 'inspection_reuse_unavailable', str(error))
                    reused = False
                if reused:
                    count += 1
                    if limit and count >= limit:
                        db.db.close()
                        return count
                    continue
                kind=classify(obj['path'])
                isolated = kind in ('ole','pdf','rtf','xml_container') or (kind == 'mime' and obj['bytes'] > 256 * 1024)
                if isolated:
                    write_json(status_path,{'state':'running','sha256':sha,'source':obj['path'],'updated_utc':utc()})
                    last_status = time.monotonic()
                    if cfg.get('reuse_inspection_process', False):
                        returncode = parser.inspect(sha)
                    else:
                        log=root/'logs/objects'/sha[:2]/(sha+'.log');log.parent.mkdir(parents=True,exist_ok=True)
                        with log.open('ab') as f:
                            child=subprocess.run([sys.executable,__file__,'one','--root',str(root),'--sha',sha],stdout=f,stderr=subprocess.STDOUT,timeout=300)
                        returncode = child.returncode
                    if returncode==76:
                        db.db.close()
                        return count
                    if returncode==75:time.sleep(30);continue
                    if returncode:raise ValueError('Object parser exited '+str(returncode))
                else:
                    before=db.db.execute('SELECT COUNT(*) FROM issues WHERE sha=?',(sha,)).fetchone()[0]
                    inspect_object(db,obj,cfg)
                    after=db.db.execute('SELECT COUNT(*) FROM issues WHERE sha=?',(sha,)).fetchone()[0]
                    db.db.execute('UPDATE objects SET state=?,inspection_revision=? WHERE sha=?',('review' if after>before else 'complete',INSPECTION_REVISION,sha));db.db.commit()
            except InspectionPaused:
                db.db.execute("UPDATE objects SET state='pending' WHERE sha=?",(sha,))
                db.db.close()
                return count
            except StorageFull:
                db.db.execute("UPDATE objects SET state='pending' WHERE sha=?",(sha,));db.db.commit();time.sleep(30);continue
            except Exception as error:
                db.issue(sha,'container_worker',type(error).__name__+': '+str(error))
                db.db.execute("UPDATE objects SET state='review',inspection_revision=? WHERE sha=?",(INSPECTION_REVISION,sha));db.db.commit()
            count+=1
            if limit and count>=limit:
                db.db.close()
                return count
    db.db.close()
    return count


def reuse(db,obj,target):
    from excel_shrink_research.enron.excel import _check_xlsx
    for previous in db.reuse_candidates(obj['sha']):
        source=Path(previous['bundle'])
        if not source.is_file():continue
        target.parent.mkdir(parents=True,exist_ok=True);pending=target.with_name(target.name+'.pending')
        try:
            if source.suffix.lower()=='.xlsx':shutil.copyfile(source,pending)
            else:
                with zipfile.ZipFile(source) as z, z.open('original.xlsx') as a, pending.open('wb') as b:shutil.copyfileobj(a,b,1024*1024)
            if digest(pending)!=previous['baseline_sha']:raise ValueError('Reusable XLSX checksum mismatch')
            error=_check_xlsx(pending)
            if error:raise ValueError(error)
            os.replace(pending,target)
            # Preserve already computed profiles with their original provenance.
            if source.suffix.lower()!='.xlsx':
                with zipfile.ZipFile(source) as z:
                    for name in ['default.xlsx','maximum-cleaning.xlsx','too-aggressive-cleaning.xlsx']:
                        if name in z.namelist():
                            dst=db.root/'reused-profiles'/obj['sha'][:2]/obj['sha']/name;dst.parent.mkdir(parents=True,exist_ok=True)
                            with z.open(name) as a,dst.open('wb') as b:shutil.copyfileobj(a,b,1024*1024)
            return {'status':'accepted','method':'reused_exact_source_hash','reuse_source':str(source),
                    'source_sha256':obj['sha'],'output_sha256':previous['baseline_sha'],'output':str(target),
                    'prompt_check':'inherited_preparation_evidence_not_repeated','updated_utc':utc()}
        except Exception as error:
            db.issue(obj['sha'],'reuse_unavailable',str(source)+': '+str(error));pending.unlink(missing_ok=True)
    return None


def conversion_suffix(obj):
    """Give Excel a staging extension matching supported content signatures."""
    with open(obj['path'],'rb') as f:prefix=f.read(4096)
    if prefix.startswith(bytes.fromhex('d0cf11e0a1b11ae1')):
        return '.xls'
    if prefix.startswith(b'PK'):
        try:
            with zipfile.ZipFile(obj['path']) as z:
                if 'xl/workbook.bin' in z.namelist():return '.xlsb'
                if 'xl/workbook.xml' in z.namelist():return '.xlsx'
                if 'mimetype' in z.namelist():return '.ods'
        except zipfile.BadZipFile:pass
    text = prefix.decode('utf-16', errors='replace') if prefix.startswith((b'\xff\xfe', b'\xfe\xff')) else prefix.decode('utf8', errors='replace')
    text = text.lstrip('\ufeff\r\n ')
    if text.startswith('ID;'):
        return '.slk'
    if '<html' in text[:2000].lower() or '<table' in text[:2000].lower():
        return '.html'
    if 'urn:schemas-microsoft-com:office:spreadsheet' in text:
        return '.xml'
    if '\t' in text.split('\n', 1)[0]:
        return '.tsv'
    if obj.get('kind') == 'legacy_excel' and any(separator in text.split('\n', 1)[0] for separator in (',', ';')):
        return '.csv'
    suffix=Path(obj['path']).suffix.lower()
    return suffix if suffix in ('.csv','.tsv','.slk','.ods','.xml','.html','.htm') else '.xls'


def convert_claimed(db, obj, cfg):
    """Convert one already claimed workbook and atomically publish its outcome.

    Each pool thread owns its Ledger connection. No database transaction spans
    file copying, Excel, or the independent reopen check. Infrastructure errors
    return the source to pending before they propagate to the scheduler.
    """
    from excel_shrink_research.enron.excel import convert_file, ExcelInfrastructureError
    root = db.root
    sha = obj['sha']
    target = root / 'xlsx' / sha[:2] / (sha + '.xlsx')
    attempt = root / 'work' / sha[:16] / uuid.uuid4().hex[:8]
    try:
        if digest(obj['path']) != sha:
            raise ValueError('Conversion source changed')
        result = reuse(db, obj, target)
        if result is None:
            attempt.mkdir(parents=True, exist_ok=True)
            local = attempt / ('source' + conversion_suffix(obj))
            shutil.copyfile(obj['path'], local)
            try:
                result = convert_file(local, target, attempt / 'excel', 90, 360,
                                      reconstruction_destination=root / 'reconstructed' / sha[:2] / (sha + '.xlsx'))
            finally:
                local.unlink(missing_ok=True)
            result['updated_utc'] = utc()
        state, output, output_sha, issue = 'review', None, None, None
        if result.get('status') == 'accepted':
            if result['source_sha256'] != sha:
                raise ValueError('Conversion source hash differs')
            state = 'reused' if result['method'] == 'reused_exact_source_hash' else 'accepted'
            output, output_sha = str(target), result['output_sha256']
        elif result.get('status') == 'reconstructed':
            if result['source_sha256'] != sha:
                raise ValueError('Reconstruction source hash differs')
            issue = ('reconstruction_review', 'Cell-copy reconstruction saved for review: ' + result['output'])
            output, output_sha = result['output'], result['output_sha256']
        else:
            issue = ('conversion', result.get('reason', 'review') + ': ' + result.get('error', ''))
        write_json(root / 'records' / sha[:2] / (sha + '.json'), dict(result, original_source=obj['path']))
        with db.transaction():
            if issue:
                db.issue(sha, issue[0], issue[1])
            db.db.execute('UPDATE objects SET conversion=?,output=?,output_sha=? WHERE sha=?',
                          (state, output, output_sha, sha))
        return state
    except ExcelInfrastructureError:
        db.db.execute("UPDATE objects SET conversion='pending' WHERE sha=?", (sha,))
        raise
    except Exception as error:
        with db.transaction():
            db.issue(sha, 'conversion_worker', type(error).__name__ + ': ' + str(error))
            db.db.execute("UPDATE objects SET conversion='review' WHERE sha=?", (sha,))
        return 'review'


def convert_loop(root, limit=0):
    """Keep the standalone serial entry point and its per-corpus worker lock."""
    from excel_shrink_research.enron.excel import ExcelInfrastructureError
    from excel_shrink_research.enron.common import RunLock
    cfg = json.loads((root / 'config.json').read_bytes())
    db = Ledger(root, cfg['reserve_bytes'])
    count = 0
    try:
        with RunLock(root / 'locks/convert'):
            db.db.execute("UPDATE objects SET conversion='pending' WHERE candidate=1 AND conversion='working'")
            while not db.stop_file.exists():
                if (root / 'pause.request').exists():
                    write_json(root / 'conversion-status.json', {'state': 'paused', 'updated_utc': utc()})
                    if limit:
                        return count
                    time.sleep(10)
                    continue
                try:
                    db.room(1024**3)
                except StorageFull:
                    write_json(root / 'conversion-status.json', {'state': 'waiting_for_space', 'updated_utc': utc()})
                    time.sleep(30)
                    continue
                obj = db.claim('convert')
                if obj is None:
                    write_json(root / 'conversion-status.json', {'state': 'waiting_for_inputs', 'updated_utc': utc()})
                    if limit:
                        return count
                    time.sleep(3)
                    continue
                write_json(root / 'conversion-status.json', {'state': 'running', 'sha256': obj['sha'],
                                                            'source': obj['path'], 'updated_utc': utc()})
                try:
                    convert_claimed(db, obj, cfg)
                    count += 1
                except ExcelInfrastructureError as error:
                    write_json(root / 'conversion-status.json', {'state': 'waiting_for_excel', 'error': str(error), 'updated_utc': utc()})
                    time.sleep(60)
                if limit and count >= limit:
                    return count
    finally:
        db.db.close()
    return count


def maximum_loop(root,limit=0):
    from excel_shrink_research.fuse_workflow import MAXIMUM,sheet_reductions
    from excel_shrink_research.enron.common import RunLock
    cfg=json.loads((root/'config.json').read_bytes());db=Ledger(root,cfg['reserve_bytes']);count=0
    with RunLock(root/'locks/maximum'):
        db.db.execute("UPDATE objects SET maximum='pending' WHERE candidate=1 AND maximum='working'");db.db.commit()
        while not db.stop_file.exists():
            if (root / 'pause.request').exists():
                if limit:
                    db.db.close()
                    return
                time.sleep(10)
                continue
            batch=db.db.execute("""SELECT o.* FROM objects o
                WHERE o.candidate=1 AND o.conversion IN ('accepted','reused') AND o.maximum='pending'
                AND EXISTS(SELECT 1 FROM memberships m WHERE m.sha=o.sha AND m.corpus='fuse')
                LIMIT 128""").fetchall()
            if not batch:
                if limit:
                    db.db.close()
                    return
                time.sleep(10);continue
            db.room(1024**3);run=root/'maximum-batches'/uuid.uuid4().hex[:12];inputs=run/'input';inputs.mkdir(parents=True)
            for obj in batch:
                dst=inputs/(obj['sha']+'.xlsx')
                try:os.link(obj['output'],dst)
                except OSError:shutil.copyfile(obj['output'],dst)
                db.db.execute("UPDATE objects SET maximum='working' WHERE sha=?",(obj['sha'],))
            db.db.commit()
            try:
                with (run/'run.log').open('wb') as f:
                    subprocess.run([cfg['shrink_executable'],str(inputs),str(run/'output'),
                                    '--batch-report',str(run/'report.json'),'--max-active-workbooks','2',
                                    '--writer-workers','1','--rayon-threads','4']+MAXIMUM,
                                   stdout=f,stderr=subprocess.STDOUT,timeout=3600)
            except subprocess.TimeoutExpired:
                db.issue('batch:'+run.name,'maximum_cleaning','Batch timeout; completed outputs are retained: '+str(run))
            for obj in batch:
                sha=obj['sha'];output=run/'output'/(sha+'.xlsx')
                if not output.is_file():
                    db.issue(sha,'maximum_cleaning','Batch failed; see '+str(run/'run.log'))
                    db.db.execute("UPDATE objects SET maximum='review' WHERE sha=?",(sha,));continue
                dest=root/'maximum'/sha[:2]/(sha+'.xlsx');dest.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(output,dest)
                stats=sheet_reductions(Path(obj['output']),dest,.001)
                write_json(root/'reductions'/sha[:2]/(sha+'.json'),{'source_sha256':sha,'output':str(dest),'sheets':stats})
                for sheet in stats:
                    if sheet['remaining_xml_pct']<=.001 and sheet.get('meaningful_cells',0):
                        write_json(root/'needles'/(sha+'-'+hashlib.sha256(sheet['part'].encode()).hexdigest()[:8]+'.json'),sheet)
                db.db.execute("UPDATE objects SET maximum='complete' WHERE sha=?",(sha,));count+=1
            db.db.commit()
            if limit and count>=limit:
                db.db.close()
                return
    db.db.close()


def main():
    p = argparse.ArgumentParser()
    p.add_argument('role', choices=['inspect','convert','maximum','one','convert-all','inspect-rest'])
    p.add_argument('--root', type=Path, required=True)
    p.add_argument('--sha')
    p.add_argument('--limit', type=int, default=0)
    p.add_argument('--lane', type=int, default=0)
    p.add_argument('--lanes', type=int, default=1)
    a = p.parse_args()
    root = a.root.resolve()
    setup(root)
    sys.path.insert(0, str(Path(__file__).parent))
    if a.role=='one':return inspect_one(root,a.sha)
    if a.role in ('convert-all', 'inspect-rest'):
        from partitioned import run_shared_worker
        return run_shared_worker(root, a.role)
    if a.role == 'inspect':
        inspect_loop(root, a.limit, a.lane, a.lanes)
    else:
        {'convert':convert_loop,'maximum':maximum_loop}[a.role](root,a.limit)
    return 0


if __name__=='__main__':sys.exit(main())
