"""Read container contents without starting their associated applications."""
import bz2
from contextlib import contextmanager, ExitStack
from email import policy
from email.parser import BytesParser
import gzip
import hashlib
import io
import json
import lzma
import os
from pathlib import Path
import re
import stat
import subprocess
import tarfile
import time
import zipfile
import csv

from ledger import InspectionPaused, digest, safe_name, utc, write_json

OLE = bytes.fromhex('d0cf11e0a1b11ae1')
EXCEL_SUFFIXES = {'.xls','.xlsx','.xlsm','.xlsb','.xlt','.xltx','.xltm','.xlw','.xlc','.xla','.xlam','.ods','.csv','.tsv','.slk','.sylk','.wk1','.wk3','.wk4','.wks'}
INSPECTION_REVISION = 1


def classify(path):
    path = Path(path)
    with path.open('rb') as f:
        prefix = f.read(65536)
    return classify_prefix(prefix, path.suffix.lower())


def tabular_prefix(prefix):
    """Recognize text spreadsheet formats without relying on their filename."""
    if prefix.startswith((b'\xff\xfe', b'\xfe\xff')):
        text = prefix.decode('utf-16', errors='replace')
    elif b'\0' not in prefix:
        text = prefix.decode('utf-8-sig', errors='replace')
    else:
        return False
    text = text.lstrip('\ufeff\r\n ')
    if text.startswith('ID;') or re.search(r'<(?:[A-Za-z0-9_]+:)?table\b', text, re.I):
        return True
    # EDRM TXT mail exports remain mail text, even when headers use tabs.
    if re.match(r'(?:Date|From|To|Subject|Received|Message-ID):', text, re.I):
        return False
    lines = [line for line in text.splitlines()[:12] if line.strip()]
    if len(lines) < 2:
        return False
    for separator in ('\t', ',', ';'):
        if separator not in lines[0]:
            continue
        try:
            rows = list(csv.reader(lines, delimiter=separator, strict=True))
        except csv.Error:
            continue
        width = len(rows[0])
        if width >= 2 and sum(len(row) == width for row in rows) >= max(2, len(rows) * .8):
            return True
    return False


def classify_prefix(prefix, suffix=''):
    """Classify content before extension hints, including misleading .bin inputs."""
    stripped = prefix.lstrip(b'\xef\xbb\xbf\x00\t\r\n ')
    if prefix.startswith(OLE): return 'ole'
    if prefix.startswith((b'PK\x03\x04', b'PK\x05\x06', b'PK\x07\x08')) or suffix in {'.zip','.docx','.pptx','.xlsx','.xlsm','.xlsb','.odt','.ods','.odp','.epub'}: return 'zip'
    if prefix.startswith(b'\x1f\x8b'): return 'gzip'
    if prefix.startswith(b'BZh'): return 'bzip2'
    if prefix.startswith(b'\xfd7zXZ\x00'): return 'xz'
    if prefix[257:262] == b'ustar' or suffix == '.tar': return 'tar'
    if stripped.startswith(b'%PDF-'): return 'pdf'
    if stripped.startswith(b'{\\rtf'): return 'rtf'
    if prefix.startswith((b'7z\xbc\xaf\x27\x1c',b'Rar!',b'MSCF',b'MZ')) or suffix in {'.7z','.rar','.cab','.iso','.arj','.lzh','.ace'}: return 'seven'
    if suffix in {'.eml','.mht','.mhtml'} or b'mime-version:' in prefix[:8192].lower(): return 'mime'
    if b'urn:schemas-microsoft-com:office:word' in prefix or b'wordprocessingml' in prefix[:4096]: return 'xml_container'
    if b'urn:schemas-microsoft-com:office:spreadsheet' in prefix: return 'legacy_excel'
    if prefix[:2] in [bytes.fromhex(s) for s in ['0900','0902','0904','0908']] and len(prefix)>=8: return 'legacy_excel'
    if b'urn:schemas-microsoft-com:office:excel' in prefix or b'application/vnd.ms-excel' in prefix: return 'legacy_excel'
    if suffix in EXCEL_SUFFIXES: return 'legacy_excel'
    if tabular_prefix(prefix): return 'legacy_excel'
    return 'other'


def inspect_object(db, obj, config):
    sha, path, depth = obj['sha'], Path(obj['path']), obj['depth'] + 1
    if depth > config.get('maximum_container_depth',128):
        raise ValueError('Container nesting exceeds the configured limit; case remains open')
    if not obj['verified']:
        if digest(path) != sha: raise ValueError('Saved source checksum changed')
        db.db.execute('UPDATE objects SET verified=1 WHERE sha=?',(sha,))
        db.db.commit()
    kind = classify(path)
    db.db.execute('UPDATE objects SET kind=? WHERE sha=?',(kind,sha));db.db.commit()
    with archive_io_slot(db, obj, config, kind):
        if kind=='legacy_excel': db.workbook(sha)
        elif kind=='zip': inspect_zip(db,obj,config,depth)
        elif kind=='tar': inspect_tar(db,obj,config,depth)
        elif kind in ('gzip','bzip2','xz'):
            opener={'gzip':gzip.open,'bzip2':bz2.open,'xz':lzma.open}[kind]
            name=path.name.rsplit('.',1)[0]
            if name==path.name:name='decompressed.bin'
            with opener(path,'rb') as stream:db.store(sha,0,name,stream,kind,depth)
        elif kind=='seven': inspect_seven(db,obj,config,depth)
        elif kind=='ole': inspect_ole(db,obj,config,depth)
        elif kind=='rtf': inspect_rtf(db,obj,config,depth)
        elif kind=='pdf': inspect_pdf(db,obj,config,depth)
        elif kind=='mime': inspect_mime(db,obj,config,depth)
        elif kind=='xml_container': inspect_xml(db,obj,config,depth)
    return kind


@contextmanager
def archive_io_slot(db, obj, config, kind):
    """Serialize large archive streams on the HDD across inspection lanes.

    Small containers and document parsing can overlap. Waiting for the large
    stream slot responds to pause requests and retains the current claim.
    The OS releases the lock if an inspector exits unexpectedly.
    """
    lock_path = config.get('archive_io_lock')
    if not lock_path or kind not in ('zip', 'tar', 'gzip', 'bzip2', 'xz', 'seven') or obj['bytes'] < 16 * 1024**2:
        yield
        return
    from excel_shrink_research.enron.common import RunLock
    with ExitStack() as held:
        last_status = 0
        while True:
            check_archive_pause(db)
            try:
                held.enter_context(RunLock(Path(lock_path)))
                break
            except RuntimeError:
                if config.get('inspection_status_path') and time.monotonic() - last_status >= 2:
                    write_json(config['inspection_status_path'], {'state': 'waiting_for_archive_io',
                               'sha256': obj['sha'], 'source': obj['path'], 'updated_utc': utc()})
                    last_status = time.monotonic()
                time.sleep(.5)
        yield


def check_archive_pause(db):
    """Stop after the previous member was committed, preserving retry evidence."""
    if db.stop_file.exists() or (db.root / 'pause.request').exists():
        raise InspectionPaused('Archive inspection paused between members')


def ordinary_archive_mail_text(name, prefix):
    """Identify ordinary EDRM mail exports after all supported format signatures.

    Restrict retention to the source's text directories and mail header layout.
    Tabular content and nested container signatures still enter normal extraction.
    """
    if not re.search(r'(^|/)text_\d+/.*\.txt$', name.replace('\\', '/'), re.I):
        return False
    if classify_prefix(prefix, '.txt') != 'other':
        return False
    encoding = 'utf-16' if prefix.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
    text = prefix.decode(encoding, errors='replace')
    return bool(re.match(r'(?:Date|From|Message-ID):', text, re.I)) and len(
        re.findall(r'^(?:Date|From|To|Subject|Message-ID):', text[:8192], re.I | re.M)) >= 2


def retain_archive_mail_text(db, obj, archive, info, ordinal):
    """Keep a proven ordinary mail body in its ZIP with a checked member receipt.

    Reading to EOF validates the ZIP CRC. The source archive retains every byte.
    Previously materialized objects and their provenance remain available.
    """
    saved = db.db.execute('SELECT member,bytes,inspection_revision FROM archive_text WHERE parent=? AND ordinal=?',
                          (obj['sha'], ordinal)).fetchone()
    if saved and saved['member'] == info.filename and saved['bytes'] == info.file_size and saved['inspection_revision'] == INSPECTION_REVISION:
        return True
    if not re.search(r'(^|/)text_\d+/.*\.txt$', info.filename.replace('\\', '/'), re.I):
        return False
    safe_name(info.filename)
    with archive.open(info) as stream:
        prefix = stream.read(65536)
        if not ordinary_archive_mail_text(info.filename, prefix):
            return False
        content_hash = hashlib.sha256(prefix)
        size = len(prefix)
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            content_hash.update(block)
            size += len(block)
    if size != info.file_size:
        raise ValueError('Truncated retained mail text')
    sha = content_hash.hexdigest()
    with db.transaction():
        db.db.execute('INSERT OR REPLACE INTO archive_text VALUES(?,?,?,?,?,?)',
                      (obj['sha'], ordinal, info.filename, size, sha, INSPECTION_REVISION))
        db.db.execute("UPDATE objects SET kind='other',state='complete',inspection_revision=? WHERE sha=? AND candidate=0 AND state='pending' AND verified=1",
                      (INSPECTION_REVISION, sha))
    return True


def inspect_zip(db,obj,config,depth):
    with zipfile.ZipFile(obj['path']) as archive:
        names=set(archive.namelist())
        workbook_package='xl/workbook.xml' in names or 'xl/workbook.bin' in names
        if workbook_package:
            db.workbook(obj['sha'])
        elif 'mimetype' in names and b'oasis.opendocument.spreadsheet' in archive.read('mimetype'):
            db.workbook(obj['sha'])
        last_status = 0
        entries = archive.infolist()
        for ordinal,info in enumerate(entries):
            check_archive_pause(db)
            if config.get('inspection_status_path') and time.monotonic() - last_status >= 2:
                write_json(config['inspection_status_path'], {'state': 'running', 'sha256': obj['sha'],
                           'source': obj['path'], 'archive_entry': ordinal, 'archive_entries': len(entries),
                           'updated_utc': utc()})
                last_status = time.monotonic()
            if info.is_dir():continue
            # Worksheet/style XML is the workbook's representation, not an
            # embedded file. Preserve it inside the XLSX; expand all embedded
            # binary payloads, including those with misleading suffixes.
            if workbook_package and info.filename.lower().endswith(('.xml','.rels')) and '/embeddings/' not in info.filename:
                continue
            try:
                if stat.S_IFMT(info.external_attr>>16)==stat.S_IFLNK:raise ValueError('Archive links are not materialized')
                if config.get('retain_archive_mail_text', False) and retain_archive_mail_text(db, obj, archive, info, ordinal):
                    continue
                with archive.open(info) as stream:
                    db.store(obj['sha'],ordinal,info.filename,stream,'zip',depth,info.file_size)
            except Exception as error:
                from ledger import StorageFull
                if isinstance(error,StorageFull):raise
                db.issue(obj['sha'],'zip_member',info.filename+': '+str(error))
        db.db.commit()


def inspect_tar(db,obj,config,depth):
    adopt=config.get('fuse_existing_output') if str(Path(obj['path']).resolve())==str(Path(config['fuse_tar']).resolve()) else None
    # Native TAR seeks directly to members; compressed TARs reach this path only
    # after their streaming decompression on F:.
    count=0
    with tarfile.open(obj['path'],'r:') as archive:
        for ordinal,info in enumerate(archive):
            check_archive_pause(db)
            if info.isdir():continue
            try:
                safe_name(info.name)
                if not info.isfile():raise ValueError('Archive link or special file is not materialized')
                count+=1
                existing=Path(adopt)/Path(*info.name.replace('\\','/').split('/')) if adopt else None
                if existing and existing.is_file() and existing.stat().st_size==info.size:
                    h=hashlib.sha256()
                    with archive.extractfile(info) as stream:
                        for block in iter(lambda:stream.read(1024*1024),b''):h.update(block)
                    expected=h.hexdigest()
                    if digest(existing)==expected:
                        db.add(existing,parent=obj['sha'],member=info.name+f' [entry {ordinal}]',method='verified_user_tar_extraction',expected=expected,verified=True,depth=depth)
                        if ordinal%50==0:db.db.commit()
                        continue
                with archive.extractfile(info) as stream:
                    db.store(obj['sha'],ordinal,info.name,stream,'tar',depth,info.size)
            except Exception as error:
                from ledger import StorageFull
                if isinstance(error,StorageFull):raise
                db.issue(obj['sha'],'tar_member',info.name+': '+str(error))
        db.db.commit()
    if adopt:
        from ledger import write_json,utc
        write_json(db.root/'fuse-tar-inventory.json',{'files':count,'expected_publication_files':249376,'finished_utc':utc()})
        if count!=249376:db.issue(obj['sha'],'fuse_source_count',f'TAR has {count} regular files; publication states 249376')


def inspect_seven(db,obj,config,depth):
    out=db.root/'seven'/obj['sha'][:24];out.mkdir(parents=True,exist_ok=True)
    listing=out/'listing.txt'
    command=[config['seven_zip'],'l','-slt','-ba','-p__NoPassword_Probe_7x9Q!',obj['path']]
    with listing.open('wb') as f:
        result=subprocess.run(command,stdout=f,stderr=subprocess.STDOUT,timeout=600)
    if result.returncode:raise ValueError('7-Zip could not list the container; see '+str(listing))
    for line in listing.read_text(encoding='utf-8-sig',errors='replace').splitlines():
        if line.startswith('Path = '):
            name=line[7:];safe_name(name)
            if ':' in name:raise ValueError('Archive path contains a Windows stream separator')
        if line.startswith(('Symbolic Link = ','Hard Link = ')):raise ValueError('Archive contains a link')
    target=out/'files';target.mkdir(exist_ok=True)
    with (out/'extract.log').open('wb') as f:
        result=subprocess.run([config['seven_zip'],'x',obj['path'],'-o'+str(target),'-y','-p__NoPassword_Probe_7x9Q!','-bb0','-bsp0'],stdout=f,stderr=subprocess.STDOUT,timeout=7200)
    if result.returncode:raise ValueError('7-Zip extraction failed or requires a password; see '+str(out/'extract.log'))
    for folder,dirs,files in os.walk(target):
        for name in files:
            path=Path(folder)/name
            if path.is_symlink():raise ValueError('Extractor produced a link')
            db.room();db.add(path,parent=obj['sha'],member=path.relative_to(target).as_posix(),method='7zip',depth=depth)
        db.db.commit()


def inspect_ole(db,obj,config,depth):
    import olefile
    from excel_shrink_research.enron import embedded_workbooks as e
    from excel_shrink_research.enron.container_embeddings import ppt_objects,native_payload
    sha=obj['sha'];ordinal=0
    with olefile.OleFileIO(obj['path']) as ole:
        root=e.StorageReader(ole,ole.root)
        stream=e.workbook_hash(root)
        if stream:db.workbook(sha,stream)
        elif ole.exists('EncryptedPackage'):
            db.workbook(sha);db.issue(sha,'encrypted_office','Encrypted Office container requires a password')
        elif Path(obj['path']).suffix.lower() in EXCEL_SUFFIXES:
            db.workbook(sha)
        def walk(storage,path=()):
            nonlocal ordinal
            for name,kind,size in e.entries(storage):
                childpath=path+(name,)
                try:
                    if kind==1:
                        child=storage.OpenStorage(name,None,0x20)
                        names={x[0] for x in e.entries(child)}
                        workbook=e.workbook_hash(child)
                        if workbook or names.intersection({'WordDocument','PowerPoint Document','EncryptedPackage','__properties_version1.0'}):
                            suffix='.xls' if workbook else '.msg' if '__properties_version1.0' in names else '.ole'
                            ordinal+=1;target=db.child_path(sha,ordinal,'/'.join(childpath)+suffix);target.parent.mkdir(parents=True,exist_ok=True)
                            e.extract(child,target)
                            childsha=db.add(target,parent=sha,member='/'.join(childpath),method='ole_storage',depth=depth)
                            if workbook:db.workbook(childsha,workbook)
                        # Continue into every storage, including Word ObjectPool
                        # and Outlook embedded-message attachment storages.
                        walk(child,childpath)
                    elif kind==2:
                        ordinal+=1
                        if name in ('Workbook','Book'):continue
                        stream=storage.OpenStream(name,None,0x20,0)
                        if name=='PowerPoint Document':
                            data=stream.read()
                            for offset,payload in ppt_objects(data):
                                ordinal+=1;db.store(sha,ordinal,'/'.join(childpath)+f'/record-{offset}.bin',io.BytesIO(payload),'powerpoint_object',depth)
                        elif name=='Package' or name.startswith('__substg1.0_37010102'):
                            db.store(sha,ordinal,'/'.join(childpath)+'.bin',stream,'ole_package',depth,size)
                        elif name=='\x01Ole10Native':
                            data=native_payload(stream.read());db.store(sha,ordinal,'/'.join(childpath)+'.bin',io.BytesIO(data),'ole_native',depth)
                        else:
                            prefix=stream.read(8)
                            if prefix.startswith((OLE,b'PK\x03\x04',b'7z\xbc\xaf',b'Rar!',b'%PDF-',b'{\\rtf')):
                                stream.seek(0);db.store(sha,ordinal,'/'.join(childpath)+'.bin',stream,'ole_embedded_stream',depth,size)
                except Exception as error:
                    from ledger import StorageFull
                    if isinstance(error,StorageFull):raise
                    db.issue(sha,'ole_object','/'.join(childpath)+': '+str(error))
        walk(root);db.db.commit()


def inspect_rtf(db,obj,config,depth):
    from oletools.rtfobj import RtfObjParser
    parser=RtfObjParser(Path(obj['path']).read_bytes());parser.parse()
    for i,item in enumerate(parser.objects):
        data=getattr(item,'olepkgdata',None) or getattr(item,'oledata',None) or getattr(item,'rawdata',None)
        if data:db.store(obj['sha'],i,'object-'+str(i)+'.bin',io.BytesIO(data),'rtf_object',depth)


def inspect_pdf(db,obj,config,depth):
    from pypdf import PdfReader
    reader=PdfReader(obj['path'],strict=True)
    if reader.is_encrypted and not reader.decrypt(''):raise ValueError('PDF requires a password')
    i=0
    for name,items in reader.attachments.items():
        for payload in items:
            db.store(obj['sha'],i,name,io.BytesIO(payload),'pdf_attachment',depth);i+=1


def inspect_mime(db,obj,config,depth):
    message=BytesParser(policy=policy.default).parsebytes(Path(obj['path']).read_bytes())
    for i,part in enumerate(message.walk()):
        if part.is_multipart():continue
        payload=part.get_payload(decode=True)
        if payload:
            if not part.get_filename() and part.get_content_disposition() != 'attachment' and part.get_content_type() in ('text/plain', 'text/html'):
                kind = classify_prefix(payload[:65536])
                if kind == 'other':
                    # The immutable EML retains the body. Record its identity
                    # without materializing another ordinary mail-text file.
                    db.db.execute('INSERT OR REPLACE INTO mime_bodies VALUES(?,?,?,?,?)',
                                  (obj['sha'], i, part.get_content_type(), len(payload), hashlib.sha256(payload).hexdigest()))
                    continue
            db.store(obj['sha'],i,part.get_filename() or f'part-{i}.bin',io.BytesIO(payload),'mime_part',depth)
    if message.defects:db.issue(obj['sha'],'mime_parse',str(message.defects))


def inspect_xml(db,obj,config,depth):
    import base64
    import xml.etree.ElementTree as ET
    i=0
    for event,node in ET.iterparse(obj['path'],events=('end',)):
        if node.tag.rsplit('}',1)[-1] in ('binData','binary-data') and node.text:
            payload=base64.b64decode(re.sub(r'\s+','',node.text),validate=True)
            name=next((v for k,v in node.attrib.items() if k.rsplit('}',1)[-1]=='name'),f'object-{i}.bin')
            db.store(obj['sha'],i,name,io.BytesIO(payload),'xml_binary',depth);i+=1
        node.clear()
