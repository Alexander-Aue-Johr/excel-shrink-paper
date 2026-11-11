"""Read uncompressed XLSX member totals from the recorded original/output pairs."""
import csv,json,pathlib,zipfile
import argparse
p=argparse.ArgumentParser(description=__doc__);p.add_argument('package',type=pathlib.Path);p.add_argument('output',type=pathlib.Path);a=p.parse_args()
with (a.package/'workbooks.csv').open(encoding='utf-8-sig',newline='') as f: rows=list(csv.DictReader(f))
a.output.parent.mkdir(parents=True,exist_ok=True)
fields=['input_file','source_xlsx_sha256','output_sha256','status','original_uncompressed_bytes','output_uncompressed_bytes','original_members','output_members','duplicate_names','error']
with a.output.open('w',encoding='utf-8-sig',newline='') as stream:
 w=csv.DictWriter(stream,fieldnames=fields);w.writeheader()
 for row in rows:
  result={key:row.get(key,'') for key in fields};result['error']=''
  if row['status']=='processed':
   try:
    for side,key,pathkey in [('original','original_uncompressed_bytes','input_file'),('output','output_uncompressed_bytes','output_file')]:
     with zipfile.ZipFile(row[pathkey]) as archive:
      members=archive.infolist();result[key]=sum(x.file_size for x in members);result[side+'_members']=len(members)
      result['duplicate_names']=int(result.get('duplicate_names') or 0)+len(members)-len({x.filename for x in members})
   except Exception as error:result['error']=str(error)
  w.writerow(result)
print(a.output,flush=True)
