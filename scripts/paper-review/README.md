# Paper scope inventory

`inventory.py` inventories the current production tree, the capture experiment
branch, every other origin head not contained in production, and tracked paper
scripts. It reads Git objects rather than temporary local files.

```powershell
python scripts/paper-review/inventory.py --repository F:/repos/excel-shrink-rust --output NEW_DIRECTORY
```

The output records revisions, all tracked paths, source hashes, role/family
classification and a lexical symbol index. It supplies an editorial coverage
ledger, not a compiler or correctness audit. No build or native Excel operation
runs. The committed family decisions are a separate human-reviewed assessment.
