# LenDenClub lending decision desk

Static GitHub Pages dashboard. Net income always means interest minus scheduled fee minus realised default loss; returned principal is excluded from income.

## Refresh

```powershell
python -m pip install -r requirements.txt
python scripts/analyze.py C:\path\to\MANUAL_LENDING_REPORT.xlsx public\data.json
```

Push to `main`; GitHub Actions deploys `public/` to Pages. The workbook is never committed. Edit fee schedule only in `scripts/analyze.py` after verifying LenDenClub's terms.
