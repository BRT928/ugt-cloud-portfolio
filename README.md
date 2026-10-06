# UGT Cloud Portfolio

Static product portfolio site generated from `UGT_Product_Portfolio_2026.xlsx` (OneDrive).

## Live site

https://brt928.github.io/ugt-cloud-portfolio/

## Pages

- Overview — `/`
- Products — `/products.html` (filters via query, e.g. `?due=2026+Q4`)
- Platforms — `/platforms.html`
- Roadmap — `/roadmap.html`
- Owners — `/owners.html`, `/owner-TGA.html`, …
- TGA & NNI / Status notes — `/status.html`
- Legend — `/glossary.html`
- Product detail — `/product-1.html`, …

## Rebuild

```bash
python3 build.py
```

Requires the Excel file at the workspace path (not committed). See `SYNC_STEPS.md` for auto-update from OneDrive.
