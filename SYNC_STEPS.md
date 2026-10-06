# Portfolio auto-sync procedure

The Excel source lives in Tornike’s OneDrive. **OneDrive access is MCP-only** — shell scripts cannot call it. A Grok Bot / agent run must perform the MCP steps, then the shell steps.

## Constants

| Key | Value |
|-----|-------|
| MCP namespace | `user-Onedrive--tornike-gabunia-ugt-ge` (**never** `user-Onedrive`) |
| Item id | `01DXQX44UQPCQGPFTZKZBJ3BMECRZ5LYJJ` |
| Drive id (informational) | `b!cwg0GibUcUyEKhWHfrabaLX1nUASQu9Kn9ssh5ENp20YI8iuV0_XRLZducsbgBhr` |
| Local Excel | `/workspace/ugt-portfolio/UGT_Product_Portfolio_2026.xlsx` |
| State file | `/workspace/ugt-portfolio/state.json` |
| Repo | `/workspace/ugt-portfolio/repo` → `BRT928/ugt-cloud-portfolio` |
| Sheet | `UGT Product Portfolio ALL` |

## Step A — MCP: cheap change check

1. `CallDynamicTool` / `get_drive_item`  
   - namespace: `user-Onedrive--tornike-gabunia-ugt-ge`  
   - arguments: `{ "itemId": "01DXQX44UQPCQGPFTZKZBJ3BMECRZ5LYJJ", "select": "id,name,size,eTag,cTag,lastModifiedDateTime" }`
2. Write the JSON response to e.g. `/tmp/ugt-drive-item.json`.
3. Shell: `python3 /workspace/ugt-portfolio/sync.py --meta /tmp/ugt-drive-item.json --check-only`  
   - Prints `NO_CHANGE` → stop.  
   - Prints `CHANGED` (+ summary) → continue.

## Step B — MCP: download if changed

4. Built-in `DownloadFile`:  
   - `connection`: `user-Onedrive--tornike-gabunia-ugt-ge`  
   - `source.fileId`: `01DXQX44UQPCQGPFTZKZBJ3BMECRZ5LYJJ`  
   - `destination.path`: `/workspace/ugt-portfolio/UGT_Product_Portfolio_2026.xlsx`

## Step C — Shell: rebuild, commit, push

5. `python3 /workspace/ugt-portfolio/sync.py --meta /tmp/ugt-drive-item.json --apply`  
   This will:
   - Update `state.json` with new eTag / cTag / lastModifiedDateTime
   - Run `build.py` → refresh `docs/` and `products.json`
   - Diff previous vs new `products.json` (added / removed / field changes)
   - rsync `docs/` + tooling into `repo/`, commit, `git push`
   - Print the change summary

## Forced rebuild (no OneDrive)

`python3 /workspace/ugt-portfolio/sync.py --force-apply`

## Notes

- Do **not** commit the `.xlsx` or any tokens (see `.gitignore`).
- GitHub Pages serves from `/docs` on `main`.
- Live site: https://brt928.github.io/ugt-cloud-portfolio/
