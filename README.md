# Nakuru Health Collections Dashboard – shared Power Query code

Nakuru County Department of Health · ICT

Every copy of the dashboard workbook loads its query code from the `queries/` folder in this repo
each time it refreshes. **Edit a `.pq` file here and every user gets the change on their next refresh.**
There is no new file to send round.

## What is in here
| Path | What |
|---|---|
| `queries/*.pq` | One file per Power Query (M code), named exactly like the query (spaces → `_`). |
| `VERSION.txt` | When the files were last exported from the master workbook. |

## Rules
- **No secrets here.** API keys stay in each user's *Settings* sheet. Never commit a key or secret.
- **Write access:** county ICT only. Anyone who can push here changes what runs in every user's Excel.
- **Test first.** Change the file on a branch, point a test workbook's `P_CodeUrl` at that branch's raw URL, refresh, then merge to `main`.
- GitHub caches raw files for a few minutes, so users may see a change up to about 5 minutes after you push.

## Changing a query
1. Open `queries/<Name>.pq` on GitHub → ✏️ Edit.
2. Paste the new M code (starting with `let`, no leading `=`).
3. Commit to `main`. Users get it on the next refresh (automatic every 15–60 min, or Data → Refresh All).

## Adding a new query
1. Add `queries/<NewName>.pq` here.
2. In the master workbook, create a query named `<NewName>` with this loader, load it, then send the updated workbook once:
```
let
    env = Record.Combine({#shared, [fnParam = fnParam, pDay = pDay, fnGetList = fnGetList, Facilities = Facilities, Lib = Lib]}),
    out = Expression.Evaluate(fnRemote("<NewName>"), env)
in
    out
```

## Workbook side
- `fnParam` and `fnRemote` stay inside the workbook. Everything else loads from here.
- The Settings sheet named cell `P_CodeUrl` holds the base raw URL, for example
  `https://raw.githubusercontent.com/amohkinuthia/nakuru-dashboard-queries/main`.
- The macros in `NakuruRepo.bas` handle export, the switch to GitHub, and restore.
