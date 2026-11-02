# Changelog

## 1.2.0

- **Add link** offers every work item link type the organization allows, grouped as Related, Dependency (Predecessor, Successor), Duplicate (Duplicate, Duplicate Of), Test (Tested By, Tests), Affects (CMMI only), and any custom types. The list is loaded from Azure DevOps after you connect and falls back to Related, Predecessor, and Successor if it can't be loaded. Parent/child still comes from Title indent; links to other organizations and Test Plans' shared step links aren't offered.
- **Add hyperlink** in the Selected work item panel adds a web link with an optional comment. Only http and https addresses are accepted; a bare address such as `example.com/spec` is treated as https.
- The related links list uses Azure DevOps' names for every link type (for example **Tested By** instead of "TestedBy-Forward"), shows hyperlink comments or addresses, and names commit, branch, and build links (for example **Fixed in Commit**).
- When Azure DevOps rejects a link (for example a second **Duplicate Of**), the panel shows Azure DevOps' message instead of the raw response.
- Links to work items in another organization open in that organization; previously they opened the same ID in the current one.
- **Open link** explains why commit, branch, build, and non-web hyperlinks can't be opened from Excel, instead of trying to open them.
- Workbooks are top-aligned. Downloaded workbooks and sheets written by the add-in (load, Refresh, Publish, Apply columns) align cells to the top, so multi-line fields such as Description read from the first line.

## 1.1.1

- Fixed **Apply changes** in Columns discarding the project's process data (picklists, states, people, area and iteration paths) that was saved to the workbook in the same step, which could leave in-cell dropdowns using outdated values.

## 1.1.0

- **Open query in Azure DevOps** in the Workbook section opens the saved query the sheet was loaded from. It's hidden for sheets opened from selected work items.
- Column changes are staged. Move (Top, Up, Down, Bottom, Alt+Up/Down), remove, and add columns instantly in the task pane, then **Apply changes** rewrites the sheet once instead of after every click. Several fields can be added at once, inserted below the selected column, and the field list can be filtered. **Discard** drops pending changes.
- Newly added columns are filled from Azure DevOps when applied, instead of staying blank until the next Refresh. The fetched values are recorded as the loaded values, so Publish doesn't treat them as edits.
- The **Selected work item** panel clears when you load a query, refresh, publish, connect, or sign in, so it no longer shows details from an earlier sheet.
- Adding or removing links, attachments, and comments checks that the selected row is the work item shown in the panel, and stops with a message if it isn't. Previously, removing a link after selecting a different row could remove an unrelated link from that row's work item.
- Removing a link or attachment fails safely if the work item changed since its details were loaded, instead of removing whatever is now in that position.
- Publish rejects a changed Work Item Type on an existing row with a clear error, instead of silently ignoring it and checking State against the wrong type. Change the type in Azure DevOps, or set the cell back.

## 1.0.1

- Switching organizations works. Connect and Sign in are available as soon as the task pane opens, even if the saved organization is slow or unreachable. Project and query lists from the previous organization can no longer overwrite the new one.
- The cached process template is tied to the organization as well as the project, and is cleared when you connect, so dropdowns and validation come from the right organization.
- Refresh, Publish, Load details, links, attachments, comments, and adding Parent Title are blocked when the sheet came from a different organization than the one you are connected to. The workbook status shows the sheet's organization and warns about a mismatch. Open in Azure DevOps uses the sheet's organization.
- Azure DevOps requests time out after 60 seconds (5 minutes for attachment uploads) with a clear message instead of hanging.

## 1.0.0

- First production release. Version raised to 1.0.0 (manifest `1.0.0.0`) so the Office manifest validator accepts it for Microsoft 365 admin deployment and Marketplace submission.
- `packages/excel-addin/manifest.xml` points at `https://bcurtisbdo.github.io/ADOExcel/`, so the add-in runs without a local web server. Install it with `npm run sideload:addin`.
- Local development uses a separate `manifest.dev.xml` (`https://localhost:3001/`) with its own add-in ID, shown as **ADO in Excel (Dev)**. Running `npm run start:addin` no longer replaces the hosted add-in.
- Entra sign-in resolves `auth.html` next to the current page, so it works when the add-in is hosted under a sub-path such as GitHub Pages. Add `https://bcurtisbdo.github.io/ADOExcel/auth.html` as a redirect URI on the app registration.
- Publish progress always reaches its total, including rows that fail to save.
- Docs: screenshots in the README, release notes moved to this file.

## 0.3.12

- Task pane **Load a query** lives under **Connection** (project + saved-query tree), not a separate section.
- **Publish** updates work items and parent links concurrently, shows progress (`Publishing n/m…`), and syncs only published rows afterward instead of a full refresh.
- Work item batch fetches run in parallel; Indent/Outdent refresh **Parent Title** without rebuilding process dropdowns (Load, Refresh, Publish, and column changes still reapply State, Area Path, and other lists).
- **Open attachment** in Selected work item opens the attachment URL when available.
- Docs: feature-parity table vs legacy extension + COM, production Office.js deployment, WebView debug sidestep (`start:addin:nodebug`), HTML-as-is on publish, and Entra vs PAT registration notes.

## 0.3.11

- Optional **Parent Title** column (add/remove from Columns); read-only, resolved from the parent work item on load, refresh, and when the column is added.
- **Selected work item** details show parent ID and title after **Load details**.

## 0.3.10

- The Azure DevOps download dialog contribution uses `ms.vss-web.external-content` so the dialog content renders with the modern extension SDK (empty or title-only dialogs with `ms.vss-web.control` are fixed).

## 0.3.9

- Download dialog failures report load failure to Azure DevOps instead of claiming success after an error message.
- Leaving a tree list clears nest-level row fills even when a Title column is not present.

## 0.3.8

- The Azure DevOps **ADO in Excel** command no longer loads two copies of the Azure DevOps SDK (that console error blocked the download dialog).
- Tree lists shade each indent level with a soft green fill so nest depth is easier to scan.

## 0.3.7

- The Azure DevOps **ADO in Excel** dialog builds the workbook and shows **Download workbook** instead of staying blank (SDK is loaded once in the dialog page).
- Downloaded workbooks keep the metadata sheet hidden. The download link no longer uses `target="_blank"`, so browsers save the file instead of opening a new tab.

## 0.3.6

- After **Load query**, the workbook stays bound to Azure DevOps. Publish and Refresh no longer report that the file is not bound.
- Process dropdowns (Work Item Type, State, Assigned To, picklists, Area Path, Iteration Path) apply again after load.

## 0.3.5

- **Download workbook** in the Azure DevOps dialog is the download action. The browser cannot start Excel; open the `.xlsx` and click **Enable Editing** if Protected View is on.
- Publish can read the binding sheet on a downloaded workbook (it is no longer stuck as “not bound” after a silent iframe download or a hidden metadata sheet).
- Date columns keep short-date display when the service returns ISO timestamps.

## 0.3.4

- Azure DevOps commands and the Excel add-in are named **ADO in Excel**, distinct from the legacy Open in Excel control.
- The browser action downloads a workbook (click **Download workbook** if the file does not start automatically). Open the `.xlsx` in Excel; Enable Editing if Protected View is on.
- Saved queries in the task pane are an expandable folder tree.
- Date fields show as short dates, not ISO timestamps.
- Downloaded workbooks keep Azure DevOps binding metadata so Publish works without loading the query again.

## 0.3.3

- Tree lists use one **Title** column. Nesting is Excel indent and outline grouping; **Indent** / **Outdent** still publish `Hierarchy-Reverse` parent links. Older Title 1…N workbooks still load.
- Task pane sections collapse and stay collapsed.
- Project list is alphabetical. Saved queries show as an indented folder tree, also alphabetical.

## 0.3.2

- Load, refresh, and publish keep the Home ribbon. Same-column updates write into the existing table instead of recreating it, so Excel does not switch to Table Design.

## 0.3.1

- **View in Browser** from ID hyperlinks, the Home ribbon, the cell context menu, and the task pane.
- **Selected work item** panel for related/predecessor/successor links, attachments, and discussion.
