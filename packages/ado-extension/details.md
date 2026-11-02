Open saved queries or selected work items in Excel, edit them in bulk, and publish the changes back to Azure DevOps.

This is a modern replacement for the Microsoft DevLabs **Azure DevOps Open in Excel** marketplace extension. That extension launched Excel through a `tfs://` protocol handler and the Windows-only Azure DevOps Office Integration COM add-in, which Microsoft no longer supports.

The Azure DevOps menu command and the Excel add-in are both named **ADO in Excel** so they are distinct from the legacy **Open in Excel** control if both extensions are installed.

## What you need

1. This Azure DevOps extension (query and work item **ADO in Excel** commands).
2. The companion **ADO in Excel** Office.js add-in in Excel for Windows — sideloaded for development, or deployed in production via Microsoft 365 admin **Integrated apps** / Centralized Deployment (or Marketplace) after hosting the add-in on HTTPS. This is a web add-in (manifest + hosted pages), not a signed COM/VSTO install.
3. Auth: **Connect with PAT** (Work Items read/write and Identity read), **or** **Sign in with Microsoft** (Entra ID). PAT-only use does not require an Entra app registration.

## Sign in with Microsoft (Entra ID)

Skip this section if you only use a Personal Access Token.

Register an app in the Azure portal, then use its Application (client) ID in Excel:

1. Register a **single-page application** (SPA).
2. Add SPA redirect URI `https://bcurtisbdo.github.io/ADOExcel/auth.html` (and `https://localhost:3001/auth.html` if you also sideload for development). The redirect is required for Entra/MSAL only—not for PAT.
3. Add delegated permission **Azure DevOps → user_impersonation**. Grant admin consent if your tenant requires it.
4. Use an **organizational** (work or school) account. Azure DevOps REST via Entra is intended for those accounts, not personal Microsoft accounts.
5. In the Excel task pane, enter your Azure DevOps organization URL and the Application (client) ID, then click **Sign in with Microsoft**.

## Open a query

1. Go to **Boards > Queries**.
2. Right-click a saved query, or open it and use **ADO in Excel** on the results toolbar.
3. In the dialog, wait for the workbook to build, then click **Download workbook** (the browser cannot start Excel). Open the `.xlsx` file in Excel for Windows. If Excel shows Protected View, click **Enable Editing** so Publish can read the binding sheet. The binding metadata sheet stays hidden.
4. Tree queries use one **Title** column; the Excel add-in indents and groups child rows and shades each nest level.
5. Edit cells, then use **Publish** in the Excel add-in task pane to write changes back. Indent/Outdent reparents; Area Path and Iteration Path changes move items in the project tree.

Unsaved and built-in queries (Assigned to me, Followed work items, Recycle bin, and similar) are not supported. Save the query under **My Queries** or **Shared Queries** first.

## Open selected work items

Select one or more work items from query results, a backlog, or a work item form, then choose **ADO in Excel**.

## Publish and refresh

The downloaded workbook is immediately useful for analysis. Publishing and refreshing require the Excel add-in (version **1.2.0**). Types, states, picklists, and area/iteration paths come from this project's process template. After you connect in Excel, **Load a query** is under **Connection**; the workbook stays bound so Publish, Refresh, and in-cell dropdowns keep working. To switch organizations, enter the new URL and credentials and click **Connect**; a sheet loaded from another organization can't be refreshed or published until you connect back to it. Tree lists shade each nest level so parents and children are easier to see.

- **Refresh** reloads current Azure DevOps values into the sheet.
- **Publish** sends changed fields concurrently (with progress in the task pane), creates rows that have no ID, updates parent/child links from Title indent, syncs published rows, and reports revision conflicts so you can refresh and retry. Changing the Work Item Type of an existing row isn't supported; Publish reports it so you can change the type in Azure DevOps or undo the edit.
- **Open query in Azure DevOps** opens the saved query the sheet was loaded from.
- **Insert work item** adds a new row of a process type (with a valid default state). Set Title, then publish.
- **Columns** add, remove, or reorder fields (Top, Up, Down, Bottom, or Alt+Up/Down; add several fields at once below the selected column), then **Apply changes** to update the sheet in one step; new columns are filled from Azure DevOps. **Discard** drops pending changes. ID, Work Item Type, Title, and Rev stay in the list. **Parent Title** is optional and read-only (resolved from the parent work item).
- **Area and Iteration** pickers apply a node to the selected cell. Excel also limits those columns to this project's tree.
- **Assigned To** uses Excel's in-cell dropdown, filled from this project's teams and people already on the sheet. Search in the task pane only if the person is not listed.
- **Convert to tree**, **Indent**, and **Outdent** match the classic Team ribbon tree list, using Title indent, row grouping, and per-level row shading instead of Title 1, Title 2, ….
- Task pane sections collapse. The project list is alphabetical; saved queries show as an expandable folder tree under **Connection**.
- Date fields display as short dates (for example 9/18/2026), not ISO timestamps.
- Cells are top-aligned, so multi-line fields such as Description read from the first line.

HTML fields such as Description edit as Markdown in the cell; publish converts Markdown to HTML. If you paste HTML into the cell instead (tags such as `div`, `p`, `br`, `b`, `a`), that HTML is published as-is. Unchanged values that still match the workbook snapshot are not rewritten.

## View in Browser

Select a work item row that already has an ID, then open it in Azure DevOps from:

- the **ID** hyperlink on the sheet
- **View in Browser** on the Home ribbon **ADO in Excel** group
- **View in Browser** on the cell context menu
- **Open in Azure DevOps** in the Excel add-in task pane

Publish new rows before opening them.

## Related links, attachments, and discussion

These stay off the grid. Select a row, click **Load details** in the task pane, then:

- see the parent work item ID and title (when one exists)
- list, open, add (by work item ID), or remove links of any work item link type your organization allows: Related, Predecessor/Successor, Duplicate/Duplicate Of, Tested By/Tests, Affects (CMMI), and custom types
- add a hyperlink (http or https) with an optional comment; commit, branch, and build links are listed and open from the work item in Azure DevOps
- list attachments, **Open attachment** (when a URL is available), attach a file, or remove one
- read discussion comments and add a comment

Add and remove act on the work item shown in the panel: if you select a different row, click **Load details** again first. The panel clears when you load a query, refresh, publish, or connect. Parent/child uses Title indent on publish; optional **Parent Title** on the sheet is display-only. If **Open attachment** cannot open a direct file URL, use the work item in Azure DevOps.
