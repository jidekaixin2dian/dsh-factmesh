# FactMesh

Review PowerPoint figures against Excel in DeepSeek Harness. Open a result to see the slide, workbook cell, period, unit, raw figures and comparison. Confirmed links can be reused on a later report.

## Installation

The verified host is DSH Web `0.2.2-rc.1` with `workspaceFiles`. Stop the running Web host for this profile, install the plugin, and start Web again.

```powershell
dsh plugin --profile web add github:jidekaixin2dian/dsh-factmesh
dsh web --host 127.0.0.1 --port 17321
```

Open the URL printed by DSH. Add a workspace, open a session in it, and select **FactMesh · Reconcile** in the sidebar. On a fresh host, choose **Add API Key** and then **Configure later** to proceed without a provider.

For offline installation, download the `.tgz` from [GitHub Releases](https://github.com/jidekaixin2dian/dsh-factmesh/releases) and pass its local path in place of the GitHub source.

## Review a report

1. Pick a workspace session, local project name, reporting month, PPTX and XLSX.
2. Read the worksheet list, choose a worksheet and select **Reconcile**. Open **Coverage** to review omitted objects.
3. Select a figure, inspect the candidate evidence, and confirm its link. Enter a cell address to inspect a different candidate.
4. For the next report, keep the same browser, Web address, workspace and project name. Choose the new files and month. The evidence panel shows the original and current locations when a link is reused.

Links are stored in browser site data. Keep the same Web origin for the project across restarts. The selected Office files are read through the host's read-only workspace interface.

The first release supports ordinary text boxes, native PowerPoint tables, and ordinary Excel value grids with month headers. It lists images, embedded charts and unsupported objects under Coverage. Formula results and ambiguous formats require verification.

[中文用法与两期演示](README.md)
