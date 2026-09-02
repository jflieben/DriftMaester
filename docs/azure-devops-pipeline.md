# Publishing DriftMaester results in Azure DevOps

Maester can run [directly from an Azure DevOps pipeline](https://maester.dev/docs/monitoring/azure-devops/).
DriftMaester works a bit differently: the tests and drift detection run unattended in an Azure Automation
account, and the results are stored in your own blob storage. 

## How it fits together

```
Azure Automation (DriftMaester runbook)          Azure DevOps pipeline
  runs Maester + drift on a schedule                downloads latest XML from blob
  writes results to blob storage         --->       checks it is recent enough
  incl. NUnit XML at                                 publishes it to the Tests tab
  <prefix>/<tenantId>/latest/TestResults.xml
```

The runbook keeps doing all the heavy lifting and holds all the state (history, trend, suppressions) in
blob storage. The pipeline is stateless. It only downloads the newest `TestResults.xml`, verifies it is
fresh, and hands it to the built-in NUnit publisher. Because every scheduled run publishes the same set of
tests, Azure DevOps tracks which tests newly fail or newly pass across runs. That run-to-run comparison in
the Tests tab is your drift, on top of the HTML drift report the runbook already emails.

- Works on a plain Microsoft-hosted agent. No managed identity on the agent, no self-hosted pool.
- The pipeline stays small and testable: download, freshness check, publish.

## What the runbook produces

Every run of `Invoke-DriftMaester` writes an NUnit 2.5 test-results file to blob storage, in addition to the
existing JSON, HTML, and Markdown reports:

- Timestamped copy: `<prefix>/<tenantId>/results/TestResults-<yyyyMMdd-HHmmss>.xml`
- Fixed "latest" pointer: `<prefix>/<tenantId>/latest/TestResults.xml`

`<prefix>` defaults to `maester` and the container is also `maester`, so the newest results live at
`maester/<tenantId>/latest/TestResults.xml` inside the `maester` container. The XML reflects the results
*after* suppressions are applied, so it matches the HTML and email reports.

Result mapping (Maester result -> NUnit outcome):

| Maester result | NUnit outcome | Shown in Azure DevOps as |
|---|---|---|
| Passed | Success | Passed |
| Failed | Failure | Failed |
| Error | Error | Failed |
| Investigate | Inconclusive | Inconclusive |
| Skipped | Ignored (not executed) | Not executed |
| NotRun | Ignored (not executed) | Not executed |

## Prerequisites

1. A working DriftMaester install writing results to a storage account. See the main
   [README](../README.md) and run `Install-DriftMaester.ps1` if you have not already.
2. An Azure DevOps project and pipeline.
3. An Azure Resource Manager service connection. Workload identity federation is recommended, exactly as in
   the [Maester Azure DevOps guide](https://maester.dev/docs/monitoring/azure-devops/#set-up-the-azure-pipeline).
   The pipeline only needs to read blobs, so the service connection can be scoped narrowly.

## Setup

### 1. Grant the service connection read access to the storage account

The pipeline reads the blob using the service connection's identity (`--auth-mode login`), so that identity
needs data-plane read access, not just control-plane Reader.

1. Open **Project settings > Service connections** and select your ARM service connection.
2. Select **Manage Service Principal** to open the app registration / service principal in Entra.
3. Note its name or object id.
4. In the Azure portal, open the DriftMaester **storage account > Access Control (IAM)**.
5. Add a role assignment: **Storage Blob Data Reader**, assigned to the service connection's principal.

If the storage account is network-restricted, either allow the agent's outbound IPs, enable
"Allow trusted Microsoft services", or run the pipeline on a self-hosted agent inside the allowed network.

### 2. Add the pipeline

Copy [`Pipelines/azure-pipelines-driftmaester.yml`](../Pipelines/azure-pipelines-driftmaester.yml) into your
Azure DevOps repository (or point a new pipeline at it) and set the variables at the top:

| Variable | What to set it to |
|---|---|
| `azureServiceConnection` | Name of the ARM service connection with Storage Blob Data Reader. |
| `storageAccountName` | The DriftMaester storage account name (contains `maester`). |
| `containerName` | Blob container, default `maester`. |
| `blobPrefix` | Blob path prefix, default `maester`. |
| `tenantId` | The Entra tenant id that was scanned. |
| `maxAgeHours` | Freshness threshold. For a daily runbook, `26` gives a little slack. |
| `failOnFindings` | `true` to fail the pipeline on failed tests, `false` to only track trends. |

### 3. Align the schedule

Point the pipeline schedule a bit after the DriftMaester Automation schedule so fresh results are already in
blob when the pipeline runs. The sample uses `0 6 * * *` (06:00 UTC); adjust to sit after your runbook time.

## Viewing results

- **Pipelines > Runs**: pick a run, then the **Tests** tab shows every Maester test with its status.
- **Test analytics** and the pass/fail history across runs show which tests regressed or improved, which is
  the drift view.
- The **driftmaester-results** artifact holds the raw `TestResults.xml` for offline analysis.

For the full DriftMaester drift report (with trend chart and per-finding detail), open the HTML report the
runbook emails, or the `<prefix>/<tenantId>/reports/DriftReport-*.html` blob in storage.

## Freshness check

The pipeline reads the blob's `lastModified` timestamp before publishing. If the newest results are older
than `maxAgeHours`, the pipeline fails with a clear error. That turns a silently broken or disabled runbook
into a visible red pipeline instead of quietly republishing yesterday's results.

## Troubleshooting

- **"No DriftMaester results found"**: the runbook has not run for this tenant yet, or `storageAccountName`,
  `containerName`, `blobPrefix`, or `tenantId` are wrong. Check the blob path
  `<blobPrefix>/<tenantId>/latest/TestResults.xml` exists in the container.
- **Results are X hours old, exceeding the threshold**: the scheduled Automation run failed or is disabled.
  Check the Automation account job history for `Invoke-DriftMaester`.
- **`AuthorizationPermissionMismatch` / 403 on download**: the service connection principal is missing
  **Storage Blob Data Reader** on the storage account, or the assignment has not propagated yet.
- **Network / firewall errors**: the storage account blocks the agent. Allow the agent network or use a
  self-hosted agent.