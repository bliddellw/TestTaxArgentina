# TestTaxArgentina

Automatically fetches Argentina individual income-tax (PIT) band data from
[PwC Tax Summaries](https://taxsummaries.pwc.com/argentina/individual/taxes-on-personal-income)
on a quarterly schedule and appends marginal-rate rows to an Excel workbook
stored in OneDrive.

---

## Repository structure

```
.
├── scripts/
│   └── fetch_ar_pit_bands.py   # Scraper + Graph API integration
├── .github/
│   └── workflows/
│       └── quarterly-ar-tax.yml  # Scheduled GitHub Actions workflow
├── requirements.txt
└── README.md
```

---

## How it works

1. The workflow runs on the 1st of **January, April, July, and October** at 06:00 UTC
   (or on demand via `workflow_dispatch`).
2. The Python script fetches the PwC Tax Summaries page for Argentina.
3. It extracts the PIT table (columns: Over, Not over, Tax on col 1, % on excess).
4. It converts the "% on excess" values to **marginal band rates**:
   - The first band (min = 0) always has rate 0%.
   - Each subsequent band's rate equals the *previous row's* "% on excess".
5. It appends the rows to `BandsTable` in the worksheet `Bands` of the
   OneDrive workbook `/TaxRates/Argentina_PIT_Bands.xlsx`.

### Output table columns

| Column          | Type              | Description                                          |
|-----------------|-------------------|------------------------------------------------------|
| `snapshot_date` | ISO date string   | Date the workflow ran (UTC), e.g. `2026-04-01`       |
| `band_min`      | integer           | Lower bound of the income band (ARS)                 |
| `band_max`      | integer or blank  | Upper bound of the income band; blank for top band   |
| `band_rate`     | decimal fraction  | Marginal rate, e.g. `0.05` for 5%                   |

Running the workflow twice appends two separate snapshot groups — existing
rows are **never** overwritten.

---

## Setup

### 1. Register an Azure AD (Entra ID) application

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).
2. Click **New registration** → give it a name (e.g. `ArgentinaTaxAgent`).
3. Under **Authentication**, choose *Accounts in this organizational directory only*.
4. Note the **Application (client) ID** and **Directory (tenant) ID**.
5. Go to **Certificates & secrets → New client secret** → copy the secret value.

### 2. Grant Microsoft Graph API permissions

In the app registration, go to **API permissions → Add a permission →
Microsoft Graph → Application permissions** and add:

| Permission                  | Reason                                       |
|-----------------------------|----------------------------------------------|
| `Files.ReadWrite.All`       | Create/read/write files in OneDrive          |

Click **Grant admin consent** for your tenant.

> **Note:** `Files.ReadWrite.All` with client credentials acts on behalf of
> the service principal (app-only). To write to a *specific user's* OneDrive,
> you may need to use a delegated flow or a shared/service-account drive.
> Alternatively, add `Sites.ReadWrite.All` to target a SharePoint document
> library and adjust `ONEDRIVE_FILE_PATH` accordingly.

### 3. Add GitHub Actions secrets

In the repository go to **Settings → Secrets and variables → Actions** and
create:

| Secret name        | Value                               |
|--------------------|-------------------------------------|
| `MS_TENANT_ID`     | Azure AD tenant ID (GUID)           |
| `MS_CLIENT_ID`     | App registration client ID (GUID)   |
| `MS_CLIENT_SECRET` | App registration client secret      |

### 4. First run

You can trigger the workflow manually:

1. Go to **Actions → Quarterly Argentina PIT Bands → Run workflow**.
2. The workbook `/TaxRates/Argentina_PIT_Bands.xlsx` will be created in
   OneDrive automatically if it does not exist.

---

## Testing locally

### Install dependencies

```bash
pip install -r requirements.txt
```

### Dry run (no OneDrive writes)

```bash
python scripts/fetch_ar_pit_bands.py --dry-run
```

This prints the parsed and transformed rows to stdout so you can verify the
scraping logic without needing Azure credentials.

### Full run

```bash
export MS_TENANT_ID=<your-tenant-id>
export MS_CLIENT_ID=<your-client-id>
export MS_CLIENT_SECRET=<your-client-secret>
python scripts/fetch_ar_pit_bands.py
```

### Custom OneDrive path

```bash
export ONEDRIVE_FILE_PATH=/MyFolder/CustomName.xlsx
python scripts/fetch_ar_pit_bands.py
```

---

## Troubleshooting

- **Table not found**: The PwC page structure may have changed. Run with
  `--dry-run` and inspect the error message for details.
- **401 Unauthorized from Graph API**: Verify your tenant/client credentials
  and that admin consent has been granted for the required permissions.
- **404 on drive**: Ensure the authenticated app/user has a OneDrive provisioned.
