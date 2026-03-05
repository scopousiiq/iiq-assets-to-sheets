# iiq-assets-to-sheets

Google Apps Script project that extracts Incident IQ asset data into Google Sheets for device reporting, Chromebook AUE tracking, and replacement cycle planning.

## What You Get

**Data loaded automatically from your iiQ instance:**
- Complete asset inventory with model, manufacturer, location, status, and purchase info
- AUE (Auto Update Expiration) dates from custom fields
- Location directory and asset status types

**Pre-built analytics sheets (formula-driven, auto-updating):**
- **Location Summary** — Asset counts, AUE status, and average age per school
- **Model Breakdown** — Device model inventory with active/retired counts
- **AUE Planning** — Replacement timeline by fiscal year
- **Budget Planning** — Estimated replacement costs per location
- **Status Overview** — Active, retired, in storage breakdown

**Calculated columns on every asset:**
- AUE Status (Expired / < 6 Months / < 1 Year / < 2 Years / OK)
- Device age in days and years
- Warranty status (Active / Expiring / Expired)
- Replacement cycle (fiscal year when AUE expires)

## Quick Start

1. Create a new Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Copy all `.gs` files from the `scripts/` directory
4. Save and refresh the spreadsheet
5. Run **iiQ Assets > Setup > Setup Spreadsheet**
6. Fill in the **Config** sheet:
   - `API_BASE_URL`: Your iiQ instance URL (e.g., `https://yourdistrict.incidentiq.com`)
   - `BEARER_TOKEN`: Your API token
7. Run **iiQ Assets > Setup > Verify Configuration**
8. Run **iiQ Assets > Load Reference Data > Refresh All Reference Data**
9. Run **iiQ Assets > Asset Data > Continue Loading** (repeat until complete for large inventories)
10. Run **iiQ Assets > Asset Data > Enrich Custom Fields** (populates AUE dates)
11. Run **iiQ Assets > Asset Data > Apply Formulas**
12. Run **iiQ Assets > Setup > Setup Automated Triggers** for weekly refresh

## AUE Date Support

AUE (Auto Update Expiration) is the date when a Chromebook stops receiving Chrome OS updates. This project supports AUE tracking via iiQ custom fields:

1. **Automatic detection**: When you run "Discover Custom Fields", the system searches for fields named like "AUE", "Auto Update Expiration", or "End of Life"
2. **Manual configuration**: If auto-detection doesn't find it, set `AUE_CUSTOM_FIELD_ID` in the Config sheet to your custom field's UUID

If your district doesn't use a custom field for AUE, the AUE-related columns and analytics will simply show blank — everything else works independently.

## Automated Refresh

After initial setup, enable automated triggers:
- **Every 10 minutes**: Continues any interrupted loading
- **Weekly (Sunday 2 AM)**: Full data refresh including reference data

## Connecting to Looker Studio / Power BI

Connect your BI tool directly to this Google Spreadsheet:
- **AssetData** sheet for detailed device-level data
- **Analytics sheets** for pre-aggregated summaries
- Formula columns (AUEStatus, AgeYears, WarrantyStatus, ReplacementCycle) are ready for filtering and grouping

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "API configuration missing" | Fill in API_BASE_URL and BEARER_TOKEN in Config sheet |
| Loading stops partway | Run "Continue Loading" again — it resumes from where it left off |
| No AUE data | Check CustomFields sheet, set AUE_CUSTOM_FIELD_ID manually if needed |
| Analytics show #REF errors | Run "Apply Formulas" after loading completes |
| Need fresh data | Run "Full Reload" (removes triggers first) or wait for weekly auto-refresh |

## Project Structure

```
scripts/
  Config.gs          - Configuration, logging, concurrency control
  ApiClient.gs       - HTTP client with retry/backoff
  AssetData.gs       - Two-phase asset loader (bulk + custom field enrichment)
  ReferenceData.gs   - Locations, status types, custom field discovery
  Setup.gs           - Sheet creation and analytics formulas
  Menu.gs            - Menu system and UI operations
  Triggers.gs        - Automated trigger management
  appsscript.json    - Apps Script manifest
```

## License

Open source — use as a base for your district's asset reporting needs.
