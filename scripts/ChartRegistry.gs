/**
 * ChartRegistry — declarative map of analytics sheets to chart specifications.
 *
 * The Dashboard web app iterates this registry, keeps entries whose sheetName
 * exists in the workbook, reads each sheet's data using the spec, and ships a
 * serializable payload to the client for Chart.js rendering.
 *
 * Entry shape:
 *   {
 *     sheetName:   string — must exactly match the sheet name created by Setup/OptionalAnalytics
 *     category:    string — full category name (section/tab grouping)
 *     tabLabel:    string — short name shown on the tab bar
 *     charts:      array  — one or more chart specs (see Spec shape below)
 *     kpiBadge:    object — alternative to `charts` for list-only sheets
 *   }
 *
 * Spec shape:
 *   {
 *     title:    string — card header
 *     type:     'bar' | 'line' | 'stackedBar' | 'horizontalBar' | 'stackedHorizontalBar'
 *     labelCol: number — 0-indexed column for category/axis labels
 *     series:   [ { header: string, col: number, color: string } ]
 *     rowStart: number — 1-indexed first data row (row 1 is normally headers)
 *     rowMode:  'contiguous' — read until labelCol goes blank
 *               | 'fixed'    — take exactly rowCount rows from rowStart
 *     rowCount: number — required when rowMode='fixed'
 *   }
 *
 * Badge shape:
 *   { label: string, color: string, rowStart: number }
 *
 * Colors map to CSS tokens defined in DashboardApp.html (IIQ.* palette).
 * Available: darkBlue, teal, gold, orange, purple, slate, red, lightBlue.
 *
 * Sheet schemas verified against Setup.gs / OptionalAnalytics.gs header arrays.
 * When adding a new analytics sheet: register it here or the dashboard won't
 * discover it.
 */

const CHART_REGISTRY = [

  // ───────────────────────────── Fleet Composition ───────────────────────────
  {
    sheetName: 'LocationSummary',
    category: 'Fleet Composition',
    tabLabel: 'Composition',
    charts: [{
      title: 'Active vs Retired Assets per Location (Top 20)',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: 'Active',  col: 2, color: 'darkBlue' },
        { header: 'Retired', col: 3, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'ModelBreakdown',
    category: 'Fleet Composition',
    tabLabel: 'Composition',
    charts: [{
      title: 'Top Device Models',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Total', col: 2, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },
  {
    sheetName: 'CategoryBreakdown',
    category: 'Fleet Composition',
    tabLabel: 'Composition',
    charts: [{
      title: 'Devices by Category',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Total', col: 1, color: 'teal' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'ManufacturerSummary',
    category: 'Fleet Composition',
    tabLabel: 'Composition',
    charts: [{
      title: 'Devices by Manufacturer',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Device Count', col: 1, color: 'purple' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Status & Operations ─────────────────────────
  {
    sheetName: 'StatusOverview',
    category: 'Status & Operations',
    tabLabel: 'Status',
    charts: [{
      title: 'Asset Status Distribution',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Count', col: 1, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'AssignmentOverview',
    category: 'Status & Operations',
    tabLabel: 'Status',
    charts: [{
      title: 'Assigned vs Unassigned per Location (Top 20)',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: 'Assigned',   col: 2, color: 'darkBlue' },
        { header: 'Unassigned', col: 3, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'DeviceReadiness',
    category: 'Status & Operations',
    tabLabel: 'Status',
    charts: [{
      title: 'Device Readiness by Location (Top 20)',
      type: 'stackedHorizontalBar',
      labelCol: 0,
      series: [
        { header: 'Deployable',  col: 2, color: 'teal' },
        { header: 'In Repair',   col: 3, color: 'gold' },
        { header: 'Lost/Stolen', col: 4, color: 'orange' },
        { header: 'Retired',     col: 5, color: 'slate' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'SpareAssets',
    category: 'Status & Operations',
    tabLabel: 'Status',
    charts: [{
      title: 'Deployable Spares per Location (Top 20)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Deployable Spares', col: 3, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },

  // ───────────────────────────── Aging & Warranty ────────────────────────────
  {
    sheetName: 'AgingAnalysis',
    category: 'Aging & Warranty',
    tabLabel: 'Aging',
    charts: [{
      title: 'Fleet Age Distribution by Year',
      type: 'stackedBar',
      labelCol: 0,
      series: [
        { header: 'Active',  col: 2, color: 'teal' },
        { header: 'Retired', col: 3, color: 'slate' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'WarrantyTimeline',
    category: 'Aging & Warranty',
    tabLabel: 'Aging',
    charts: [{
      title: 'Devices with Expiring Warranty by Quarter',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Devices Expiring', col: 1, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },
  {
    sheetName: 'DeviceLifecycle',
    category: 'Aging & Warranty',
    tabLabel: 'Aging',
    charts: [{
      title: 'Average Lifespan by Model (Years)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Avg Lifespan (Years)', col: 3, color: 'purple' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },

  // ───────────────────────────── Replacement Planning ────────────────────────
  {
    sheetName: 'BudgetPlanning',
    category: 'Replacement Planning',
    tabLabel: 'Budget',
    charts: [{
      title: 'Estimated Replacement Cost per Location (Top 20)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Est. Replacement Cost', col: 5, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'ReplacementPlanning',
    category: 'Replacement Planning',
    tabLabel: 'Budget',
    charts: [{
      title: 'New Replacements Needed by Target Date (Top 20)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'New Replacements Needed', col: 4, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'ReplacementForecast',
    category: 'Replacement Planning',
    tabLabel: 'Budget',
    charts: [{
      title: 'Devices to Replace by Year',
      type: 'bar',
      labelCol: 0,
      series: [
        { header: 'Device Count', col: 1, color: 'darkBlue' }
      ],
      rowStart: 2,
      rowMode: 'contiguous'
    }]
  },

  // ───────────────────────────── Service & Risk ──────────────────────────────
  {
    sheetName: 'ServiceImpact',
    category: 'Service & Risk',
    tabLabel: 'Service',
    charts: [{
      title: 'Tickets per Device by Model (Top 15)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Tickets / Device', col: 4, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },
  {
    sheetName: 'BreakRate',
    category: 'Service & Risk',
    tabLabel: 'Service',
    charts: [{
      // BreakRate has a model-summary block in cols I-M (0-indexed 8-12).
      // Row 1 is its own header row; data starts at row 2.
      title: 'Avg Tickets per Device by Model (Top 15)',
      type: 'horizontalBar',
      labelCol: 8,
      series: [
        { header: 'Avg Tickets/Device', col: 11, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },
  {
    sheetName: 'HighTicketLocations',
    category: 'Service & Risk',
    tabLabel: 'Service',
    charts: [{
      title: 'Tickets per Device by Location (Top 20)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Tickets / Device', col: 4, color: 'red' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 20
    }]
  },
  {
    sheetName: 'LostStolenRate',
    category: 'Service & Risk',
    tabLabel: 'Service',
    charts: [{
      title: 'Loss Rate (%) by Location (Top 15)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Rate (%)', col: 5, color: 'orange' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },
  {
    sheetName: 'ModelFragmentation',
    category: 'Service & Risk',
    tabLabel: 'Service',
    charts: [{
      title: 'Distinct Models per Location (Top 15)',
      type: 'horizontalBar',
      labelCol: 0,
      series: [
        { header: 'Distinct Models', col: 2, color: 'purple' }
      ],
      rowStart: 2,
      rowMode: 'fixed',
      rowCount: 15
    }]
  },

  // ───────────────────────────── Lookups (interactive) ───────────────────────
  // These tabs render a form (dropdown or text input) and call the named server
  // function via google.script.run when submitted. Results render as an HTML
  // table. Only included when the source sheet exists in the workbook.
  {
    sheetName: 'IndividualLookup',
    category: 'Individual Lookup',
    tabLabel: 'Individual',
    lookup: {
      type: 'individual',
      inputType: 'dropdown',
      inputLabel: 'Select User:',
      placeholder: '— Choose a user —',
      serverFunction: 'dashboardLookupIndividual',
      optionsFunction: 'dashboardGetOwnerNames'
    }
  },
  {
    sheetName: 'VerificationLookup',
    category: 'Verification Lookup',
    tabLabel: 'Verification',
    lookup: {
      type: 'verification',
      inputType: 'text',
      inputLabel: 'Asset Tag or Serial Number:',
      placeholder: 'Paste or type and press Enter',
      serverFunction: 'dashboardLookupVerification'
    }
  }
];

/** Ordered category list — drives tab bar order. */
const CATEGORY_ORDER = [
  'Fleet Composition',
  'Status & Operations',
  'Aging & Warranty',
  'Replacement Planning',
  'Service & Risk',
  'Individual Lookup',
  'Verification Lookup'
];
