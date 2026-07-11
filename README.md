# Construction Cost Dashboard (Excel / VBA / Power Query)

A supplier-based cost tracking and reporting system for utility-scale construction projects — built entirely in Excel with VBA, Power Query, and dynamic Form Controls.

Developed as a personal tool to solve a real problem on renewable energy construction sites (solar PV & battery storage): **site teams had no live visibility of cost commitments** across dozens of suppliers, hire contracts, and delivery streams. Commercial reporting arrived weeks after the spend happened. This workbook closes that gap at site level.

> **This is a fully anonymised demo.** All supplier names, prices, and project data are fictional samples. The structure, formulas, macros, and queries are identical to the production tool.

---

## What it does

- **Single-screen dashboard UI** — all 28 working sheets are hidden; the user navigates entirely through buttons and Form Controls on one DASHBOARD sheet
- **Supplier-based cost filtering** — a ComboBox drives dynamic shape-based displays showing any supplier's cost history, invoice status, and totals
- **Four cost capture streams** — automated main table, manual auxiliary table, hired equipment (with hire-period date logic), and contracts/subcontracts
- **Price benchmarking** — a maintained price list with per-supplier, per-material rates and multi-criteria lookups (supplier × material × hire type)
- **HGV / deliveries tracking** — combined delivery, collection, and inventory streams appended via Power Query into a pivot-driven logistics view with PV/BESS and WEEK/TOTAL toggle switches
- **Invoice status tracking** — conditional formatting flags paid / pending / overdue items at a glance
- **Category & principle-based analysis** — cost roll-ups by work principle, activity, and material/service category with dedicated chart sheets

## Architecture

```
DASHBOARD (only visible sheet)
   │  Form Controls (ComboBox → linked cell → VBA)
   │
   ├── Navigation macros (GoTo_*) unhide/activate working sheets on demand
   │
   ├── MAIN TABLE-1 AUTOMATED  ─┐
   ├── AUX. TABLE-2 MANUEL      ├─ 13 Power Query queries append into
   ├── TABLE-3 HIRED EQUIPMENT  │  MASTER LIST / HGV_Combined
   ├── CONTRACT TABLE-4 MANUAL ─┘
   │
   ├── PRICES  (multi-criteria INDEX/MATCH & SUMIF lookups)
   ├── SUPPLIERS (source list feeding ComboBox + data validation)
   └── CHART sheets (material/service, work-based, principle-based, PV vs BESS)
```

**Key mechanics:**

- 33 VBA subs: navigation, filter clearing, query refresh (`RefreshAllQueries`), toggle switches (`ToggleSwitch`, `ToggleWeekTotal` — animated shape-based PV/BESS and WEEK/TOTAL switches), and targeted filters (`FilterDiesel`, `FilterDrinkingWater`, `FilterEffluent`, ...)
- 13 Power Query (M) queries — all internal `$Workbook$` sources, no external dependencies
- Pivot caches driving the logistics dashboard
- Designed to survive a **OneDrive/AutoSave environment** — Form Controls instead of ActiveX throughout (ActiveX is unstable under AutoSave)

## Why Excel?

Because it's what a construction site actually has. No IT approval, no licences, no server. One `.xlsm` file on the site OneDrive gives the whole engineering team live cost visibility. The constraint was the point.

## Files

| Path | Contents |
|---|---|
| `Construction_Cost_Dashboard_DEMO.xlsm` | The full workbook (macros enabled) with sample data |
| `vba_modules/` | All VBA code exported as readable `.bas` / `.cls` files |

## Usage

1. Download the `.xlsm` and open in desktop Excel (macros must be enabled)
2. Everything starts from the DASHBOARD sheet — use the buttons
3. Working sheets unhide as you navigate; the `GoToDashboard` button returns and re-hides

## Author

**Mehmet Baran Akat** — Civil Engineer / Site Lead Engineer, UK renewable energy sector (BESS & solar PV).
Built on Sundays, refined on site.

- LinkedIn: [linkedin.com/in/mehmet-baran-akat-1ba0b6a9](https://linkedin.com/in/mehmet-baran-akat-1ba0b6a9)
