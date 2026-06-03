---
name: power-bi-project-builder
description: Build a complete .pbip Power BI project from any data file (Excel/CSV). Creates the full folder structure — semantic model (TMDL), DAX measures, styled report page, and a custom SVG background — ready to open in Power BI Desktop. Invoke when the user uploads or points to a data file and asks to build a dashboard, report, or Power BI project.
tools: Write, Read, Bash, Glob
---

# Power BI Project Builder

## Goal
Turn any data file + user intent into a fully working `.pbip` project with semantic model (TMDL), measures, visuals, and a custom SVG background. Output opens in Power BI Desktop (May 2026+) with zero manual setup.

## Prerequisites
- Python + pandas + openpyxl: `pip install pandas openpyxl`
- User must provide: data file path, output folder, desired KPIs/charts
- Optional: brand colors, dark/light theme preference

---

## CRITICAL RULES — Never Break These

| Rule | Detail |
|------|--------|
| **No BOM** | Write every file with Python `open(p,'w',encoding='utf-8',newline='\n')`. Never use PowerShell `Set-Content`. Strip BOM at end with re-encode step. |
| **No model.bim** | Never create `model.bim` alongside TMDL `definition/` folder. PBI rejects both formats coexisting. Use TMDL only. |
| **No `//` comments in TMDL** | TMDL has NO comment syntax. Any `//` line causes a parse error. Empty files must be truly empty (blank line only). |
| **TMDL uses tabs only** | All indentation must be real tab characters (`\t`), never spaces. Write files with Python to guarantee this. |
| **measure names with spaces → single-quoted** | `measure 'Total Sales' = ...` not `measure Total Sales = ...`. Unquoted spaces cause "Invalid indentation" parse errors. |
| **`ref table` at root level** | In `model.tmdl`, `ref table` lines must be at zero indentation (root level), NOT indented under `model Model`. |
| **`formatString:` not `format:`** | Column date/time format property is `formatString: Short Date`, never `format: Short Date`. |
| **Reserved name** | Table name `Measures` is blocked. Always use `Calculations`. |
| **Proper UUIDs for lineageTag** | Always generate real UUIDs (`import uuid; str(uuid.uuid4())`). Short strings like `t1`, `c1` cause parse errors. |
| **Entity match** | `"Entity"` in every visual JSON must exactly match the TMDL `table` name (case-sensitive). |
| **sourceColumn exact** | TMDL `sourceColumn` must match the Excel header byte-for-byte. |
| **SVG title placement** | SVG labels must be at the TOP of each card/chart zone (just below accent bar). PBI visuals must be positioned BELOW the label so the text shows. |

---

## Workflow

### Step 1 — Read & Profile the Data
```python
import pandas as pd
sheets = pd.read_excel('file.xlsx', sheet_name=None)
for name, df in sheets.items():
    print(f'=== {name} | {df.shape} ===')
    print(df.dtypes)
    print(df.head(3).to_string())
```
Identify: text columns (→ axes/categories), numeric columns (→ measures), date columns, primary grouping field.

### Step 2 — Clarify with User
Ask once:
1. Which KPIs and chart types?
2. Output folder path?
3. Brand colors? (default light theme: bg `#F0F4FF`, header `#0D47A1`, accent `#1565C0`)
4. Dark or light background?

### Step 3 — Design the SVG Background

**Canvas**: 1500 × 844 px (standard PBI widescreen)

**Layout zones**:
- Header band `y=0–110`: gradient rect + title text + subtitle
- Section label `y≈135`: small caps label ("KEY METRICS")
- KPI card rects `y=148`, height=120: white rect + 4px colored top accent + label text at `y=168` (top of card)
- Section label `y≈300`: ("WORKFORCE ANALYTICS")
- Chart zones row 1 `y=310`, height=250: white rect + 4px accent + title text at `y=330`
- Section label `y≈582`
- Chart zones row 2 `y=592`, height=225: white rect + 4px accent + title text at `y=612`
- Footer `y=820–844`

**SVG label positioning rule**: Place `<text>` label at `cardTop + 20px`. Place PBI visual starting at `cardTop + 30px` so label is always visible above the visual data.

**Filename**: `{ProjectName}{8digits}.svg`
**Path**: `{Name}.Report/StaticResources/RegisteredResources/`

### Step 4 — Build Folder Structure

```
{Name}.pbip
{Name}.SemanticModel/
  .platform
  definition.pbism
  definition/
    database.tmdl
    model.tmdl
    relationships.tmdl
    tables/
      {DataTable}.tmdl
      Calculations.tmdl
{Name}.Report/
  .platform
  definition.pbir
  definition/
    report.json
    version.json
    pages/
      pages.json
      {pageId}/
        page.json
        visuals/
          {visualId}/
            visual.json
  StaticResources/
    RegisteredResources/
      {svg}
    SharedResources/
      BaseThemes/
        CY26SU02.json   ← copy from a working reference project
```

### Step 5 — Core Project Files (exact schemas required)

#### `{Name}.pbip`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
  "version": "1.0",
  "artifacts": [{"report": {"path": "{Name}.Report"}}],
  "settings": {"enableAutoRecovery": true}
}
```

#### `{Name}.Report/.platform`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
  "metadata": {"type": "Report", "displayName": "{Name}"},
  "config": {"version": "2.0", "logicalId": "<uuid>"}
}
```

#### `{Name}.SemanticModel/.platform`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
  "metadata": {"type": "SemanticModel", "displayName": "{Name}"},
  "config": {"version": "2.0", "logicalId": "<uuid>"}
}
```

#### `definition.pbir`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
  "version": "4.0",
  "datasetReference": {"byPath": {"path": "../{Name}.SemanticModel"}}
}
```

#### `definition.pbism`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
  "version": "4.2",
  "settings": {}
}
```

#### `definition/version.json`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
  "version": "2.0.0"
}
```

### Step 6 — TMDL Semantic Model Files

Write ALL `.tmdl` files using Python to guarantee tab indentation:

#### `database.tmdl`
```
database {Name}
	compatibilityLevel: 1702
```

#### `model.tmdl`
```
model Model
	culture: en-US
	defaultPowerBIDataSourceVersion: powerBI_V3
	sourceQueryCulture: en-US
	dataAccessOptions
		legacyRedirects
		returnErrorValuesAsNull

annotation PBI_QueryOrder = ["{DataTable}","Calculations"]

ref table {DataTable}
ref table Calculations
```
**Critical**: `ref table` and `annotation` lines are at ROOT level (zero indent), NOT under `model Model`.

#### `relationships.tmdl`
Must be a blank file (single newline). No comments, no text.

#### `{DataTable}.tmdl`
```
table {DataTable}
	lineageTag: <uuid>

	column ColumnName
		dataType: string
		lineageTag: <uuid>
		summarizeBy: none
		sourceColumn: ColumnName

	column NumericCol
		dataType: int64
		lineageTag: <uuid>
		summarizeBy: sum
		sourceColumn: NumericCol

	column DateCol
		dataType: dateTime
		formatString: Short Date
		lineageTag: <uuid>
		summarizeBy: none
		sourceColumn: DateCol

	partition {DataTable} = m
		mode: import
		source =
				let
				    Source = Excel.Workbook(File.Contents("C:\\path\\to\\file.xlsx"), null, true),
				    Sheet = Source{[Item="SheetName",Kind="Sheet"]}[Data],
				    PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),
				    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders,{...})
				in
				    ChangedTypes

	annotation PBI_ResultType = Table
```

**Column property rules**:
- Property order: `dataType`, `formatString` (dates only), `lineageTag`, `summarizeBy`, `sourceColumn`
- Use `formatString: Short Date` NOT `format: Short Date`
- All lineageTags must be real UUIDs

#### `Calculations.tmdl`
```
table Calculations
	lineageTag: <uuid>

	measure 'Total Records' = COUNTROWS({DataTable})
		formatString: #,0
		lineageTag: <uuid>

	measure 'My Rate' = DIVIDE([Numerator], [Denominator])
		formatString: 0.0%;-0.0%;0.0%
		lineageTag: <uuid>

	measure 'Avg Value' = AVERAGE({DataTable}[Column])
		formatString: $#,0
		lineageTag: <uuid>

	partition Calculations = m
		mode: import
		source =
				let
				    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),
				    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}}),
				    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column1"})
				in
				    #"Removed Columns"

	annotation PBI_ResultType = Table
```

**Measure name rules**:
- Names with spaces MUST be in single quotes: `measure 'My Measure' = ...`
- Names without spaces need no quotes: `measure TotalCount = ...`
- Measure properties (`formatString`, `lineageTag`) indented with 2 tabs

### Step 7 — Report JSON Files

#### `report.json`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json",
  "themeCollection": {
    "baseTheme": {
      "name": "CY26SU02",
      "reportVersionAtImport": {"visual": "2.6.0", "report": "3.1.0", "page": "2.3.0"},
      "type": "SharedResources"
    }
  },
  "resourcePackages": [
    {
      "name": "SharedResources",
      "type": "SharedResources",
      "items": [{"name": "CY26SU02", "path": "BaseThemes/CY26SU02.json", "type": "BaseTheme"}]
    },
    {
      "name": "RegisteredResources",
      "type": "RegisteredResources",
      "items": [{"name": "{SvgFilename}.svg", "path": "{SvgFilename}.svg", "type": "Image"}]
    }
  ],
  "settings": {
    "useStylableVisualContainerHeader": true,
    "exportDataMode": "AllowSummarized",
    "defaultDrillFilterOtherVisuals": true,
    "allowChangeFilterTypes": true,
    "useEnhancedTooltips": true,
    "useDefaultAggregateDisplayName": true
  }
}
```
**Note**: SVG resource type is `"Image"` not `"svg"`. Theme type is `"SharedResources"` not `2`.

#### `pages/pages.json`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
  "pageOrder": ["{pageId}"],
  "activePageName": "{pageId}"
}
```

#### `pages/{pageId}/page.json`
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
  "name": "{pageId}",
  "displayName": "Dashboard",
  "displayOption": "FitToPage",
  "height": 844,
  "width": 1500,
  "objects": {
    "background": [
      {
        "properties": {
          "image": {
            "image": {
              "name": {"expr": {"Literal": {"Value": "'{SvgFilename}.svg'"}}},
              "url": {
                "expr": {
                  "ResourcePackageItem": {
                    "PackageName": "RegisteredResources",
                    "PackageType": 1,
                    "ItemName": "{SvgFilename}.svg"
                  }
                }
              },
              "scaling": {"expr": {"Literal": {"Value": "'Fill'"}}}
            }
          },
          "transparency": {"expr": {"Literal": {"Value": "0D"}}}
        }
      }
    ]
  }
}
```
**Critical**: Background uses `objects.background[].properties.image.image` with `name/url/scaling` each wrapped in `expr`. NOT a flat `image` object.

### Step 8 — Visual JSON Files

Every `visual.json` MUST have `$schema` and use `position` object (not flat x/y/z).

**Schema**: `https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json`

#### KPI Card visual (`cardVisual` type, `Data` bucket)
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json",
  "name": "card1",
  "position": {"x": 30.0, "y": 178.0, "z": 1000, "height": 88.0, "width": 330.0, "tabOrder": 1000},
  "visual": {
    "visualType": "cardVisual",
    "query": {
      "queryState": {
        "Data": {
          "projections": [{
            "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Calculations"}}, "Property": "Total Headcount"}},
            "queryRef": "Calculations.Total Headcount",
            "nativeQueryRef": "Total Headcount"
          }]
        }
      },
      "sortDefinition": {
        "sort": [{"field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Calculations"}}, "Property": "Total Headcount"}}, "direction": "Descending"}],
        "isDefaultSort": true
      }
    },
    "objects": {
      "divider":    [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}, "selector": {"id": "default"}}],
      "label":      [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}, "selector": {"id": "default"}}],
      "fillCustom": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "layout":     [{"properties": {"backgroundShow": {"expr": {"Literal": {"Value": "false"}}}}, "selector": {"id": "default"}}],
      "outline":    [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}, "selector": {"id": "default"}}],
      "value":      [{"properties": {"fontSize": {"expr": {"Literal": {"Value": "28D"}}}, "horizontalAlignment": {"expr": {"Literal": {"Value": "'left'"}}}}, "selector": {"id": "default"}}]
    },
    "visualContainerObjects": {
      "title":        [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "background":   [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "border":       [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
    },
    "drillFilterOtherVisuals": true
  }
}
```

**Visual type rules**:
- KPI single value → `cardVisual` with `Data` bucket (NOT `card` with `Values`)
- Bar chart (horizontal) → `barChart` with `Category` + `Y` buckets
- Column chart (vertical) → `clusteredColumnChart` with `Category` + `Y` buckets
- Line chart → `lineChart` with `Category` + `Y` buckets

#### Bar/Column chart visual
```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json",
  "name": "chart1",
  "position": {"x": 30.0, "y": 338.0, "z": 2000, "height": 218.0, "width": 440.0, "tabOrder": 2000},
  "visual": {
    "visualType": "barChart",
    "query": {
      "queryState": {
        "Category": {
          "projections": [{
            "field": {"Column": {"Expression": {"SourceRef": {"Entity": "TableName"}}, "Property": "ColumnName"}},
            "queryRef": "TableName.ColumnName",
            "nativeQueryRef": "ColumnName",
            "active": true
          }]
        },
        "Y": {
          "projections": [{
            "field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Calculations"}}, "Property": "MeasureName"}},
            "queryRef": "Calculations.MeasureName",
            "nativeQueryRef": "MeasureName"
          }]
        }
      },
      "sortDefinition": {
        "sort": [{"field": {"Measure": {"Expression": {"SourceRef": {"Entity": "Calculations"}}, "Property": "MeasureName"}}, "direction": "Descending"}],
        "isDefaultSort": true
      }
    },
    "visualContainerObjects": {
      "title":        [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "background":   [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "border":       [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}],
      "visualHeader": [{"properties": {"show": {"expr": {"Literal": {"Value": "false"}}}}}]
    },
    "drillFilterOtherVisuals": true
  }
}
```

### Step 9 — SharedResources Base Theme

Copy `CY26SU02.json` from any working reference PBIP project into:
`{Name}.Report/StaticResources/SharedResources/BaseThemes/CY26SU02.json`

If no reference project exists, use pbi-cli to export one, or omit `SharedResources` from `report.json` entirely and use a plain theme name.

### Step 10 — Re-encode All Files (Run Last, Always)
```python
import os
root = 'path/to/project'
for dirpath, _, files in os.walk(root):
    for f in files:
        if f.endswith(('.json','.tmdl','.pbir','.pbism','.pbip','.svg','.bim')):
            p = os.path.join(dirpath, f)
            c = open(p, encoding='utf-8-sig').read()
            open(p, 'w', encoding='utf-8', newline='\n').write(c)
```

---

## SVG Design Guidelines

**Title positioning** (critical for visual overlay):
- SVG labels must appear at the TOP of each zone, just after the accent bar
- Formula: `label_y = zone_top + 20`
- Formula: `visual_y = zone_top + 30` (PBI visual starts 30px below zone top)
- This ensures the SVG label text is NOT covered by the PBI visual

**Card zone layout**:
```
zone_top=148 ──── 4px accent bar
y=168         ──── SVG label text (TOTAL HEADCOUNT)
y=178         ──── PBI cardVisual starts here  
zone_bottom=268
```

**Chart zone layout**:
```
zone_top=310 ──── 4px accent bar  
y=330        ──── SVG title text (Headcount by Status)
y=338        ──── PBI chart visual starts here
zone_bottom=560
```

**Hide PBI titles**: Always set `visualContainerObjects.title.show = false` on every visual since SVG provides the titles.

---

## Validation Checklist

Before opening in Power BI Desktop, verify:

- [ ] No UTF-8 BOM in any file (re-encode step run)
- [ ] No `model.bim` file present alongside TMDL `definition/` folder
- [ ] All `.platform` files have `$schema` property with correct URL
- [ ] `definition.pbir` has `version: "4.0"` and `$schema`
- [ ] `definition.pbism` has `version: "4.2"` and `$schema`
- [ ] `version.json` has `$schema` and `version: "2.0.0"`
- [ ] `pages.json` has `$schema`
- [ ] `page.json` has `$schema` and correct background `objects` structure
- [ ] Every `visual.json` has `$schema` and uses `position` object (not flat x/y/z)
- [ ] KPI visuals use `cardVisual` type with `Data` bucket (not `card` + `Values`)
- [ ] `ref table` lines in `model.tmdl` are at ROOT level (zero indent)
- [ ] `relationships.tmdl` is blank — no comments
- [ ] All TMDL files use tabs (not spaces) for indentation
- [ ] Measure names with spaces are single-quoted in TMDL
- [ ] All `lineageTag` values are real UUIDs
- [ ] Column date format is `formatString:` not `format:`
- [ ] No table named `Measures` — only `Calculations`
- [ ] SVG item type in `report.json` is `"Image"` not `"svg"`
- [ ] Theme `type` in `report.json` is `"SharedResources"` not integer `2`
- [ ] `CY26SU02.json` base theme file present in `SharedResources/BaseThemes/`
- [ ] All `sourceColumn` values match data file headers exactly
- [ ] PBI visual `y` positions start BELOW SVG label text (at least 10px below label)
- [ ] `title: show = false` set in `visualContainerObjects` on all visuals

---

## Common Errors & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| `Expected '$schema' property in .platform` | Missing `$schema` in `.platform` files | Add schema URL to both `.platform` files |
| `Missing required artifact 'model.bim'` | Old `definition.pbism` format (version 1.0) | Set `version: "4.2"` and add `$schema` to `definition.pbism` |
| `Cannot have both TMDL and TMSL` | `model.bim` present alongside `definition/` folder | Delete `model.bim` — keep TMDL only |
| `ref table — Unexpected line type: ReferenceObject` | `ref table` indented under `model Model` | Move `ref table` lines to root level (zero indent) |
| `// No relationships — Unexpected line type: Other` | TMDL does not support `//` comments | Make `relationships.tmdl` a blank file |
| `Invalid indentation — measure Total Headcount` | Measure name has spaces without single quotes | Use `measure 'Total Headcount' = ...` |
| `Invalid indentation — measure [Total Headcount]` | Brackets not valid in TMDL measure declaration | Remove brackets; use `measure 'Name' = ...` |
| `Invalid indentation` (generic, still tabs) | Wrong `format:` property on date columns | Change `format:` to `formatString:` |
| No visuals show in report | Wrong visual format: flat x/y/z, wrong type/bucket | Use `position` object, `cardVisual`+`Data`, charts need `Category`+`Y` |
| SVG background missing | Wrong `page.json` background format | Use `objects.background[].properties.image.image` with `expr` wrappers |
| SVG labels hidden by visuals | PBI visual positioned over SVG text | Move PBI visual `y` to `zone_top + 30`, label at `zone_top + 20` |


---

## Complete Working Example — Full Build Script

This is the **exact verified Python script** that produces a working PBIP project.
Set the 5 variables at the top, run it — opens in Power BI Desktop with zero errors.

All TMDL is written via Python `open(..., newline='\n')` to guarantee real tab characters.

```python
import os, json, uuid, shutil

# ── Configure these 5 values ──────────────────────────────────────────────────
PROJECT_NAME = 'SalesDashboard'      # No spaces
DATA_FILE    = 'C:/Data/Sales.xlsx'  # Absolute path to Excel
SHEET_NAME   = 'SalesData'           # Sheet name in the Excel file
OUTPUT_DIR   = 'C:/Projects'         # Parent folder; project created inside
# Copy CY26SU02.json from any existing working PBIP project:
THEME_SRC    = 'C:/Ref.Report/StaticResources/SharedResources/BaseThemes/CY26SU02.json'
# ─────────────────────────────────────────────────────────────────────────────

ROOT    = f'{OUTPUT_DIR}/{PROJECT_NAME}'
SM      = f'{ROOT}/{PROJECT_NAME}.SemanticModel'
REP     = f'{ROOT}/{PROJECT_NAME}.Report'
DSM     = f'{SM}/definition'
DREP    = f'{REP}/definition'
PAGE_ID = 'ReportSection'
SVG     = f'{PROJECT_NAME}00000001.svg'
VS      = 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json'
PS      = 'https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json'


def w(p, c):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, 'w', encoding='utf-8', newline='\n').write(c)

def j(p, o): w(p, json.dumps(o, indent=2))
def uid(): return str(uuid.uuid4())
def lit(v): return {'expr': {'Literal': {'Value': v}}}
def hide():
    return {'title':        [{'properties': {'show': lit('false')}}],
            'background':   [{'properties': {'show': lit('false')}}],
            'border':       [{'properties': {'show': lit('false')}}],
            'visualHeader': [{'properties': {'show': lit('false')}}]}


# Folders
for d in [f'{DSM}/tables', f'{DREP}/pages/{PAGE_ID}/visuals',
          f'{REP}/StaticResources/RegisteredResources',
          f'{REP}/StaticResources/SharedResources/BaseThemes']:
    os.makedirs(d, exist_ok=True)


# ── .pbip ─────────────────────────────────────────────────────────────────────
j(f'{ROOT}/{PROJECT_NAME}.pbip', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json',
    'version': '1.0',
    'artifacts': [{'report': {'path': f'{PROJECT_NAME}.Report'}}],
    'settings': {'enableAutoRecovery': True}
})

# ── .platform (both need $schema + correct metadata.type) ─────────────────────
j(f'{REP}/.platform', {'$schema': PS, 'metadata': {'type': 'Report',        'displayName': PROJECT_NAME}, 'config': {'version': '2.0', 'logicalId': uid()}})
j(f'{SM}/.platform',  {'$schema': PS, 'metadata': {'type': 'SemanticModel', 'displayName': PROJECT_NAME}, 'config': {'version': '2.0', 'logicalId': uid()}})

# ── definition.pbir  version MUST be "4.0" ────────────────────────────────────
j(f'{REP}/definition.pbir', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json',
    'version': '4.0',
    'datasetReference': {'byPath': {'path': f'../{PROJECT_NAME}.SemanticModel'}}
})

# ── definition.pbism  version MUST be "4.2" ───────────────────────────────────
j(f'{SM}/definition.pbism', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json',
    'version': '4.2', 'settings': {}
})

# ── version.json ──────────────────────────────────────────────────────────────
j(f'{DREP}/version.json', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json',
    'version': '2.0.0'
})

# ── TMDL: database ───────────────────────────────────────────────────────────
w(f'{DSM}/database.tmdl', f'database {PROJECT_NAME}\n\tcompatibilityLevel: 1702\n')

# ── TMDL: model ───────────────────────────────────────────────────────────────
# CRITICAL: ref table lines at ROOT level (zero indent) NOT under model Model
w(f'{DSM}/model.tmdl',
  'model Model\n'
  '\tculture: en-US\n'
  '\tdefaultPowerBIDataSourceVersion: powerBI_V3\n'
  '\tsourceQueryCulture: en-US\n'
  '\tdataAccessOptions\n'
  '\t\tlegacyRedirects\n'
  '\t\treturnErrorValuesAsNull\n'
  '\n'
  f'annotation PBI_QueryOrder = ["{SHEET_NAME}","Calculations"]\n'
  '\n'
  f'ref table {SHEET_NAME}\n'
  'ref table Calculations\n')

# ── TMDL: relationships  (BLANK — no comments, TMDL has no comment syntax) ────
w(f'{DSM}/relationships.tmdl', '\n')

# ── TMDL: data table ──────────────────────────────────────────────────────────
# Property order: dataType > formatString > lineageTag > summarizeBy > sourceColumn
# Dates use "formatString: Short Date" NOT "format: Short Date"
# All lineageTag values must be real UUIDs
w(f'{DSM}/tables/{SHEET_NAME}.tmdl',
  f'table {SHEET_NAME}\n'
  f'\tlineageTag: {uid()}\n\n'
  '\tcolumn ID\n'
  f'\t\tdataType: int64\n\t\tlineageTag: {uid()}\n\t\tsummarizeBy: none\n\t\tsourceColumn: ID\n\n'
  '\tcolumn Name\n'
  f'\t\tdataType: string\n\t\tlineageTag: {uid()}\n\t\tsummarizeBy: none\n\t\tsourceColumn: Name\n\n'
  '\tcolumn Date\n'
  f'\t\tdataType: dateTime\n\t\tformatString: Short Date\n\t\tlineageTag: {uid()}\n\t\tsummarizeBy: none\n\t\tsourceColumn: Date\n\n'
  '\tcolumn Value\n'
  f'\t\tdataType: int64\n\t\tlineageTag: {uid()}\n\t\tsummarizeBy: sum\n\t\tsourceColumn: Value\n\n'
  f'\tpartition {SHEET_NAME} = m\n'
  '\t\tmode: import\n'
  '\t\tsource =\n'
  '\t\t\t\tlet\n'
  f'\t\t\t\t    Source = Excel.Workbook(File.Contents("{DATA_FILE}"), null, true),\n'
  f'\t\t\t\t    Sheet = Source{{[Item="{SHEET_NAME}",Kind="Sheet"]}}[Data],\n'
  '\t\t\t\t    PromotedHeaders = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true]),\n'
  '\t\t\t\t    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders,'
  '{{"ID",Int64.Type},{"Name",type text},{"Date",type datetime},{"Value",Int64.Type}})\n'
  '\t\t\t\tin\n'
  '\t\t\t\t    ChangedTypes\n\n'
  '\tannotation PBI_ResultType = Table\n')

# ── TMDL: Calculations ────────────────────────────────────────────────────────
# Measure names with spaces MUST be single-quoted: measure 'My Name' = ...
w(f'{DSM}/tables/Calculations.tmdl',
  f'table Calculations\n'
  f'\tlineageTag: {uid()}\n\n'
  f"\tmeasure 'Total Records' = COUNTROWS({SHEET_NAME})\n"
  f'\t\tformatString: #,0\n\t\tlineageTag: {uid()}\n\n'
  f"\tmeasure 'Total Value' = SUM({SHEET_NAME}[Value])\n"
  f'\t\tformatString: #,0\n\t\tlineageTag: {uid()}\n\n'
  f"\tmeasure 'Avg Value' = AVERAGE({SHEET_NAME}[Value])\n"
  f'\t\tformatString: #,0.0\n\t\tlineageTag: {uid()}\n\n'
  '\tpartition Calculations = m\n'
  '\t\tmode: import\n'
  '\t\tsource =\n'
  '\t\t\t\tlet\n'
  '\t\t\t\t    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i44FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Column1 = _t]),\n'
  '\t\t\t\t    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Column1", type text}}),\n'
  '\t\t\t\t    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column1"})\n'
  '\t\t\t\tin\n'
  '\t\t\t\t    #"Removed Columns"\n\n'
  '\tannotation PBI_ResultType = Table\n')

# ── SVG background ────────────────────────────────────────────────────────────
# label_y = zone_top + 20   PBI visual_y = zone_top + 30
# Cards: zone_top=148  label=168  visual=178
# Charts: zone_top=310  label=330  visual=338
svg_lines = [
    '<svg xmlns="http://www.w3.org/2000/svg" width="1500" height="844">',
    '  <defs>',
    '    <linearGradient id="hg" x1="0" y1="0" x2="1" y2="0">',
    '      <stop offset="0%" stop-color="#0D47A1"/><stop offset="100%" stop-color="#1565C0"/>',
    '    </linearGradient>',
    '    <linearGradient id="bg" x1="0" y1="0" x2="0" y2="1">',
    '      <stop offset="0%" stop-color="#F0F4FF"/><stop offset="100%" stop-color="#E8EEF9"/>',
    '    </linearGradient>',
    '  </defs>',
    '  <rect width="1500" height="844" fill="url(#bg)"/>',
    '  <rect x="0" y="0" width="1500" height="110" fill="url(#hg)"/>',
    '  <circle cx="1440" cy="20" r="100" fill="#1976D2" opacity="0.15"/>',
    f'  <text x="40" y="46" font-family="Segoe UI,Arial,sans-serif" font-size="28" font-weight="700" fill="#FFFFFF">{PROJECT_NAME}</text>',
    '  <text x="40" y="76" font-family="Segoe UI,Arial,sans-serif" font-size="13" fill="#BBDEFB">Analytics Dashboard</text>',
    '  <!-- KEY METRICS section -->',
    '  <text x="40" y="135" font-family="Segoe UI,Arial,sans-serif" font-size="11" font-weight="600" fill="#5C6BC0" letter-spacing="1.5">KEY METRICS</text>',
    '  <rect x="30" y="148" width="450" height="120" rx="8" fill="#FFFFFF"/>',
    '  <rect x="30" y="148" width="450" height="4" rx="2" fill="#1565C0"/>',
    '  <text x="50" y="168" font-family="Segoe UI,Arial,sans-serif" font-size="11" font-weight="600" fill="#78909C">TOTAL RECORDS</text>',
    '  <rect x="530" y="148" width="450" height="120" rx="8" fill="#FFFFFF"/>',
    '  <rect x="530" y="148" width="450" height="4" rx="2" fill="#00897B"/>',
    '  <text x="550" y="168" font-family="Segoe UI,Arial,sans-serif" font-size="11" font-weight="600" fill="#78909C">TOTAL VALUE</text>',
    '  <rect x="1030" y="148" width="440" height="120" rx="8" fill="#FFFFFF"/>',
    '  <rect x="1030" y="148" width="440" height="4" rx="2" fill="#F57C00"/>',
    '  <text x="1050" y="168" font-family="Segoe UI,Arial,sans-serif" font-size="11" font-weight="600" fill="#78909C">AVG VALUE</text>',
    '  <!-- ANALYTICS section -->',
    '  <text x="40" y="300" font-family="Segoe UI,Arial,sans-serif" font-size="11" font-weight="600" fill="#5C6BC0" letter-spacing="1.5">ANALYTICS</text>',
    '  <rect x="30" y="310" width="700" height="250" rx="8" fill="#FFFFFF"/>',
    '  <rect x="30" y="310" width="700" height="4" rx="2" fill="#1565C0"/>',
    '  <text x="50" y="330" font-family="Segoe UI,Arial,sans-serif" font-size="12" font-weight="600" fill="#37474F">Records by Name</text>',
    '  <rect x="780" y="310" width="690" height="250" rx="8" fill="#FFFFFF"/>',
    '  <rect x="780" y="310" width="690" height="4" rx="2" fill="#F57C00"/>',
    '  <text x="800" y="330" font-family="Segoe UI,Arial,sans-serif" font-size="12" font-weight="600" fill="#37474F">Value by Name</text>',
    '  <rect x="0" y="820" width="1500" height="24" fill="#E3EAF6"/>',
    f'  <text x="40" y="836" font-family="Segoe UI,Arial,sans-serif" font-size="10" fill="#90A4AE">{PROJECT_NAME} Dashboard</text>',
    '</svg>',
]
w(f'{REP}/StaticResources/RegisteredResources/{SVG}', '\n'.join(svg_lines))

# ── report.json ───────────────────────────────────────────────────────────────
# SVG item type = "Image" (NOT "svg")  |  theme type = "SharedResources" (NOT integer 2)
j(f'{DREP}/report.json', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.3.0/schema.json',
    'themeCollection': {'baseTheme': {'name': 'CY26SU02', 'reportVersionAtImport': {'visual': '2.6.0', 'report': '3.1.0', 'page': '2.3.0'}, 'type': 'SharedResources'}},
    'resourcePackages': [
        {'name': 'SharedResources', 'type': 'SharedResources', 'items': [{'name': 'CY26SU02', 'path': 'BaseThemes/CY26SU02.json', 'type': 'BaseTheme'}]},
        {'name': 'RegisteredResources', 'type': 'RegisteredResources', 'items': [{'name': SVG, 'path': SVG, 'type': 'Image'}]}],
    'settings': {'useStylableVisualContainerHeader': True, 'exportDataMode': 'AllowSummarized',
                 'defaultDrillFilterOtherVisuals': True, 'allowChangeFilterTypes': True,
                 'useEnhancedTooltips': True, 'useDefaultAggregateDisplayName': True}})

# ── pages.json ────────────────────────────────────────────────────────────────
j(f'{DREP}/pages/pages.json', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json',
    'pageOrder': [PAGE_ID], 'activePageName': PAGE_ID})

# ── page.json ─────────────────────────────────────────────────────────────────
# Background: objects.background[0].properties.image.image with expr wrappers
j(f'{DREP}/pages/{PAGE_ID}/page.json', {
    '$schema': 'https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json',
    'name': PAGE_ID, 'displayName': 'Dashboard', 'displayOption': 'FitToPage',
    'height': 844, 'width': 1500,
    'objects': {'background': [{'properties': {
        'image': {'image': {
            'name':    {'expr': {'Literal': {'Value': f"'{ SVG}'"}}},
            'url':     {'expr': {'ResourcePackageItem': {'PackageName': 'RegisteredResources', 'PackageType': 1, 'ItemName': SVG}}},
            'scaling': {'expr': {'Literal': {'Value': "'Fill'"}}}}},
        'transparency': {'expr': {'Literal': {'Value': '0D'}}}}}]}})

# ── Visual helpers ────────────────────────────────────────────────────────────
def write_v(vid, vj):
    d = f'{DREP}/pages/{PAGE_ID}/visuals/{vid}'
    os.makedirs(d, exist_ok=True)
    j(f'{d}/visual.json', vj)

def make_card(vid, x, y, cw, ch, z, mname):
    qr = f'Calculations.{mname}'
    return {'$schema': VS, 'name': vid,
            'position': {'x': float(x), 'y': float(y), 'z': z, 'height': float(ch), 'width': float(cw), 'tabOrder': z},
            'visual': {
                'visualType': 'cardVisual',   # NOT 'card'
                'query': {
                    'queryState': {'Data': {'projections': [{  # 'Data' bucket, NOT 'Values'
                        'field': {'Measure': {'Expression': {'SourceRef': {'Entity': 'Calculations'}}, 'Property': mname}},
                        'queryRef': qr, 'nativeQueryRef': mname}]}},
                    'sortDefinition': {'sort': [{'field': {'Measure': {'Expression': {'SourceRef': {'Entity': 'Calculations'}}, 'Property': mname}}, 'direction': 'Descending'}], 'isDefaultSort': True}},
                'objects': {
                    'divider':    [{'properties': {'show': lit('false')}, 'selector': {'id': 'default'}}],
                    'label':      [{'properties': {'show': lit('false')}, 'selector': {'id': 'default'}}],
                    'fillCustom': [{'properties': {'show': lit('false')}}],
                    'layout':     [{'properties': {'backgroundShow': lit('false')}, 'selector': {'id': 'default'}}],
                    'outline':    [{'properties': {'show': lit('false')}, 'selector': {'id': 'default'}}],
                    'value':      [{'properties': {'fontSize': lit('28D'), 'horizontalAlignment': lit("'left'")}, 'selector': {'id': 'default'}}]},
                'visualContainerObjects': hide(), 'drillFilterOtherVisuals': True}}

def make_chart(vid, x, y, cw, ch, z, vtype, cent, ccol, mname):
    return {'$schema': VS, 'name': vid,
            'position': {'x': float(x), 'y': float(y), 'z': z, 'height': float(ch), 'width': float(cw), 'tabOrder': z},
            'visual': {
                'visualType': vtype,  # 'barChart' | 'clusteredColumnChart' | 'lineChart'
                'query': {
                    'queryState': {
                        'Category': {'projections': [{'field': {'Column': {'Expression': {'SourceRef': {'Entity': cent}}, 'Property': ccol}}, 'queryRef': f'{cent}.{ccol}', 'nativeQueryRef': ccol, 'active': True}]},
                        'Y':        {'projections': [{'field': {'Measure': {'Expression': {'SourceRef': {'Entity': 'Calculations'}}, 'Property': mname}}, 'queryRef': f'Calculations.{mname}', 'nativeQueryRef': mname}]}},
                    'sortDefinition': {'sort': [{'field': {'Measure': {'Expression': {'SourceRef': {'Entity': 'Calculations'}}, 'Property': mname}}, 'direction': 'Descending'}], 'isDefaultSort': True}},
                'visualContainerObjects': hide(), 'drillFilterOtherVisuals': True}}

# Visuals — cards at y=178 (zone_top=148+30), charts at y=338 (zone_top=310+28)
write_v('card1',  make_card('card1',   30, 178, 450, 88, 1000, 'Total Records'))
write_v('card2',  make_card('card2',  530, 178, 450, 88, 2000, 'Total Value'))
write_v('card3',  make_card('card3', 1030, 178, 440, 88, 3000, 'Avg Value'))
write_v('chart1', make_chart('chart1',  30, 338, 700, 218, 4000, 'barChart',            SHEET_NAME, 'Name', 'Total Records'))
write_v('chart2', make_chart('chart2', 780, 338, 690, 218, 5000, 'clusteredColumnChart', SHEET_NAME, 'Name', 'Total Value'))

# ── Copy base theme ───────────────────────────────────────────────────────────
if os.path.exists(THEME_SRC):
    shutil.copy2(THEME_SRC, f'{REP}/StaticResources/SharedResources/BaseThemes/CY26SU02.json')
else:
    print('WARNING: Set THEME_SRC to CY26SU02.json from any existing PBIP project.')

# ── Strip BOM — ALWAYS run last ───────────────────────────────────────────────
for dp, _, fs in os.walk(ROOT):
    for f in fs:
        if f.endswith(('.json','.tmdl','.pbir','.pbism','.pbip','.svg')):
            p = os.path.join(dp, f)
            c = open(p, encoding='utf-8-sig').read()
            open(p, 'w', encoding='utf-8', newline='\n').write(c)

print(f'Done. Open: {ROOT}/{PROJECT_NAME}.pbip')
```
