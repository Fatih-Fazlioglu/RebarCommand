# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build Commands

```bash
# Build (from project directory)
dotnet build

# Build x64 Debug (recommended for AutoCAD 2024)
dotnet build -c Debug -p:Platform=x64

# Build Release
dotnet build -c Release -p:Platform=x64
```

Output: `bin\x64\Debug\RebarCommand.dll`

## Project Overview

VB.NET AutoCAD 2024 plugin for automated rebar detailing in structural beam drawings. Reads beam reinforcement data from Excel and generates plan views with rebars, dimensions, and quantity takeoff tables (metraj).

## Architecture

### Core Modules

**ReinforcementBarAutomation.vb** - Main entry point
- `DONATI_PLAN` command: Generates complete rebar plan from Excel data
- Handles beam geometry collection, rebar drawing, duplicate sets, and web rebars
- Contains `BeamExcelRow` class for parsed Excel data

**PosManager.vb** - Rebar position (poz) block management
- `Manager` class: Tracks unique rebar positions, creates POZ_MARKER blocks
- Methods: `AddDuplicateRebar`, `AddDuplicateRebarWithAdetText`, `AddTiebarPos`, `AddStirrupPos`
- ADET field supports both integer and "nxm" format (e.g., "3x2" for grouped rebars)

**MetrajGenerator.vb** - Quantity takeoff table
- `METRAJ_HESAPLA` command: Creates bill of materials table from POZ_MARKER blocks
- Parses "nxm" ADET format as n*m for calculations
- Groups by position number, calculates weights by diameter

### Supporting Modules

| Module | Purpose |
|--------|---------|
| `GeometryCollector.vb` | Collects columns, beams, labels from drawing |
| `LayerManager.vb` | Layer creation and management |
| `DimensionUtils.vb` | Dimension and length label creation |
| `SectionsModule.vb` | Cross-section view generation |
| `StirrupModule.vb` | Stirrup drawing logic |
| `TieBarsHelper.vb` | Tie bar (çiroz) specification resolution |
| `MathUtils.vb` | Geometry utilities (Round5Up, line intersections) |

### Key Data Structures

**BeamExcelRow** (in ReinforcementBarAutomation.vb):
- Parsed from Excel: KNumber, dimensions, rebar quantities (QtyMontaj, QtyLower, QtyAddlLeft/Right)
- Phi values: PhiMontaj, PhiLower, PhiWeb, PhiAddlLeft/Right

**PosManager.Instance**:
- Key, PosNo, Mid, PhiMm, TotalLen, Yer, Adet (String), MiterCount, Miter1/2, StraightPart
- Flags: IsRebar, IsTie, IsStirrup, IsForMetraj

### Layer Conventions

- `LYR_REBAR1`: Original rebars
- `LYR_REBAR2`: Duplicate rebars
- `LYR_DIM`: Dimensions
- `poz`: Position markers

### Drawing Constants

- `DUP_OFFSET`: 80.0 - Vertical offset for duplicate set
- `DUP_SPACING`: 15.0 - Spacing between duplicate rebar lines
- `EPS`: 0.001 - Tolerance for geometric comparisons
- Miter angles: Left=225°, Right=315° (45° diagonal hooks)

## Dependencies

- AutoCAD 2024 .NET API (accoremgd, acdbmgd, acmgd)
- ExcelDataReader 3.7.0 - Excel file parsing
- .NET Framework 4.8
