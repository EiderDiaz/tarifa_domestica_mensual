# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

This is a Python desktop utility for processing **CFE (Comisión Federal de Electricidad)** monthly domestic electricity bills in Mexico. It extracts data from PDF receipts and populates an Excel quotation template for solar photovoltaic system proposals.

The entire application lives in a single file: `tarifa_domestica_mensual.py`, which exposes one public function `procesar_tarifa_domestica_mensual()` designed to be imported by a parent app (`app_maestro`), or run standalone via `__main__`.

## How It Works (Pipeline)

1. **PDF selection** — User picks a CFE receipt PDF via a `tkinter` file dialog.
2. **Validation** — Double verification: tarifa must be domestic (1, 1A–1F, DAC) and billing period must be monthly (< 45 days). Rejects bimestral receipts.
3. **Data extraction** — Uses `pdfplumber` + regex to parse: customer name, address, RPU (service number), tarifa, wire count, consumption/payment history (up to 12 periods), and cost breakdown (Suministro, IVA, DAP).
4. **State detection** — Matches Mexican state abbreviations from the address to a numeric code (used in energy calculations).
5. **Savings calculation** — Computes average real payments and subtracts base cost (suministro × IVA + DAP) to estimate savings.
6. **Excel population** — Writes extracted data into four sheets of a `.xlsm` template: `PROMEDIO DE CONSUMO`, `FORMATO DE COTIZACION`, `COTIZACIÓN`, `CALCULO DE ENERGIA`.
7. **Chart adjustment** — Uses `win32com.client` (COM automation) to open the saved Excel, read computed values from `RECUPERACION` sheet, and adjust Y-axis scaling on a chart.

## Dependencies

- `pdfplumber` — PDF text extraction
- `openpyxl` — Excel .xlsm read/write (with VBA macro preservation via `keep_vba=True`)
- `win32com.client` (`pywin32`) — Excel COM automation for chart/macro operations (Windows-only)
- `tkinter` — GUI file dialogs and message boxes (ships with Python)

There is no `requirements.txt`; install manually:
```
pip install pdfplumber openpyxl pywin32
```

## Running

```bash
python tarifa_domestica_mensual.py
```

This opens a file picker for a CFE PDF receipt. The Excel template path is hardcoded to `D:/SECOM/Cotizaciones José/COTIZACION SISTEMA FOTOVOLTAICO MENSUAL.xlsm`.

## Platform Notes

- The `win32com.client` dependency means chart adjustment only works on **Windows**. The `abrir_archivo()` helper handles cross-platform file opening, but the COM automation section will fail on macOS/Linux.
- The Excel template path and output directory (`D:/SECOM/Cotizaciones José`) are hardcoded Windows paths.

## Key Regex Patterns

These are critical to the PDF parsing and are sensitive to CFE receipt format changes:

- **Tarifa**: `TARIFA: *([1DAC]{1}[A-F]?)`
- **Periodo**: `PERIODO FACTURADO:\s*(\d{2}) ([A-Z]{3}) (\d{2})\s*-\s*(\d{2}) ([A-Z]{3}) (\d{2})`
- **RPU**: `NO\.? DE SERVICIO: *(\d+)`
- **Consumo**: `(\d{1,3}[,\d]*)\s+(\d{1,3}[,\d]*)\s+(\d{1,3}[,\d]*)`
- **Pago total**: `TOTAL A PAGAR:\s*\$?([\d,]+)`
- **Historial**: `del \d{2} [A-Z]{3} \d{2} al \d{2} [A-Z]{3} \d{2} (\d+) \$([\d,]+\.\d{2})`

## Code Conventions

- All code and comments are in **Spanish**.
- Helper functions and constants are defined **inside** `procesar_tarifa_domestica_mensual()` rather than at module level, to keep the module's public API as a single callable.
- Error handling uses `tkinter.messagebox` popups — there is no logging framework.
