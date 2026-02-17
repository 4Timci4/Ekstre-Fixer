# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Ekstre Fixer is a Windows desktop Electron app that processes Turkish bank statement Excel files. It normalizes, validates, and reformats bank extracts (ekstre) with proper balance calculations. The user is non-technical and communicates in Turkish — explain what/why in plain Turkish, skip code details unless asked.

## Commands

```bash
npm start          # Run app in dev mode (hot reload enabled)
```

## Architecture

Monolithic Electron app with two files handling all logic:

- **main.js** — Electron main process. Creates window (480x620), handles IPC (file dialogs, folder pickers, message boxes via `select-files`, `save-file`, `select-directory`, `select-folder`, `open-folder`, `show-message`). Uses `nodeIntegration: true` / `contextIsolation: false` (no preload script).

- **renderer.js** (~790 lines) — All business logic and UI state. Two processing modes:
  - **Genel Ekstre (Tab 1):** Processes all records, aggregates collection transactions by date, calculates running balance.
  - **Tarihli Ekstre (Tab 2):** Requires start date, calculates devir (carryover) from prior transactions, filters from start date onward.

- **index.html** — Dark-themed UI with Turkish language, persistent source/output folder settings in localStorage.

## Data Processing Pipeline

1. Read Excel with `xlsx` library (header at row 4, data from row 5)
2. Extract metadata (company name, debtor name) and validate confirmation row
3. Normalize Turkish number format (`1.234,56` → `1234.56`) and dates
4. Split amounts into positive/negative columns
5. Aggregate by transaction type (Tahsilat, GeriDevir, Senet, Çek)
6. Check for unmatched balances (EslenmemisTahsilat, EslenmemisCek, EslenmemisSenet)
7. Write formatted output with `exceljs` — balance column uses Excel formulas

## Key Dependencies

- `xlsx` — reading/parsing input Excel files
- `exceljs` — writing formatted Excel output with formulas and styling
- `dayjs` — date parsing and formatting (DD.MM.YYYY / DD/MM/YYYY)
- `electron-reload` — dev-only hot reload

## Output Format

Columns: Fatura Tarihi | Fatura Vadesi | Fatura No | İşlem Türü | İşlem Tarihi | İşlem Tutarı (+) | İşlem Tutarı (-) | Bakiye

Balance formula: `Previous Bakiye + Amount(+) + Amount(-)`. Number format: `#,##0.00`.

## Test Data

`Real-Examples/` contains sample input Excel files for manual testing (devir, genel ekstre, real company statements).
