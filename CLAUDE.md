# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project for sales/inventory tracking (penjualan stand) in Google Sheets. It uses clasp for local development and deployment, with Rollup and Babel for TypeScript transpilation.

## Commands

- `npm run build` - Transpile TypeScript to JavaScript using Rollup
- `npm run push` - Build and deploy to Google Apps Script
- `npm run open` - Open the Google Apps Script editor in browser
- `npm run lint` - Run ESLint with auto-fix
- `npm run format` - Format code with Prettier

## Architecture

### Build Pipeline

Source TypeScript (`src/index.ts` and `src/helpers.ts`) is transpiled via Rollup + Babel to `build/` directory. A custom Rollup plugin prevents tree-shaking since Google Apps Script requires all top-level functions to be available as menu handlers.

### Spreadsheet Integration

The script operates across two Google Spreadsheets:

- **Source spreadsheet**: Daily sales data (YUMBENTO PTC) - referenced by `sourceSpreadsheetId`
- **Target spreadsheet**: Monthly recap (Rekap yumbento PTC 2025) - always the active spreadsheet running the script

The `targetSheetName` variable at the top of index.ts must be updated when working on a new month.

### Custom Menu Functions

The `onOpen()` function creates a "Custom Menu" with these operations:

- **Clear Ranges**: Clears specific columns (D, G, J, M, P...) in rows 2-437
- **subtotal**: Sets formulas multiplying quantity columns by unit price (column C)
- **profit**: Sets formulas calculating profit (subtotal - quantity \* cost in column B)
- **isi formula sticker / Profit Sticker**: Copies values from row 285 to row 297 for sticker items
- **Proses Data Penjualan (Multi-Tanggal)**: Combined function that processes a date range:
    1. Prompts for date range (e.g., "1-3" or "5")
    2. For each date: adds new products (yellow rows), validates product names match, then copies sales data
    3. Stops and shows error if validation fails, otherwise shows summary at end
- **Masukkan data penjualan (Manual)**: Copies sales data from source to target spreadsheet via user prompts
- **masukkan barang baru (Manual)**: Processes yellow-highlighted rows to add new products

### Column Layout Pattern

The sheet uses a repeating 3-column pattern (31 groups for days of month):

- Column pattern: Quantity | Subtotal | Profit (e.g., D|E|F, G|H|I, J|K|L...)
- Column A: Product names
- Column B: Cost price
- Column C: Selling price
- Day to column mapping: Day 1 = D (index 4), Day 2 = G (index 7), formula: `4 + (day - 1) * 3`

## Interaction Preferences

When asked to ask clarifying questions, use multiple choice format instead of open-ended questions.

## Development Workflow

After making code changes, always run the following to verify everything works:

1. `npm run lint` - Check for linting errors
2. `npm run format` - Format source files using prettier
3. `npm run build` - Verify the build succeeds
4. `npm test` - Run unit tests
