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

Source TypeScript (`src/index.ts`) is transpiled via Rollup + Babel to `build/` directory. A custom Rollup plugin prevents tree-shaking since Google Apps Script requires all top-level functions to be available as menu handlers.

### Spreadsheet Integration

The script operates across two Google Spreadsheets:
- **Source spreadsheet**: Daily sales data (YUMBENTO PTC)
- **Target spreadsheet**: Monthly recap (Rekap yumbento PTC 2025)

The `targetSheetName` variable at the top of index.ts must be updated when working on a new month.

### Custom Menu Functions

The `onOpen()` function creates a "Custom Menu" with these operations:
- **Clear Ranges**: Clears specific columns (D, G, J, M, P...) in rows 2-437
- **subtotal**: Sets formulas multiplying quantity columns by unit price (column C)
- **profit**: Sets formulas calculating profit (subtotal - quantity * cost in column B)
- **isi formula sticker / Profit Sticker**: Copies values from row 285 to row 297 for sticker items
- **Masukkan data penjualan**: Copies sales data from source to target spreadsheet via user prompts
- **masukkan barang baru**: Processes yellow-highlighted rows to add new products

### Column Layout Pattern

The sheet uses a repeating 3-column pattern (31 groups for days of month):
- Column pattern: Quantity | Subtotal | Profit (e.g., D|E|F, G|H|I, J|K|L...)
- Column A: Product names
- Column B: Cost price
- Column C: Selling price
