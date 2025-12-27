# Penjualan Stand

A Google Apps Script project for sales and inventory tracking in Google Sheets. Designed to manage daily sales data with batch processing, formula management, and data validation.

## Features

- **Batch Date Processing** - Process multiple dates at once with validation
- **Automatic Formula Management** - Set subtotal and profit formulas across columns
- **Data Validation** - Validates product names match between source and target sheets
- **New Product Detection** - Identifies and processes yellow-highlighted rows as new products
- **Custom Menu Integration** - All functions accessible via Google Sheets custom menu

## Prerequisites

- Node.js (v18+)
- npm
- Google account with access to Google Apps Script
- [clasp](https://github.com/google/clasp) CLI tool installed globally

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd penjualan-stand
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Login to clasp (if not already):
   ```bash
   clasp login
   ```

4. Create `.clasp.json` with your script ID:
   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID",
     "rootDir": "."
   }
   ```

## Scripts

| Command | Description |
|---------|-------------|
| `npm run build` | Transpile TypeScript to JavaScript |
| `npm run push` | Build and deploy to Google Apps Script |
| `npm run open` | Open Google Apps Script editor |
| `npm run lint` | Run ESLint with auto-fix |
| `npm run format` | Format code with Prettier |
| `npm test` | Run tests |
| `npm run test:watch` | Run tests in watch mode |

## Project Structure

```
penjualan-stand/
├── src/
│   ├── index.ts          # Main Apps Script entry point
│   ├── helpers.ts        # Utility functions
│   └── __tests__/        # Unit tests
├── build/                # Compiled output (generated)
├── rollup.config.js      # Build configuration
├── tsconfig.json         # TypeScript configuration
└── appsscript.json       # Apps Script manifest
```

## Spreadsheet Layout

The target spreadsheet uses a repeating 3-column pattern for each day:

| Column A | Column B | Column C | Day 1 (D-F) | Day 2 (G-I) | ... |
|----------|----------|----------|-------------|-------------|-----|
| Product  | Cost     | Price    | Qty\|Sub\|Profit | Qty\|Sub\|Profit | ... |

**Column formula:** `4 + (day - 1) * 3`

## Configuration

Update `targetSheetName` in `src/index.ts` when working on a new month:

```typescript
const targetSheetName = "des'25"; // Change this for each month
```

## Menu Functions

Once deployed, a "Custom Menu" appears in Google Sheets with:

- **Clear Ranges** - Clear quantity columns
- **subtotal** - Set subtotal formulas
- **profit** - Set profit formulas
- **Proses Data Penjualan (Multi-Tanggal)** - Batch process date range
- **Masukkan data penjualan (Manual)** - Manual data entry
- **masukkan barang baru (Manual)** - Add new products

## Development

1. Make changes in `src/`
2. Run `npm run lint` to check for issues
3. Run `npm test` to verify helpers work correctly
4. Run `npm run push` to deploy changes

## License

ISC
