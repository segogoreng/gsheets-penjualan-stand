# Performance Guidelines

When working with Google Sheets API, follow these patterns to ensure good performance:

## 1. Use Batch Read/Write Operations

Always use `getValues()` / `setValues()` for multiple cells instead of single-cell operations in loops.

```typescript
// BAD - Makes 100 API calls
for (let i = 1; i <= 100; i++) {
  const value = sheet.getRange(i, 1).getValue();
  sheet.getRange(i, 2).setValue(value * 2);
}

// GOOD - Makes 2 API calls
const values = sheet.getRange(1, 1, 100, 1).getValues();
const results = values.map(row => [row[0] * 2]);
sheet.getRange(1, 2, 100, 1).setValues(results);
```

## 2. Minimize API Calls

Cache sheet and range references outside of loops. Avoid calling `getRange()` repeatedly.

```typescript
// BAD - Repeated API calls inside loop
for (const day of days) {
  const sheet = ss.getSheetByName('January');
  const range = sheet.getRange(`A${day}:C${day}`);
  // ...
}

// GOOD - Cache references outside loop
const sheet = ss.getSheetByName('January');
const allData = sheet.getRange('A1:C31').getValues();
for (const day of days) {
  const rowData = allData[day - 1];
  // ...
}
```

## 3. In-Memory Processing

Read all data into arrays first, process in memory, then write back in a single operation.

```typescript
// GOOD - Read once, process in memory, write once
const data = sheet.getRange('A1:D100').getValues();

// Process in memory
const processed = data.map(row => {
  const [name, qty, price] = row;
  return [name, qty, price, qty * price]; // Add calculated column
});

// Write back once
sheet.getRange('A1:D100').setValues(processed);
```
