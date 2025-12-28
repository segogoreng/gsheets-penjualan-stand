// Helper: Extract month abbreviation from targetSheetName (e.g., "des'25" → "des")
export function extractMonthAbbreviation(sheetName: string): string {
    const match = sheetName.match(/^([a-zA-Z]+)'/);
    if (match) {
        return match[1].toLowerCase();
    }
    throw new Error(`Cannot extract month from sheet name: ${sheetName}`);
}

// Helper: Build source sheet name from day and month (e.g., (1, "des") → "1des")
export function buildSourceSheetName(day: number, monthAbbr: string): string {
    return `${day}${monthAbbr}`;
}

// Helper: Convert day number to column index (Day 1 = 4 (D), Day 2 = 7 (G), etc.)
export function dayToColumnIndex(day: number): number {
    return 4 + (day - 1) * 3;
}

// Helper: Parse date range input ("1-3" → {start: 1, end: 3} or "5" → {start: 5, end: 5})
export function parseDateRange(
    input: string
): { start: number; end: number } | null {
    const trimmed = input.trim();

    // Handle single date: "5"
    if (/^\d+$/.test(trimmed)) {
        const day = parseInt(trimmed, 10);
        if (day >= 1 && day <= 31) {
            return { start: day, end: day };
        }
        return null;
    }

    // Handle range: "1-3"
    const rangeMatch = trimmed.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rangeMatch) {
        const start = parseInt(rangeMatch[1], 10);
        const end = parseInt(rangeMatch[2], 10);
        if (start >= 1 && end <= 31 && start <= end) {
            return { start, end };
        }
    }

    return null;
}
