import { describe, it, expect } from 'vitest';
import {
    extractMonthAbbreviation,
    buildSourceSheetName,
    dayToColumnIndex,
    parseDateRange,
} from '../helpers';

describe('extractMonthAbbreviation', () => {
    it('extracts month from standard format', () => {
        expect(extractMonthAbbreviation("des'25")).toBe('des');
        expect(extractMonthAbbreviation("jan'24")).toBe('jan');
        expect(extractMonthAbbreviation("agt'25")).toBe('agt');
    });

    it('converts to lowercase', () => {
        expect(extractMonthAbbreviation("JAN'24")).toBe('jan');
        expect(extractMonthAbbreviation("DES'25")).toBe('des');
    });

    it('throws error for invalid format', () => {
        expect(() => extractMonthAbbreviation('invalid')).toThrow(
            'Cannot extract month from sheet name: invalid'
        );
        expect(() => extractMonthAbbreviation("'25")).toThrow();
        expect(() => extractMonthAbbreviation('123')).toThrow();
    });
});

describe('buildSourceSheetName', () => {
    it('builds sheet name from day and month', () => {
        expect(buildSourceSheetName(1, 'des')).toBe('1des');
        expect(buildSourceSheetName(15, 'jan')).toBe('15jan');
        expect(buildSourceSheetName(31, 'agt')).toBe('31agt');
    });
});

describe('dayToColumnIndex', () => {
    it('converts day 1 to column D (index 4)', () => {
        expect(dayToColumnIndex(1)).toBe(4);
    });

    it('converts day 2 to column G (index 7)', () => {
        expect(dayToColumnIndex(2)).toBe(7);
    });

    it('converts day 3 to column J (index 10)', () => {
        expect(dayToColumnIndex(3)).toBe(10);
    });

    it('follows the pattern: 4 + (day - 1) * 3', () => {
        expect(dayToColumnIndex(10)).toBe(4 + 9 * 3); // 31
        expect(dayToColumnIndex(31)).toBe(4 + 30 * 3); // 94
    });
});

describe('parseDateRange', () => {
    describe('single date input', () => {
        it('parses single digit', () => {
            expect(parseDateRange('5')).toEqual({ start: 5, end: 5 });
            expect(parseDateRange('1')).toEqual({ start: 1, end: 1 });
        });

        it('parses double digit', () => {
            expect(parseDateRange('15')).toEqual({ start: 15, end: 15 });
            expect(parseDateRange('31')).toEqual({ start: 31, end: 31 });
        });

        it('trims whitespace', () => {
            expect(parseDateRange('  5  ')).toEqual({ start: 5, end: 5 });
        });

        it('returns null for day 0', () => {
            expect(parseDateRange('0')).toBeNull();
        });

        it('returns null for day > 31', () => {
            expect(parseDateRange('32')).toBeNull();
        });
    });

    describe('range input', () => {
        it('parses range format', () => {
            expect(parseDateRange('1-3')).toEqual({ start: 1, end: 3 });
            expect(parseDateRange('10-20')).toEqual({ start: 10, end: 20 });
        });

        it('handles spaces around dash', () => {
            expect(parseDateRange('1 - 3')).toEqual({ start: 1, end: 3 });
            expect(parseDateRange('1  -  3')).toEqual({ start: 1, end: 3 });
        });

        it('returns null for invalid range (start > end)', () => {
            expect(parseDateRange('5-3')).toBeNull();
        });

        it('returns null for out of bounds range', () => {
            expect(parseDateRange('0-5')).toBeNull();
            expect(parseDateRange('1-32')).toBeNull();
        });

        it('allows same start and end', () => {
            expect(parseDateRange('5-5')).toEqual({ start: 5, end: 5 });
        });
    });

    describe('invalid input', () => {
        it('returns null for non-numeric input', () => {
            expect(parseDateRange('abc')).toBeNull();
            expect(parseDateRange('one-three')).toBeNull();
        });

        it('returns null for empty string', () => {
            expect(parseDateRange('')).toBeNull();
            expect(parseDateRange('   ')).toBeNull();
        });

        it('returns null for malformed range', () => {
            expect(parseDateRange('1-')).toBeNull();
            expect(parseDateRange('-3')).toBeNull();
            expect(parseDateRange('1-2-3')).toBeNull();
        });
    });
});
