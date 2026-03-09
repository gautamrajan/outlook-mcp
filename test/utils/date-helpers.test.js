const {
  resolveIanaTimezone,
  formatAllDayRange,
  formatEmailDate,
} = require('../../utils/date-helpers');

describe('date helpers', () => {
  afterEach(() => {
    jest.restoreAllMocks();
    jest.useRealTimers();
  });

  describe('resolveIanaTimezone', () => {
    test('maps Windows timezone names to IANA names', () => {
      expect(resolveIanaTimezone('W. Europe Standard Time')).toBe('Europe/Berlin');
    });

    test('accepts IANA timezone names directly', () => {
      expect(resolveIanaTimezone('Europe/Berlin')).toBe('Europe/Berlin');
    });

    test('falls back to Pacific time for unknown timezone names', () => {
      const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});

      expect(resolveIanaTimezone('Mars/Phobos')).toBe('America/Los_Angeles');
      expect(warnSpy).toHaveBeenCalledWith(
        expect.stringContaining('Unmapped timezone "Mars/Phobos"')
      );
    });
  });

  describe('formatAllDayRange', () => {
    test('formats a single-day all-day event without midnight times', () => {
      expect(
        formatAllDayRange(
          '2026-03-09T00:00:00.0000000',
          '2026-03-10T00:00:00.0000000',
          { year: 2026, month: 3, day: 9 }
        )
      ).toBe('Mon, Mar 9 (All day) (Today)');
    });

    test('formats a multi-day all-day event using the inclusive end date', () => {
      expect(
        formatAllDayRange(
          '2026-03-09T00:00:00.0000000',
          '2026-03-12T00:00:00.0000000',
          { year: 2026, month: 3, day: 8 }
        )
      ).toBe('Mon, Mar 9 – Wed, Mar 11 (All day) (Tomorrow)');
    });
  });

  describe('formatEmailDate', () => {
    test('uses the configured timezone year when deciding whether to show the year', () => {
      jest.useFakeTimers();
      jest.setSystemTime(new Date('2025-12-31T23:30:00Z'));

      expect(
        formatEmailDate('2025-12-31T15:30:00Z', 'Asia/Tokyo')
      ).toBe('Thu, Jan 1 · 12:30 AM');
    });

    test('includes the year when the email is outside the current year in that timezone', () => {
      jest.useFakeTimers();
      jest.setSystemTime(new Date('2025-12-31T23:30:00Z'));

      expect(
        formatEmailDate('2024-12-31T15:30:00Z', 'Asia/Tokyo')
      ).toBe('Wed, Jan 1, 2025 · 12:30 AM');
    });
  });
});
