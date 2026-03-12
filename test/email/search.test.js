const handleSearchEmails = require('../../email/search');
const { callGraphAPIPaginated } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');
const { resolveFolderPath } = require('../../email/folder-utils');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');
jest.mock('../../email/folder-utils');

const {
  ENABLE_RECENT_EMAILS_FALLBACK,
  sanitizeKqlValue,
  buildKqlBooleans,
  buildSearchParams,
  addBooleanFilters,
  progressiveSearch,
  formatSearchResults,
} = handleSearchEmails._internal;

const MOCK_TOKEN = 'dummy_access_token';
const MOCK_ENDPOINT = 'me/mailFolders/inbox/messages';

const mockEmail = (id, subject) => ({
  id,
  subject,
  from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
  receivedDateTime: '2024-06-01T10:00:00Z',
  isRead: true,
});

beforeEach(() => {
  callGraphAPIPaginated.mockClear();
  ensureAuthenticated.mockClear();
  resolveFolderPath.mockClear();
  jest.spyOn(console, 'error').mockImplementation(() => {});
  ensureAuthenticated.mockResolvedValue(MOCK_TOKEN);
  resolveFolderPath.mockResolvedValue(MOCK_ENDPOINT);
});

afterEach(() => {
  console.error.mockRestore();
});

// ---------------------------------------------------------------------------
// Bug 3: sanitizeKqlValue
// ---------------------------------------------------------------------------
describe('sanitizeKqlValue', () => {
  test('returns falsy values unchanged', () => {
    expect(sanitizeKqlValue('')).toBe('');
    expect(sanitizeKqlValue(null)).toBe(null);
    expect(sanitizeKqlValue(undefined)).toBe(undefined);
  });

  test('strips single quotes', () => {
    expect(sanitizeKqlValue("O'Connell")).toBe('OConnell');
  });

  test('strips backslashes', () => {
    expect(sanitizeKqlValue('path\\to\\file')).toBe('pathtofile');
  });

  test('strips colons', () => {
    expect(sanitizeKqlValue('Re: Hello')).toBe('"Re Hello"');
  });

  test('strips double quotes', () => {
    expect(sanitizeKqlValue('say "hello"')).toBe('"say hello"');
  });

  test('handles combination of special characters', () => {
    expect(sanitizeKqlValue(`O'Brien: "test\\123"`)).toBe('"OBrien test123"');
  });

  test('trims leading/trailing whitespace after stripping', () => {
    expect(sanitizeKqlValue("' hello '")).toBe('hello');
  });
});

// ---------------------------------------------------------------------------
// buildKqlBooleans
// ---------------------------------------------------------------------------
describe('buildKqlBooleans', () => {
  test('returns empty string when no boolean filters set', () => {
    expect(buildKqlBooleans({})).toBe('');
    expect(buildKqlBooleans({ hasAttachments: false, unreadOnly: false })).toBe('');
  });

  test('returns hasAttachments KQL when set', () => {
    expect(buildKqlBooleans({ hasAttachments: true })).toBe('hasAttachments:true');
  });

  test('does not emit isRead for unreadOnly (handled via post-filter)', () => {
    expect(buildKqlBooleans({ unreadOnly: true })).toBe('');
  });

  test('returns only hasAttachments when both set (isRead handled via post-filter)', () => {
    expect(buildKqlBooleans({ hasAttachments: true, unreadOnly: true }))
      .toBe('hasAttachments:true');
  });
});

// ---------------------------------------------------------------------------
// Bug 1: buildSearchParams — no $filter when $search is present
// ---------------------------------------------------------------------------
describe('buildSearchParams', () => {
  test('uses $search with KQL for subject + hasAttachments (no $filter)', () => {
    const params = buildSearchParams(
      { query: '', from: '', to: '', subject: 'Budget' },
      { hasAttachments: true },
      10
    );
    expect(params.$search).toBe('"subject:Budget hasAttachments:true"');
    expect(params.$filter).toBeUndefined();
  });

  test('uses $search for from + unreadOnly (no isRead in KQL, no $filter)', () => {
    const params = buildSearchParams(
      { query: '', from: 'john@example.com', to: '', subject: '' },
      { unreadOnly: true },
      10
    );
    // isRead is NOT a valid KQL property — unreadOnly handled via post-filter
    expect(params.$search).toBe('"from:john@example.com"');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });

  test('uses $search with KQL for all terms + both booleans (no $filter, no isRead in KQL)', () => {
    const params = buildSearchParams(
      { query: 'hello', from: 'alice', to: 'bob', subject: 'Meeting' },
      { hasAttachments: true, unreadOnly: true },
      10
    );
    expect(params.$search).toBe('"hello subject:Meeting from:alice to:bob hasAttachments:true"');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });

  test('uses $filter (not $search) when only boolean filters and no search terms', () => {
    const params = buildSearchParams(
      { query: '', from: '', to: '', subject: '' },
      { hasAttachments: true, unreadOnly: true },
      10
    );
    expect(params.$search).toBeUndefined();
    expect(params.$filter).toBe('hasAttachments eq true and isRead eq false');
  });

  test('no $search and no $filter when nothing is specified', () => {
    const params = buildSearchParams(
      { query: '', from: '', to: '', subject: '' },
      {},
      10
    );
    expect(params.$search).toBeUndefined();
    expect(params.$filter).toBeUndefined();
  });

  test('sanitizes special characters and phrase-quotes multi-word values', () => {
    const params = buildSearchParams(
      { query: '', from: "O'Connell", to: '', subject: 'Re: Hello' },
      {},
      10
    );
    // Single quote removed, colon removed, multi-word values phrase-quoted
    expect(params.$search).toBe('"subject:"Re Hello" from:OConnell"');
    expect(params.$filter).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// addBooleanFilters (unchanged, but verify still works for filter-only path)
// ---------------------------------------------------------------------------
describe('addBooleanFilters', () => {
  test('adds $filter for hasAttachments', () => {
    const params = {};
    addBooleanFilters(params, { hasAttachments: true });
    expect(params.$filter).toBe('hasAttachments eq true');
  });

  test('adds $filter for unreadOnly', () => {
    const params = {};
    addBooleanFilters(params, { unreadOnly: true });
    expect(params.$filter).toBe('isRead eq false');
  });

  test('does not add $filter when no booleans', () => {
    const params = {};
    addBooleanFilters(params, {});
    expect(params.$filter).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Bug 1 (integration): combined search never sends $filter + $search together
// ---------------------------------------------------------------------------
describe('progressiveSearch — combined search params', () => {
  test('subject + hasAttachments sends $search without $filter', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ subject: 'Budget', hasAttachments: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toContain('subject:Budget');
    expect(params.$search).toContain('hasAttachments:true');
    expect(params.$filter).toBeUndefined();
  });

  test('from + unreadOnly sends $search without isRead (post-filtered instead)', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Test')] });

    await handleSearchEmails({ from: 'alice', unreadOnly: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toContain('from:alice');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Bug 1 (integration): single-term fallback also folds booleans into KQL
// ---------------------------------------------------------------------------
describe('progressiveSearch — single-term fallback params', () => {
  test('when combined fails, single-term search still folds booleans into KQL', async () => {
    // First call (combined) fails
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('Simulated combined failure'))
      // Second call (single-term subject) succeeds
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ subject: 'Budget', hasAttachments: true });

    // The second call is the single-term fallback
    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    const fallbackParams = callGraphAPIPaginated.mock.calls[1][3];
    expect(fallbackParams.$search).toContain('subject:Budget');
    expect(fallbackParams.$search).toContain('hasAttachments:true');
    expect(fallbackParams.$filter).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Bug 2: _searchInfo attached on fallback strategies
// ---------------------------------------------------------------------------
describe('progressiveSearch — _searchInfo on fallbacks', () => {
  test('single-term fallback attaches _searchInfo', async () => {
    // Combined fails, single-term subject succeeds
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Test')] });

    const result = await handleSearchEmails({ subject: 'Locked', from: 'bob' });

    // Result should mention fallback strategy
    expect(result.content[0].text).toContain('single-term-subject');
  });

  test('boolean-only fallback attaches _searchInfo with warning when search terms were dropped', async () => {
    // Combined fails, single-term subject fails, single-term from fails, boolean-only succeeds
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))   // combined
      .mockRejectedValueOnce(new Error('fail'))   // single subject
      .mockRejectedValueOnce(new Error('fail'))   // single from
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Unrelated')] }); // boolean-only

    const result = await handleSearchEmails({
      subject: 'Budget',
      from: 'alice',
      hasAttachments: true,
    });

    // The response should contain the strategy info
    expect(result.content[0].text).toContain('boolean-filters-only');
  });

  test('final recent-emails fallback attaches _searchInfo', async () => {
    // Everything fails; with fallback disabled, no unrelated recent emails are returned
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))   // combined
      .mockRejectedValueOnce(new Error('fail'))   // single subject
      .mockRejectedValueOnce(new Error('fail'));  // boolean-only

    const result = await handleSearchEmails({
      subject: 'Budget',
      hasAttachments: true,
    });

    expect(result.content[0].text).toContain('No emails found matching your search criteria.');
    expect(result.content[0].text).toContain(
      'Strategies attempted: combined-search, single-term-subject, boolean-filters-only.'
    );
    expect(result.content[0].text).toContain(
      'All search strategies failed. No results returned (recent-emails fallback is disabled).'
    );
    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(3);
  });
});

describe('progressiveSearch — recent-emails fallback flag', () => {
  test('returns empty results with _searchInfo when recent-emails fallback is disabled', async () => {
    expect(ENABLE_RECENT_EMAILS_FALLBACK).toBe(false);

    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))   // combined
      .mockRejectedValueOnce(new Error('fail'))   // single subject
      .mockRejectedValueOnce(new Error('fail'));  // boolean-only

    const response = await progressiveSearch(
      MOCK_ENDPOINT,
      MOCK_TOKEN,
      { query: '', from: '', to: '', subject: 'Budget' },
      { hasAttachments: true },
      10
    );

    expect(response).toEqual({
      value: [],
      _searchInfo: {
        attemptsCount: 3,
        strategies: ['combined-search', 'single-term-subject', 'boolean-filters-only'],
        originalTerms: { query: '', from: '', to: '', subject: 'Budget' },
        filterTerms: { hasAttachments: true },
        warning: 'All search strategies failed. No results returned (recent-emails fallback is disabled).'
      }
    });
  });
});

describe('formatSearchResults', () => {
  test('includes attempted strategies in no-results message when _searchInfo is present', () => {
    const result = formatSearchResults({
      value: [],
      _searchInfo: {
        strategies: [
          'combined-search',
          'single-term-subject',
          'single-term-from',
          'boolean-filters-only'
        ]
      }
    });

    expect(result.content[0].text).toBe(
      'No emails found matching your search criteria. ' +
      'Strategies attempted: combined-search, single-term-subject, single-term-from, boolean-filters-only.'
    );
  });
});

// ---------------------------------------------------------------------------
// Bug 3 (integration): special characters in real search
// ---------------------------------------------------------------------------
describe('progressiveSearch — special characters in search terms', () => {
  test("apostrophe in subject is sanitized in KQL", async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', "O'Connell Report")] });

    await handleSearchEmails({ subject: "O'Connell" });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toBe('"subject:OConnell"');
    // No raw apostrophe
    expect(params.$search).not.toContain("'");
  });

  test('colon in query is sanitized', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Re: test')] });

    await handleSearchEmails({ query: 'Re: Hello' });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toBe('"\"Re Hello\""');
  });
});

// ---------------------------------------------------------------------------
// Existing behavior preserved: boolean-only search with no search terms
// ---------------------------------------------------------------------------
describe('boolean-only search (no search terms)', () => {
  test('uses $filter when only hasAttachments is set (no search terms)', async () => {
    // Combined search returns no results (no search terms → no $search)
    // Falls through to boolean-only
    callGraphAPIPaginated
      .mockResolvedValueOnce({ value: [] })  // combined (empty)
      .mockResolvedValueOnce({ value: [mockEmail('1', 'With Attachment')] }); // boolean-only

    await handleSearchEmails({ hasAttachments: true });

    // The boolean-only call should use $filter
    const booleanCall = callGraphAPIPaginated.mock.calls[1][3];
    expect(booleanCall.$filter).toBe('hasAttachments eq true');
    expect(booleanCall.$search).toBeUndefined();
  });
});
