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
  tokenizeKqlValue,
  buildBareQueryTerms,
  buildFieldQueryTerms,
  buildTextKqlTerms,
  buildKqlBooleans,
  buildSearchParams,
  addBooleanFilters,
  progressiveSearch,
  formatSearchResults,
} = handleSearchEmails._internal;

const MOCK_TOKEN = 'dummy_access_token';
const MOCK_ENDPOINT = 'me/mailFolders/inbox/messages';

const mockEmail = (id, subject, isRead = true) => ({
  id,
  subject,
  from: { emailAddress: { name: 'Sender', address: 'sender@example.com' } },
  receivedDateTime: '2024-06-01T10:00:00Z',
  isRead,
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

  test('strips colons without adding phrase quotes', () => {
    expect(sanitizeKqlValue('Re: Hello')).toBe('Re Hello');
  });

  test('strips double quotes without adding phrase quotes', () => {
    expect(sanitizeKqlValue('say "hello"')).toBe('say hello');
  });

  test('handles combination of special characters', () => {
    expect(sanitizeKqlValue(`O'Brien: "test\\123"`)).toBe('OBrien test123');
  });

  test('trims leading and trailing whitespace after stripping', () => {
    expect(sanitizeKqlValue("' hello '")).toBe('hello');
  });
});

describe('tokenizeKqlValue', () => {
  test('splits sanitized text into non-empty tokens', () => {
    expect(tokenizeKqlValue('  Josh   Yellin  ')).toEqual(['Josh', 'Yellin']);
  });

  test('returns an empty array when sanitization removes all content', () => {
    expect(tokenizeKqlValue(':"\\')).toEqual([]);
  });
});

describe('buildBareQueryTerms', () => {
  test('returns fuzzy tokens by default', () => {
    expect(buildBareQueryTerms('budget meeting')).toEqual(['budget', 'meeting']);
  });

  test('returns a single phrase term when exact phrase mode is enabled', () => {
    expect(buildBareQueryTerms('budget meeting', true)).toEqual(['"budget meeting"']);
  });

  test('keeps single-token exact phrases unquoted', () => {
    expect(buildBareQueryTerms('budget', true)).toEqual(['budget']);
  });
});

describe('buildFieldQueryTerms', () => {
  test('returns repeated field clauses in fuzzy mode', () => {
    expect(buildFieldQueryTerms('from', 'Josh Yellin')).toEqual(['from:Josh', 'from:Yellin']);
  });

  test('returns a single field-scoped phrase in exact mode', () => {
    expect(buildFieldQueryTerms('from', 'Josh Yellin', true)).toEqual(['from:"Josh Yellin"']);
  });

  test('keeps single-token exact field matches unquoted', () => {
    expect(buildFieldQueryTerms('from', 'Josh', true)).toEqual(['from:Josh']);
  });

  test('drops empty field values after sanitization', () => {
    expect(buildFieldQueryTerms('from', ':"\\')).toEqual([]);
  });
});

describe('buildTextKqlTerms', () => {
  test('preserves stable query/subject/from/to ordering', () => {
    expect(buildTextKqlTerms({
      query: 'hello world',
      from: 'Josh Yellin',
      to: 'Alice',
      subject: 'Q4 Report',
    })).toEqual([
      'hello',
      'world',
      'subject:Q4',
      'subject:Report',
      'from:Josh',
      'from:Yellin',
      'to:Alice',
    ]);
  });

  test('supports mixed fuzzy and exact phrase options', () => {
    expect(buildTextKqlTerms(
      { query: '', from: 'Josh Yellin', to: '', subject: 'Q4 Report' },
      { fromExactPhrase: false, subjectExactPhrase: true }
    )).toEqual([
      'subject:"Q4 Report"',
      'from:Josh',
      'from:Yellin',
    ]);
  });
});

describe('buildKqlBooleans', () => {
  test('returns empty string when no boolean filters set', () => {
    expect(buildKqlBooleans({})).toBe('');
    expect(buildKqlBooleans({ hasAttachments: false, unreadOnly: false })).toBe('');
  });

  test('returns hasAttachments KQL when set', () => {
    expect(buildKqlBooleans({ hasAttachments: true })).toBe('hasAttachments:true');
  });

  test('does not emit isRead for unreadOnly', () => {
    expect(buildKqlBooleans({ unreadOnly: true })).toBe('');
  });

  test('returns only hasAttachments when both booleans are set', () => {
    expect(buildKqlBooleans({ hasAttachments: true, unreadOnly: true })).toBe('hasAttachments:true');
  });
});

describe('buildSearchParams', () => {
  test('uses fuzzy KQL for subject plus hasAttachments without $filter', () => {
    const params = buildSearchParams(
      { query: '', from: '', to: '', subject: 'Budget' },
      { hasAttachments: true },
      10
    );

    expect(params.$search).toBe('"subject:Budget hasAttachments:true"');
    expect(params.$filter).toBeUndefined();
  });

  test('uses fuzzy KQL for from plus unreadOnly without isRead in KQL', () => {
    const params = buildSearchParams(
      { query: '', from: 'john@example.com', to: '', subject: '' },
      { unreadOnly: true },
      10
    );

    expect(params.$search).toBe('"from:john@example.com"');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });

  test('uses fuzzy repeated field clauses for multi-word sender search', () => {
    const params = buildSearchParams(
      { query: '', from: 'Josh Yellin', to: '', subject: '' },
      {},
      10
    );

    expect(params.$search).toBe('"from:Josh from:Yellin"');
  });

  test('uses exact field phrase when explicitly requested', () => {
    const params = buildSearchParams(
      { query: '', from: 'Josh Yellin', to: '', subject: '' },
      {},
      10,
      { fromExactPhrase: true }
    );

    expect(params.$search).toBe('from:"Josh Yellin"');
  });

  test('uses fuzzy bare query terms by default', () => {
    const params = buildSearchParams(
      { query: 'budget meeting', from: '', to: '', subject: '' },
      {},
      10
    );

    expect(params.$search).toBe('"budget meeting"');
  });

  test('uses exact bare query phrase when explicitly requested', () => {
    const params = buildSearchParams(
      { query: 'budget meeting', from: '', to: '', subject: '' },
      {},
      10,
      { queryExactPhrase: true }
    );

    expect(params.$search).toBe('"budget meeting"');
  });

  test('supports mixed exact and fuzzy modes while preserving term order', () => {
    const params = buildSearchParams(
      { query: '', from: 'Josh Yellin', to: '', subject: 'Q4 Report' },
      {},
      10,
      { fromExactPhrase: false, subjectExactPhrase: true }
    );

    expect(params.$search).toBe('subject:"Q4 Report" from:Josh from:Yellin');
  });

  test('uses $search with all terms and hasAttachments in KQL', () => {
    const params = buildSearchParams(
      { query: 'hello', from: 'alice', to: 'bob', subject: 'Meeting' },
      { hasAttachments: true, unreadOnly: true },
      10
    );

    expect(params.$search).toBe('"hello subject:Meeting from:alice to:bob hasAttachments:true"');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });

  test('uses $filter when only boolean filters are present', () => {
    const params = buildSearchParams(
      { query: '', from: '', to: '', subject: '' },
      { hasAttachments: true, unreadOnly: true },
      10
    );

    expect(params.$search).toBeUndefined();
    expect(params.$filter).toBe('hasAttachments eq true and isRead eq false');
  });

  test('omits empty text terms after sanitization instead of emitting malformed KQL', () => {
    const params = buildSearchParams(
      { query: '', from: ':"\\', to: '', subject: '' },
      { hasAttachments: true },
      10
    );

    expect(params.$search).toBeUndefined();
    expect(params.$filter).toBe('hasAttachments eq true');
  });

  test('sanitizes special characters and keeps fuzzy defaults', () => {
    const params = buildSearchParams(
      { query: 'Re: Hello', from: "O'Connell", to: '', subject: 'Q4 Report' },
      {},
      10
    );

    expect(params.$search).toBe('"Re Hello subject:Q4 subject:Report from:OConnell"');
  });
});

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

  test('does not add $filter when no booleans are set', () => {
    const params = {};
    addBooleanFilters(params, {});
    expect(params.$filter).toBeUndefined();
  });
});

describe('handleSearchEmails and progressiveSearch', () => {
  test('defaults to fuzzy multi-word field matching when no phrase flag is provided', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ from: 'Josh Yellin' });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toBe('"from:Josh from:Yellin"');
  });

  test('accepts exact phrase flags through the handler', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ from: 'Josh Yellin', fromExactPhrase: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toBe('from:"Josh Yellin"');
  });

  test('combined search keeps boolean KQL and omits $filter', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ subject: 'Budget', hasAttachments: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toContain('subject:Budget');
    expect(params.$search).toContain('hasAttachments:true');
    expect(params.$filter).toBeUndefined();
  });

  test('combined search with unreadOnly still omits isRead from KQL', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'Test')] });

    await handleSearchEmails({ from: 'alice', unreadOnly: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toContain('from:alice');
    expect(params.$search).not.toContain('isRead');
    expect(params.$filter).toBeUndefined();
  });

  test('punctuation-only text falls back to filter-only params in the combined attempt', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'With Attachment')] });

    await handleSearchEmails({ subject: ':"\\', hasAttachments: true });

    const params = callGraphAPIPaginated.mock.calls[0][3];
    expect(params.$search).toBeUndefined();
    expect(params.$filter).toBe('hasAttachments eq true');
  });

  test('single-term fallback reuses fuzzy builder rules for field searches', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('combined failed'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ from: 'Josh Yellin' });

    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    const fallbackParams = callGraphAPIPaginated.mock.calls[1][3];
    expect(fallbackParams.$search).toBe('"from:Josh from:Yellin"');
  });

  test('single-term fallback reuses exact phrase rules for field searches', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('combined failed'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ from: 'Josh Yellin', fromExactPhrase: true });

    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    const fallbackParams = callGraphAPIPaginated.mock.calls[1][3];
    expect(fallbackParams.$search).toBe('from:"Josh Yellin"');
  });

  test('single-term fallback reuses fuzzy builder rules for query searches', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('combined failed'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ query: 'budget meeting' });

    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    const fallbackParams = callGraphAPIPaginated.mock.calls[1][3];
    expect(fallbackParams.$search).toBe('"budget meeting"');
  });

  test('single-term fallback reuses exact phrase rules for query searches', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('combined failed'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Budget')] });

    await handleSearchEmails({ query: 'budget meeting', queryExactPhrase: true });

    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    const fallbackParams = callGraphAPIPaginated.mock.calls[1][3];
    expect(fallbackParams.$search).toBe('"budget meeting"');
  });

  test('skips single-term fallback when the term sanitizes to empty and uses boolean-only fallback', async () => {
    callGraphAPIPaginated
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [] });

    const response = await progressiveSearch(
      MOCK_ENDPOINT,
      MOCK_TOKEN,
      { query: '', from: '', to: '', subject: ':"\\' },
      { hasAttachments: true },
      10
    );

    expect(callGraphAPIPaginated).toHaveBeenCalledTimes(2);
    expect(callGraphAPIPaginated.mock.calls[0][3].$filter).toBe('hasAttachments eq true');
    expect(callGraphAPIPaginated.mock.calls[1][3].$filter).toBe('hasAttachments eq true');
    expect(response._searchInfo.strategies).toEqual(['combined-search', 'boolean-filters-only']);
  });

  test('over-fetches and post-filters unread results only when effective KQL terms exist', async () => {
    callGraphAPIPaginated.mockResolvedValue({
      value: [
        mockEmail('1', 'A', false),
        mockEmail('2', 'B', true),
        mockEmail('3', 'C', false),
      ],
    });

    const response = await progressiveSearch(
      MOCK_ENDPOINT,
      MOCK_TOKEN,
      { query: '', from: 'alice bob', to: '', subject: '' },
      { unreadOnly: true },
      2
    );

    expect(callGraphAPIPaginated.mock.calls[0][3].$top).toBe(6);
    expect(callGraphAPIPaginated.mock.calls[0][4]).toBe(6);
    expect(response.value.map(email => email.id)).toEqual(['1', '3']);
  });

  test('does not over-fetch unread results when text input sanitizes to nothing', async () => {
    callGraphAPIPaginated.mockResolvedValue({ value: [mockEmail('1', 'A', false)] });

    await progressiveSearch(
      MOCK_ENDPOINT,
      MOCK_TOKEN,
      { query: ':"\\', from: '', to: '', subject: '' },
      { unreadOnly: true },
      2
    );

    expect(callGraphAPIPaginated.mock.calls[0][3].$search).toBeUndefined();
    expect(callGraphAPIPaginated.mock.calls[0][3].$filter).toBe('isRead eq false');
    expect(callGraphAPIPaginated.mock.calls[0][4]).toBe(2);
  });

  test('returns no results without calling Graph when text input sanitizes away and no filters remain', async () => {
    const response = await progressiveSearch(
      MOCK_ENDPOINT,
      MOCK_TOKEN,
      { query: ':"\\', from: '', to: '', subject: '' },
      {},
      2
    );

    expect(callGraphAPIPaginated).not.toHaveBeenCalled();
    expect(response.value).toEqual([]);
    expect(response._searchInfo.warning).toContain('removed during sanitization');
  });

  test('single-term fallback attaches _searchInfo to successful fallback results', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Test')] });

    const result = await handleSearchEmails({ subject: 'Locked', from: 'bob' });

    expect(result.content[0].text).toContain('single-term-subject');
  });

  test('boolean-only fallback attaches warning when raw search terms were dropped', async () => {
    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))
      .mockRejectedValueOnce(new Error('fail'))
      .mockRejectedValueOnce(new Error('fail'))
      .mockResolvedValueOnce({ value: [mockEmail('1', 'Unrelated')] });

    const result = await handleSearchEmails({
      subject: 'Budget',
      from: 'alice',
      hasAttachments: true,
    });

    expect(result.content[0].text).toContain('boolean-filters-only');
  });

  test('returns empty results with _searchInfo when recent-emails fallback is disabled', async () => {
    expect(ENABLE_RECENT_EMAILS_FALLBACK).toBe(false);

    callGraphAPIPaginated
      .mockRejectedValueOnce(new Error('fail'))
      .mockRejectedValueOnce(new Error('fail'))
      .mockRejectedValueOnce(new Error('fail'));

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
  test('includes attempted strategies in the no-results message when _searchInfo is present', () => {
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
