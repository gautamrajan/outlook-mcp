# Memory Index

## Outlook MCP Architecture
- Main entry: `index.js`, modules: `auth/`, `calendar/`, `email/`, `folder/`, `rules/`, `utils/`
- Email search: `email/search.js` uses progressive fallback (combined -> individual terms -> boolean filters -> recent)
- Search is folder-scoped by default (no cross-folder search)
- `bodyPreview` is fetched from Graph API but NOT surfaced in formatted search/list output (gap identified)
- Max 50 results per search call, KQL via `$search` parameter
- No date-range filter on search-emails tool
- No conversationId exposed in results
- Graph API client: `utils/graph-api.js` with pagination support and 401 auto-retry
- Config: `config.js` defines `EMAIL_SELECT_FIELDS`, `EMAIL_DETAIL_FIELDS`, etc.

## Design Documents
- Email excavation skill design: `docs/email-excavation-skill-design.md`
