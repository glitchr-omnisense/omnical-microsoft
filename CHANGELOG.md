# Changelog

## 1.x (unreleased)

- First version: Microsoft Graph's calendars as an omnical provider (`microsoft`).
  - calendars, events (calendarView with paging, or series), one event, create, update, delete;
  - getSchedule free/busy for work or school accounts;
  - OAuth 2.0 through the core's `OAuth2Client`, with tenant, PKCE and refresh-token rotation;
  - one refresh and one retry on a 401;
  - `Prefer` for UTC times, immutable ids and text bodies;
  - Windows time zones read as IANA (`Zones`: ext-intl, or a table);
  - `PatternedRecurrence` between Graph's patterns and RRULE, both ways;
  - a read-then-write `changeKey` check;
  - `policy.json` with Outlook's throttling limit as 1,000 a minute per user, and Microsoft's terms.
