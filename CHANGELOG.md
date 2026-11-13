# Changelog

## [Unreleased]

## [1.0.0] - 2026-11-13

The first stable version. It requires `omnisense/contracts` `^1.0`, `omnisense/omnical` `^1.0` with no stability flag, so `composer require omnical/microsoft` installs it in a project with Composer's default minimum-stability; the 1.x series follows semantic versioning. A tag `vX.Y.Z` is released by `.github/workflows/release.yml` once the package's CI passes on it.

- 2026-11-10, on omnisense/contracts' round 2 (5bac28b or later, still `^1.0@dev`): Outlook's
  limit declared as it is, 10,000 every 10 minutes per user (`"every": 10`, the 1,000 a minute
  stand-in gone), and its 4 requests at a time as `"concurrency": 4`.
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
