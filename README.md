# omnical/microsoft

Microsoft Graph for [omnisense/omnical](https://github.com/glitchr-omnisense/omnical), under any
PHP host: the signed-in user's Outlook and Microsoft 365 calendars over OAuth 2.0.

- the user's calendars, the default one first;
- their events read: each occurrence of a series, or the series themselves;
- events written: created, updated and deleted, attendees included;
- time zones kept: Graph's Windows names (`Romance Standard Time`) are read back as IANA names;
- recurrence as RRULE data, both ways (Graph's `patternedRecurrence`);
- a check on `changeKey` before a write, against lost updates;
- the busy time of people by their address (`getSchedule`);
- Microsoft's throttling limits declared and enforced.

```
composer require omnical/microsoft
```

## Registering the application (Microsoft Entra)

1. In the Entra admin centre (App registrations), register an application. Pick the accounts it
   serves. That choice is the `tenant` option:
   - `common`: work, school and personal accounts;
   - `organizations`;
   - `consumers`;
   - your tenant's id.
2. Add a **Web** redirect URI, the address the user comes back to with the code
   (`https://example.org/oauth/microsoft`). It must match exactly.
3. Under Certificates & secrets, create a client secret. Keep it as a secret
   (`MICROSOFT_CLIENT_SECRET` in the environment, the vault): the provider takes its **reference**,
   never its value.
4. Under API permissions, add the delegated permissions:
   - `Calendars.ReadWrite` to write, or `Calendars.Read` to read only;
   - `offline_access` for a refresh token.

   None needs an administrator's consent for delegated use. `getSchedule`'s pages disagree on its
   least permission (`Calendars.ReadBasic` in its reference, `Calendars.Read` in the free/busy
   guide); `Calendars.Read` satisfies both.

## The consent step, then the provider

```php
use Omnical\Auth\InMemoryTokenStore;
use Omnical\Auth\OAuth2Client;
use Omnical\Microsoft\MicrosoftFactory;
use Omnical\Model\EventQuery;
use Omnical\Registry;

$tokens = new MyDoctrineTokenStore();                  // an Omnical\Auth\TokenStoreInterface
$factory = new MicrosoftFactory($httpClient, $tokens);
$options = ['client_id' => '…', 'client_secret' => 'MICROSOFT_CLIENT_SECRET', 'token' => 'user:42:microsoft'];

// 1. Send the user to consent (PKCE: keep the verifier in the session).
[$verifier, $challenge] = OAuth2Client::pkce();
$url = $factory->oauth($options)->authorizationUrl('https://example.org/oauth/microsoft', MicrosoftFactory::scopes(), $state, $challenge);

// 2. Back on the redirect URI: the code becomes the user's token, kept in the store.
$tokens->save('user:42:microsoft', $factory->oauth($options)->exchange($_GET['code'], 'https://example.org/oauth/microsoft', $verifier));

// 3. Then, on any request:
$registry = new Registry([$factory], ['outlook' => ['factory' => 'microsoft', 'options' => $options]]);
$events = $registry->calendarEvents('outlook', 'user-42')->events(new EventQuery($calendarId, $from, $to));
```

The token is refreshed when it expires, or once when Graph answers 401. The store gets the new
refresh token each time, because Microsoft's refresh tokens replace themselves. A refresh token
lasts 90 days by default and can be revoked at any time; then an `AuthenticationException` says
the user has to consent again.

## Options

| Option | Default | |
|---|---|---|
| `token` | — | the key of the user's token in the factory's token store |
| `access_token` | — | instead of `token`: the reference of a token given as it is (never refreshed) |
| `client_id` | — | the application (client) id; needed with `token` |
| `client_secret` | — | the **reference** of the client secret (an environment variable by default, or the factory's `SecretsInterface`) |
| `tenant` | `common` | `organizations`, `consumers`, a tenant id or domain |
| `base_uri` | `https://graph.microsoft.com/v1.0` | a national cloud's Graph |
| `timeout` | the HTTP client's | seconds |
| `availability_interval` | `30` | getSchedule's minutes per slot, 5 to 1440 |
| `policy`, `usage` | — | the site's tightening of the policy, what the site says of itself |

The factory takes the HTTP client, the token store, the secrets (`EnvSecrets` by default) and a
clock, in that order.

## The calls

Every request carries the token as `Authorization: Bearer` and one `Prefer` field (RFC 7240) with
three preferences:
- `outlook.timezone="UTC"`: times come back in UTC and are set in the event's own zone;
- `IdType="ImmutableId"`: an event keeps its id when it moves between calendars;
- `outlook.body-content-type="text"`.

| Contract | Graph |
|---|---|
| `calendars()` | `GET /me/calendars`, every page; `canEdit`, `isDefaultCalendar`, `hexColor`, `owner`, `changeKey` |
| `events()`, `expand` (default) | `GET /me/calendars/{id}/calendarView?startDateTime&endDateTime&$top&$orderby=start/dateTime`, `@odata.nextLink` followed up to the query's limit (on Graph's host only) |
| `events()`, no `expand` | `GET /me/calendars/{id}/events?$filter=start/dateTime lt '…' and end/dateTime gt '…'`: single events and series masters |
| `event()` | `GET /me/events/{id}`; a 404 is null |
| `create()` | `POST /me/calendars/{id}/events` (201) |
| `update()` | `PATCH /me/events/{id}` |
| `delete()` | `DELETE /me/events/{id}` (204; a 404 is nothing to delete) |
| `freeBusy()` | `POST /me/calendar/getSchedule`: people by address. Each `scheduleItems` entry is a busy span: `busy`, `tentative`, `oof` (out of office), `unknown` (read as busy); `free` and `workingElsewhere` are free. The schedule's `error` is `FreeBusy::$error`. |

How each field is read and written:
- `subject` is the title, `body.content` the description, `location.displayName` the location.
- `start`/`end` are written in the event's IANA zone, or UTC when it has none. An all-day event is
  written as two midnights in one zone, as Graph requires.
- `originalStartTimeZone` gives the event's zone, through `Zones`:
  - with ext-intl, any Windows name;
  - without it, a table of 43 common zones;
  - Outlook's custom zones are read as UTC.
- `showAs`: `free` is a transparent event, `tentative` a tentative one. `isCancelled` is a
  cancelled event.
- Attendees: `type` gives the role (`resource` is a room or equipment). `status.response` gives
  the answer (`tentativelyAccepted`, `notResponded`…).
- The organizer is an attendee with `organizer: true`. It is never written: Graph sets it.
- An occurrence or exception carries `seriesMasterId` (as `recurringEventId`) and `originalStart`.
- Other fields: `iCalUId` is the uid, `webLink` the url, `changeKey` the etag, and
  `lastModifiedDateTime` the update time.

### Recurrence

`PatternedRecurrence` maps Graph's patterns to RRULE and back:

| Graph | RRULE |
|---|---|
| `daily` | `FREQ=DAILY;INTERVAL=n` |
| `weekly` | `FREQ=WEEKLY;BYDAY=…`, with `WKST` unless Monday |
| `absoluteMonthly` | `BYMONTHDAY` |
| `relativeMonthly` | `BYDAY=2TU`, or `BYDAY=MO,…,FR;BYSETPOS=-1` |
| `absoluteYearly` | `BYMONTH` + `BYMONTHDAY` |
| `relativeYearly` | `BYMONTH` + `BYDAY` |
| range `endDate` | `UNTIL`: the last day's end in the recurrence's zone, in UTC; a date for an all-day series |
| range `numbered` | `COUNT` |
| range `noEnd` | neither |

A weekdays-only daily rule (`FREQ=DAILY;BYDAY=MO,…,FR`) is written as a weekly pattern. A rule
Graph cannot express throws `NotSupportedException` instead of another rule being written in its
place:
- `BYHOUR`;
- `FREQ=MINUTELY`;
- several rules, `EXRULE`, `RDATE`, `EXDATE` (to take one occurrence out, delete it);
- the fifth Monday;
- a negative or repeated `BYMONTHDAY`.

### Lost updates: a check, not a lock

Graph documents no `If-Match` on events. When `update()` or `delete()` is given an etag (the
`changeKey` read with the event), the provider first reads the event and compares the two. If the
`changeKey` changed, it throws `ConflictException` (412) and writes nothing. If the event is gone,
`update()` throws the same, and `delete()` does nothing. This is a read-then-write check, not an
atomic condition: a change that lands between the read and the write is not seen. Without an etag,
the write goes through as it is.

## Usage policy

`policy.json` (omnisense/contracts' schema), read on 2026-11-10:

- **Rate:**
  - Outlook allows 10,000 requests in any 10 minutes per application and mailbox, 4 at a time,
    and 150 MB of uploads per 5 minutes.
  - The schema has no 10-minute window, so **1,000 a minute per user** is declared. It never
    exceeds the real limit, and it counts per the `$user` given to `Registry::calendarEvents()` and
    `freeBusy()`.
  - A 429 and its `Retry-After` hold every call until the wait is over.
- **Rights:**
  - commercial use allowed: no ban, though Microsoft may set fees;
  - redistribution, resale and sublicensing of the data forbidden;
  - no attribution.
  - Caching and storage are allowed only as the intended scenario needs. Data kept stays current,
    follows retention and deletion rules, and is deleted when the user leaves.
  - Ask for the minimum permissions.
  - No advertising use.
- **Terms:** <https://learn.microsoft.com/en-us/legal/microsoft-apis/terms-of-use>.
- **Limits:** <https://learn.microsoft.com/en-us/graph/throttling-limits>.

## Verified, and not

- **Verified** against Microsoft's documentation on 2026-11-10:
  - every path, parameter and status above;
  - the event, calendar, dateTimeTimeZone, patternedRecurrence and scheduleInformation resources;
  - the identity platform's endpoints, PKCE and refresh-token rotation;
  - the throttling limits and the terms.
- **Tested:** the tests answer from fixtures written from those documented shapes (made-up ids and
  addresses), and assert each request's method, URL, query, headers and body.
- **Not verified:** no call has been made to Microsoft Graph itself.
  - The comma-separated `Prefer` field (RFC 7240) has not been tried against Graph. The docs'
    examples send one preference per field.
  - IANA names on write are documented with a caveat: Microsoft says "create or update might not
    support all" of them.
  - The `$filter` of a non-expanded query misses a series whose first occurrence lies before the
    span; ask with `expand` for occurrences.
  - The query's words are matched here, on the title, description and location; no `$search` is
    sent.

MIT licence.
