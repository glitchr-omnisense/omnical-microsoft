# omnical/microsoft

Microsoft Graph's calendars for omnical: the user's calendars, their events read and written, and
people's busy time. See the [README](../README.md) for:
- the Entra registration;
- the consent step;
- the options;
- every call and how each field is read;
- the recurrence mapping;
- the usage policy;
- what is verified and what is not.

## In plain PHP

```php
use Omnical\Auth\InMemoryTokenStore;
use Omnical\Auth\OAuthToken;
use Omnical\Microsoft\MicrosoftFactory;
use Omnical\Model\Event;
use Omnical\Model\FreeBusy;
use Omnical\Model\FreeBusyQuery;
use Omnical\Registry;

$tokens = new InMemoryTokenStore(['me' => OAuthToken::fromArray($savedToken)]);
$registry = new Registry([new MicrosoftFactory(null, $tokens)], [
    'outlook' => ['factory' => 'microsoft', 'options' => ['client_id' => '…', 'client_secret' => 'MICROSOFT_CLIENT_SECRET', 'token' => 'me']],
]);

$calendar = $registry->calendarEvents('outlook', 'me');
$primary = $calendar->calendars()[0];
$created = $calendar->create(new Event('Consultation', new DateTimeImmutable('2026-11-23 10:00', new DateTimeZone('Europe/Paris')), new DateTimeImmutable('2026-11-23 10:45', new DateTimeZone('Europe/Paris')), calendarId: $primary->id));
$calendar->update($created->moved($newStart, $newEnd));   // ConflictException when someone changed it since

$busy = $registry->freeBusy('outlook', 'me')->freeBusy(new FreeBusyQuery(['anna@contoso.example', 'paul@contoso.example'], $from, $to));
$slots = FreeBusy::common($busy, $from, $to, 3600);         // the hours both are free
```

## Testing your application

`Omnical\Testing\RecordedHttpClient` answers Graph's calls from your own fixtures. The package's
tests (`Tests/MicrosoftProviderTest.php`) show the routes and the documented answer shapes.
