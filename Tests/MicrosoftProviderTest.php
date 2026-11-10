<?php

namespace Omnical\Microsoft\Tests;

use Omnical\Auth\ArraySecrets;
use Omnical\Auth\InMemoryTokenStore;
use Omnical\Auth\OAuthToken;
use Omnical\Exception\AuthenticationException;
use Omnical\Exception\ConflictException;
use Omnical\Exception\InvalidConfigException;
use Omnical\Http\Data;
use Omnical\Microsoft\Graph;
use Omnical\Microsoft\MicrosoftFactory;
use Omnical\Microsoft\MicrosoftProvider;
use Omnical\Microsoft\Zones;
use Omnical\Model\Attendee;
use Omnical\Model\AttendeeRole;
use Omnical\Model\BusySlot;
use Omnical\Model\Calendar;
use Omnical\Model\Event;
use Omnical\Model\EventQuery;
use Omnical\Model\EventStatus;
use Omnical\Model\FreeBusyQuery;
use Omnical\Model\FreeBusyStatus;
use Omnical\Model\ParticipationStatus;
use Omnical\Model\Recurrence;
use Omnical\Model\Transparency;
use Omnical\Testing\RecordedHttpClient;
use PHPUnit\Framework\TestCase;

/**
 * Microsoft Graph's calls checked against recorded answers. No live call was
 * made: the fixtures are written from the resources Graph documents
 * (calendar, event, dateTimeTimeZone, patternedRecurrence,
 * scheduleInformation, as learn.microsoft.com gave them on 2026-11-10), with
 * made-up ids and addresses; every request's method, URL, query, headers and
 * body is asserted.
 */
final class MicrosoftProviderTest extends TestCase
{
    private const GRAPH = 'https://graph.microsoft.com/v1.0';
    private const CALENDAR = 'AAMkAGI2-default=';

    public function testTheCalendarsPrimaryFirstWithTheTokenAndPreferences(): void
    {
        $http = self::http();
        $calendars = self::provider($http)->calendars();

        self::assertSame([
            [self::CALENDAR, 'Calendar', true, false, null, 'anna@contoso.example', 'ck-cal-1'],
            ['AAMkAGI2-shared=', 'Planning du cabinet', false, true, '#8BC34A', 'cabinet@contoso.example', 'ck-cal-2'],
        ], array_map(static fn (Calendar $c) => [$c->id, $c->name, $c->primary, $c->readOnly, $c->color, $c->owner, $c->version], $calendars));
        self::assertSame(['GET', self::GRAPH.'/me/calendars', ['$top' => 100]], [$http->requests[0]['method'], $http->requests[0]['url'], $http->requests[0]['options']['query'] ?? null]);
        self::assertSame('Bearer at-1', $http->header('Authorization'));
        self::assertSame('outlook.timezone="UTC", IdType="ImmutableId", outlook.body-content-type="text"', $http->header('Prefer'));
        self::assertSame('application/json', $http->header('Accept'));
    }

    public function testOccurrencesPageByPageInTheEventsZone(): void
    {
        $http = self::http();
        $events = self::provider($http)->events(new EventQuery(self::CALENDAR, new \DateTimeImmutable('2026-11-16T01:00', new \DateTimeZone('Europe/Paris')), new \DateTimeImmutable('2026-12-31T00:00:00Z'), limit: 10));

        self::assertCount(4, $events);
        self::assertCount(2, $http->requests, 'the next page followed');
        self::assertSame(self::GRAPH.'/me/calendars/AAMkAGI2-default%3D/calendarView', $http->requests[0]['url']);
        self::assertSame(['startDateTime' => '2026-11-16T00:00:00Z', 'endDateTime' => '2026-12-31T00:00:00Z', '$top' => 10, '$orderby' => 'start/dateTime'], $http->requests[0]['options']['query'] ?? null);
        self::assertStringStartsWith(self::GRAPH.'/me/calendars/AAMkAGI2-default=/calendarView?startDateTime=', $http->requests[1]['url']);
        self::assertSame('Bearer at-1', $http->header('Authorization', 1));

        [$occurrence, $meeting, $closed, $cancelled] = $events;
        self::assertSame(['Point hebdomadaire', '2026-11-16T08:00:00+01:00', 'Europe/Paris', 'AAMkSeries-0009', '2026-11-16T07:00:00+00:00', 'ck-0001'], [$occurrence->title, $occurrence->start->format(\DATE_ATOM), $occurrence->timeZone, $occurrence->recurringEventId, $occurrence->originalStart?->format(\DATE_ATOM), $occurrence->etag]);
        self::assertTrue($occurrence->isOccurrence());
        self::assertSame(['Dossier Durand', 'Cabinet, salle 2', 2700, self::CALENDAR], [$meeting->description, $meeting->location, $meeting->duration(), $meeting->calendarId]);
        self::assertSame(['anna@contoso.example', true, AttendeeRole::Chair], [$meeting->organizer?->email, $meeting->organizer?->organizer, $meeting->organizer?->role]);
        self::assertSame([
            ['paul.durand@fabrikam.example', ParticipationStatus::Accepted, AttendeeRole::Required, false],
            ['lea@fabrikam.example', ParticipationStatus::Tentative, AttendeeRole::Optional, false],
            ['salle2@contoso.example', ParticipationStatus::NeedsAction, AttendeeRole::NonParticipant, true],
        ], array_map(static fn (Attendee $a) => [$a->email, $a->status, $a->role, $a->resource], $meeting->attendees));
        self::assertSame([true, '2026-12-24T00:00:00+01:00', '2026-12-27T00:00:00+01:00', Transparency::Transparent], [$closed->allDay, $closed->start->format(\DATE_ATOM), $closed->end->format(\DATE_ATOM), $closed->transparency]);
        self::assertSame([EventStatus::Cancelled, 'UTC', false], [$cancelled->status, $cancelled->timeZone, $cancelled->busy()], 'a custom Outlook zone is read as UTC');
        self::assertSame(['040000008200E00074C5B7101A82E008-0002', '2026-11-10T09:12:44+00:00'], [$meeting->uid, $meeting->updated?->format(\DATE_ATOM)]);
    }

    public function testTheLimitStopsThePagingAndWordsAreMatchedHere(): void
    {
        $http = self::http();
        $query = static fn (int $limit, ?string $search = null) => new EventQuery(self::CALENDAR, new \DateTimeImmutable('2026-11-16'), new \DateTimeImmutable('2026-12-31'), limit: $limit, search: $search);
        self::assertCount(2, self::provider($http)->events($query(2)));
        self::assertCount(1, $http->requests);
        self::assertSame(2, self::option($http, 0, 'query', '$top'));

        $found = self::provider($http)->events($query(1, 'reportée'));
        self::assertSame(['Audience reportée'], array_map(static fn (Event $e) => $e->title, $found), 'found on the second page');
    }

    public function testSeriesWithoutExpandingAndOneEventWithItsRule(): void
    {
        $http = self::http();
        $provider = self::provider($http);
        $provider->events(new EventQuery(self::CALENDAR, new \DateTimeImmutable('2026-11-16T00:00:00Z'), new \DateTimeImmutable('2026-12-31T00:00:00Z'), expand: false));
        self::assertSame(self::GRAPH.'/me/calendars/AAMkAGI2-default%3D/events', $http->requests[0]['url']);
        self::assertSame("start/dateTime lt '2026-12-31T00:00:00' and end/dateTime gt '2026-11-16T00:00:00'", self::option($http, 0, 'query', '$filter'));

        $series = $provider->event(self::CALENDAR, 'AAMkSeries-0009');
        self::assertSame(self::GRAPH.'/me/events/AAMkSeries-0009', $http->requests[1]['url']);
        self::assertSame(['FREQ=WEEKLY;BYDAY=MO;WKST=SU;UNTIL=20261228T225959Z'], $series?->recurrence?->rules);
        self::assertSame('2026-11-02T08:00:00+01:00', $series->start->format(\DATE_ATOM));
        self::assertNull($provider->event(self::CALENDAR, 'AAMkGone-0000'), 'a 404 is no event');
    }

    public function testAnEventIsCreatedInItsZoneWithItsRule(): void
    {
        $http = self::http();
        $paris = new \DateTimeZone('Europe/Paris');
        $event = new Event('Consultation', new \DateTimeImmutable('2026-11-23T10:00', $paris), new \DateTimeImmutable('2026-11-23T10:45', $paris), calendarId: self::CALENDAR, description: 'Premier rendez-vous',
            attendees: [new Attendee('paul.durand@fabrikam.example', 'Paul Durand'), new Attendee('salle2@contoso.example', resource: true), new Attendee('anna@contoso.example', organizer: true)],
            recurrence: Recurrence::rule('FREQ=WEEKLY;COUNT=4'));
        $created = self::provider($http)->create($event);

        self::assertSame(['POST', self::GRAPH.'/me/calendars/AAMkAGI2-default%3D/events'], [$http->requests[0]['method'], $http->requests[0]['url']]);
        self::assertSame([
            'subject' => 'Consultation',
            'body' => ['contentType' => 'text', 'content' => 'Premier rendez-vous'],
            'start' => ['dateTime' => '2026-11-23T10:00:00', 'timeZone' => 'Europe/Paris'],
            'end' => ['dateTime' => '2026-11-23T10:45:00', 'timeZone' => 'Europe/Paris'],
            'isAllDay' => false,
            'location' => ['displayName' => ''],
            'showAs' => 'busy',
            'attendees' => [
                ['emailAddress' => ['address' => 'paul.durand@fabrikam.example', 'name' => 'Paul Durand'], 'type' => 'required'],
                ['emailAddress' => ['address' => 'salle2@contoso.example'], 'type' => 'resource'],
            ],
            'recurrence' => [
                'pattern' => ['type' => 'weekly', 'interval' => 1, 'daysOfWeek' => ['monday'], 'firstDayOfWeek' => 'monday'],
                'range' => ['type' => 'numbered', 'numberOfOccurrences' => 4, 'startDate' => '2026-11-23', 'recurrenceTimeZone' => 'Europe/Paris'],
            ],
        ], $http->requests[0]['options']['json'] ?? null);
        self::assertSame(['AAMkNew-0005', 'ck-0005', 'Europe/Paris', '2026-11-23T10:00:00+01:00'], [$created->id, $created->etag, $created->timeZone, $created->start->format(\DATE_ATOM)]);
    }

    public function testAnAllDayEventIsWrittenAsMidnightsInOneZone(): void
    {
        $body = MicrosoftProvider::write(new Event('Fermeture', new \DateTimeImmutable('2026-12-24'), new \DateTimeImmutable('2026-12-27'), allDay: true, timeZone: 'Europe/Paris', transparency: Transparency::Transparent));
        self::assertSame([['dateTime' => '2026-12-24T00:00:00', 'timeZone' => 'Europe/Paris'], ['dateTime' => '2026-12-27T00:00:00', 'timeZone' => 'Europe/Paris'], true, 'free'], [$body['start'], $body['end'], $body['isAllDay'], $body['showAs']]);
    }

    public function testAStaleChangeKeyIsAConflictAndNothingIsWritten(): void
    {
        $http = self::http();
        $provider = self::provider($http);
        $meeting = $provider->event(self::CALENDAR, 'AAMkMeet-0002');
        self::assertNotNull($meeting);
        try {
            $provider->update($meeting->stored('AAMkMeet-0002', self::CALENDAR, 'ck-old'));
            self::fail('A stale changeKey was written over.');
        } catch (ConflictException $e) {
            self::assertSame(412, $e->status());
        }
        self::assertSame(['GET', 'GET'], array_column($http->requests, 'method'), 'read, compared, nothing patched');

        $updated = $provider->update($meeting->moved(new \DateTimeImmutable('2026-11-16T11:00:00Z'), new \DateTimeImmutable('2026-11-16T11:45:00Z')));
        self::assertSame(['GET', 'PATCH'], array_column(\array_slice($http->requests, 2), 'method'));
        self::assertSame(self::GRAPH.'/me/events/AAMkMeet-0002', $http->requests[3]['url']);
        self::assertSame(['dateTime' => '2026-11-16T12:00:00', 'timeZone' => 'Europe/Paris'], self::option($http, 3, 'json', 'start'));
        self::assertSame('AAMkMeet-0002', $updated->id);
    }

    public function testDeletingWithAndWithoutTheChangeKey(): void
    {
        $http = self::http();
        $provider = self::provider($http);
        $provider->delete(self::CALENDAR, 'AAMkMeet-0002', 'ck-0002');
        self::assertSame([['GET', self::GRAPH.'/me/events/AAMkMeet-0002'], ['DELETE', self::GRAPH.'/me/events/AAMkMeet-0002']], array_map(static fn (array $r) => [$r['method'], $r['url']], $http->requests));
        $provider->delete(self::CALENDAR, 'AAMkGone-0000', 'ck-x');
        self::assertCount(3, $http->requests, 'gone already: nothing to delete');
        $provider->delete(self::CALENDAR, 'AAMkGone-0000');
        self::assertSame('DELETE', $http->requests[3]['method'], 'a 404 on DELETE is no error');
    }

    public function testPeoplesBusyTimeThroughGetSchedule(): void
    {
        $http = self::http();
        $answers = self::provider($http)->freeBusy(new FreeBusyQuery(['anna@contoso.example', 'paul@fabrikam.example', 'Lea@contoso.example', 'AAMkAGI2-shared='], new \DateTimeImmutable('2026-11-16T08:00', new \DateTimeZone('Europe/Paris')), new \DateTimeImmutable('2026-11-16T18:00', new \DateTimeZone('Europe/Paris'))));

        self::assertSame(['POST', self::GRAPH.'/me/calendar/getSchedule'], [$http->requests[0]['method'], $http->requests[0]['url']]);
        self::assertSame([
            'schedules' => ['anna@contoso.example', 'paul@fabrikam.example', 'Lea@contoso.example'],
            'startTime' => ['dateTime' => '2026-11-16T07:00:00', 'timeZone' => 'UTC'],
            'endTime' => ['dateTime' => '2026-11-16T17:00:00', 'timeZone' => 'UTC'],
            'availabilityViewInterval' => 30,
        ], $http->requests[0]['options']['json'] ?? null);
        self::assertSame(['09:00-10:00 busy', '11:00-11:30 tentative'], array_map(static fn (BusySlot $s) => $s->start->format('H:i').'-'.$s->end->format('H:i').' '.$s->status->value, $answers[0]->slots), 'working elsewhere is free');
        self::assertSame('Unable to retrieve free/busy information for the user.', $answers[1]->error);
        self::assertSame([FreeBusyStatus::OutOfOffice], array_map(static fn (BusySlot $s) => $s->status, $answers[2]->slots), 'matched without case');
        self::assertStringContainsString('by address', (string) $answers[3]->error);
    }

    public function testA401RefreshesTheTokenOnceAndKeepsTheRotatedOne(): void
    {
        $calls = 0;
        $http = new RecordedHttpClient(static function (string $method, string $url, array $options) use (&$calls): array {
            if (str_contains($url, 'login.microsoftonline.com')) {
                return [200, '{"token_type":"Bearer","scope":"Calendars.ReadWrite","expires_in":3600,"access_token":"at-2","refresh_token":"rt-2"}'];
            }

            return 1 === ++$calls ? [401, '{"error":{"code":"InvalidAuthenticationToken","message":"Access token has expired."}}'] : [200, (string) file_get_contents(__DIR__.'/Fixtures/calendars.json')];
        });
        $store = new InMemoryTokenStore(['user:42:microsoft' => new OAuthToken('at-1', 'rt-1', 2000000)]);
        $provider = (new MicrosoftFactory($http, $store, new ArraySecrets(['MICROSOFT_CLIENT_SECRET' => 'not-a-real-secret']), static fn () => 1000000))
            ->create(['client_id' => 'client-1', 'client_secret' => 'MICROSOFT_CLIENT_SECRET', 'tenant' => 'organizations', 'token' => 'user:42:microsoft']);
        self::assertInstanceOf(MicrosoftProvider::class, $provider);

        self::assertCount(2, $provider->calendars());
        self::assertSame(['GET', 'POST', 'GET'], array_column($http->requests, 'method'));
        self::assertSame('https://login.microsoftonline.com/organizations/oauth2/v2.0/token', $http->requests[1]['url']);
        self::assertSame(['grant_type' => 'refresh_token', 'refresh_token' => 'rt-1', 'client_id' => 'client-1', 'client_secret' => 'not-a-real-secret'], $http->requests[1]['options']['body'] ?? null);
        self::assertSame('Bearer at-2', $http->header('Authorization'));
        self::assertSame(['at-2', 'rt-2', 1003600], [$store->get('user:42:microsoft')?->accessToken, $store->get('user:42:microsoft')?->refreshToken, $store->get('user:42:microsoft')?->expiresAt]);
    }

    public function testATokenRefusedTwiceAsksForConsent(): void
    {
        $http = RecordedHttpClient::routes(['~~' => [401, '{"error":{"code":"InvalidAuthenticationToken"}}']]);
        $provider = (new MicrosoftFactory($http, null, new ArraySecrets(['MS_TOKEN' => 'not-a-real-token'])))->create(['access_token' => 'MS_TOKEN']);
        self::assertInstanceOf(MicrosoftProvider::class, $provider);
        try {
            $provider->calendars();
            self::fail('A refused token read calendars.');
        } catch (AuthenticationException $e) {
            self::assertStringContainsString('consent again', $e->getMessage());
            self::assertStringNotContainsString('not-a-real-token', $e->getMessage());
        }
        self::assertCount(1, $http->requests);
    }

    public function testANextLinkOffGraphIsNotFollowed(): void
    {
        $http = RecordedHttpClient::routes(['~/calendarView~' => [200, '{"value":[],"@odata.nextLink":"https://elsewhere.example/steal"}']]);
        self::provider($http)->events(new EventQuery(self::CALENDAR, new \DateTimeImmutable('2026-11-16'), new \DateTimeImmutable('2026-11-17')));
        self::assertCount(1, $http->requests);
        self::assertFalse((new Graph(...self::graphParts($http)))->owns('https://elsewhere.example/steal'));
    }

    public function testTheFactorysOptionsAndTheConsentStep(): void
    {
        $factory = new MicrosoftFactory(RecordedHttpClient::routes([]), new InMemoryTokenStore());
        $url = $factory->oauth(['client_id' => 'client-1', 'tenant' => 'consumers'])->authorizationUrl('https://site.example/back', MicrosoftFactory::scopes(), 'st', 'challenge');
        self::assertStringStartsWith('https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?response_type=code&client_id=client-1&redirect_uri=https%3A%2F%2Fsite.example%2Fback&scope=offline_access%20https%3A%2F%2Fgraph.microsoft.com%2FCalendars.ReadWrite&state=st&code_challenge=challenge&code_challenge_method=S256', $url);
        self::assertSame(['offline_access', 'https://graph.microsoft.com/Calendars.Read'], MicrosoftFactory::scopes(false));

        foreach ([
            [[], 'needs the key of the user\'s token'],
            [['token' => 'u', 'tenant' => 'a/b', 'client_id' => 'c'], 'is no tenant'],
            [['token' => 'u'], 'needs its application\'s client_id'],
            [['access_token' => 'X', 'base_uri' => 'http://graph.example'], 'needs an https base_uri'],
            [['access_token' => 'X', 'availability_interval' => 2], '5 to 1440 minutes'],
        ] as [$options, $message]) {
            try {
                $factory->create($options);
                self::fail('Built with '.json_encode($options));
            } catch (InvalidConfigException $e) {
                self::assertStringContainsString($message, $e->getMessage());
            }
        }
        $this->expectExceptionMessage('token store');
        (new MicrosoftFactory(RecordedHttpClient::routes([])))->create(['token' => 'u', 'client_id' => 'c']);
    }

    public function testWindowsZonesAreReadAsIana(): void
    {
        self::assertSame(['Europe/Paris', 'Asia/Tokyo', 'America/New_York', 'Europe/Paris', 'UTC', null, null], array_map(Zones::toIana(...), ['Romance Standard Time', 'Tokyo Standard Time', 'Eastern Standard Time', 'Europe/Paris', 'UTC', 'tzone://Microsoft/Custom', 'Nowhere Standard Time']));
        foreach (Zones::WINDOWS as $windows => $iana) {
            self::assertNotNull(Zones::toIana($iana), $iana.' is a zone PHP knows');
        }
    }

    private static function provider(RecordedHttpClient $http): MicrosoftProvider
    {
        $provider = (new MicrosoftFactory($http, new InMemoryTokenStore(['user:42:microsoft' => new OAuthToken('at-1', 'rt-1', 2000000)]), new ArraySecrets([]), static fn () => 1000000))
            ->create(['client_id' => 'client-1', 'token' => 'user:42:microsoft']);

        return $provider instanceof MicrosoftProvider ? $provider : throw new \LogicException('No Microsoft provider.');
    }

    /** @return array{0: \Omnical\Http\JsonClient, 1: \Omnical\Http\RawClient, 2: \Omnical\Auth\StaticAccessToken} */
    private static function graphParts(RecordedHttpClient $http): array
    {
        return [new \Omnical\Http\JsonClient($http, 'Microsoft Graph'), new \Omnical\Http\RawClient($http, 'Microsoft Graph'), new \Omnical\Auth\StaticAccessToken(new ArraySecrets(['T' => 'x']), 'T')];
    }

    private static function option(RecordedHttpClient $http, int $request, string $option, string $key): mixed
    {
        $value = $http->requests[$request]['options'][$option] ?? null;

        return \is_array($value) ? ($value[$key] ?? null) : null;
    }

    /** The second event of the first calendarView page, alone. */
    private static function meeting(): string
    {
        return (string) json_encode(Data::of(json_decode((string) file_get_contents(__DIR__.'/Fixtures/calendarview-1.json'), true))->list('value')[1]->toArray());
    }

    private static function http(): RecordedHttpClient
    {
        $fixture = static fn (string $name) => '@'.__DIR__.'/Fixtures/'.$name.'.json';

        return RecordedHttpClient::routes([
            '~^GET https://graph\.microsoft\.com/v1\.0/me/calendars\?~' => [200, $fixture('calendars')],
            '~^GET https://graph\.microsoft\.com/v1\.0/me/calendars/AAMkAGI2-default%3D/calendarView\?~' => [200, $fixture('calendarview-1')],
            '~^GET https://graph\.microsoft\.com/v1\.0/me/calendars/AAMkAGI2-default=/calendarView\?.*%24skip=3~' => [200, $fixture('calendarview-2')],
            '~^GET https://graph\.microsoft\.com/v1\.0/me/calendars/AAMkAGI2-default%3D/events\?~' => [200, '{"value":[]}'],
            '~^GET https://graph\.microsoft\.com/v1\.0/me/events/AAMkSeries-0009$~' => [200, $fixture('event-series')],
            '~^GET https://graph\.microsoft\.com/v1\.0/me/events/AAMkMeet-0002$~' => [200, self::meeting()],
            '~^PATCH https://graph\.microsoft\.com/v1\.0/me/events/AAMkMeet-0002$~' => [200, self::meeting()],
            '~^DELETE https://graph\.microsoft\.com/v1\.0/me/events/AAMkGone~' => [404, '{"error":{"code":"ErrorItemNotFound"}}'],
            '~^DELETE https://graph\.microsoft\.com/v1\.0/me/events/~' => [204, ''],
            '~^POST https://graph\.microsoft\.com/v1\.0/me/calendars/AAMkAGI2-default%3D/events$~' => [201, $fixture('event-created')],
            '~^POST https://graph\.microsoft\.com/v1\.0/me/calendar/getSchedule$~' => [200, $fixture('schedule')],
        ]);
    }
}
