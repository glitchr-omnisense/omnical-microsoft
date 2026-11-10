<?php

namespace Omnical\Microsoft;

use Omnical\CalendarEventsInterface;
use Omnical\Capabilities;
use Omnical\Exception\ConflictException;
use Omnical\Exception\InvalidArgumentException;
use Omnical\FreeBusyInterface;
use Omnical\Http\Data;
use Omnical\Model\Attendee;
use Omnical\Model\AttendeeRole;
use Omnical\Model\BusySlot;
use Omnical\Model\Calendar;
use Omnical\Model\Event;
use Omnical\Model\EventQuery;
use Omnical\Model\EventStatus;
use Omnical\Model\FreeBusy;
use Omnical\Model\FreeBusyQuery;
use Omnical\Model\FreeBusyStatus;
use Omnical\Model\ParticipationStatus;
use Omnical\Model\Transparency;
use Omnisense\Contracts\Policy\DeclaresUsagePolicyInterface;
use Omnisense\Contracts\Policy\UsagePolicy;

/**
 * Microsoft Graph's calendars (v1.0), the signed-in user's ("/me"): Outlook
 * and Microsoft 365, work or school accounts and personal ones.
 *
 * - calendars(): GET /me/calendars (canEdit, isDefaultCalendar, hexColor, owner), every page.
 * - events(): with EventQuery::$expand, GET /me/calendars/{id}/calendarView (startDateTime,
 *   endDateTime, $top up to 1000, $orderby=start/dateTime, @odata.nextLink up to the query's
 *   limit): each occurrence on its own; without, GET /me/calendars/{id}/events filtered on the
 *   events' own start and end ($filter=start/dateTime lt '…' and end/dateTime gt '…'): single
 *   events and series masters - a series whose first occurrence lies before the span is not
 *   among them. $search is not sent: the query's words are matched here, on the title,
 *   description and location.
 * - event(): GET /me/events/{id} (404: null).
 * - create(): POST /me/calendars/{id}/events. update(): PATCH /me/events/{id}. delete(): DELETE
 *   /me/events/{id} (204; a 404 is nothing to delete).
 * - freeBusy(): POST /me/calendar/getSchedule, people by their address (work or school accounts
 *   only; personal Microsoft accounts have no getSchedule): each schedule's items as busy spans,
 *   its error as FreeBusy::$error.
 *
 * Every request says Prefer: outlook.timezone="UTC" (times read as UTC, then set in the event's
 * zone, its originalStartTimeZone read as IANA by Zones), IdType="ImmutableId" (ids survive a
 * move between calendars) and outlook.body-content-type="text". The event's version is its
 * changeKey ($etag): Graph documents no If-Match on events, so update() and delete() given an
 * etag read the event first and throw ConflictException when its changeKey changed - a check
 * read then written, not atomic. Recurrence goes through PatternedRecurrence.
 */
final class MicrosoftProvider implements CalendarEventsInterface, FreeBusyInterface, DeclaresUsagePolicyInterface
{
    /** Microsoft Graph's usage policy, as its terms and throttling limits read on the day it says (provenance.verifiedAt). */
    public const POLICY = __DIR__.'/policy.json';

    /** The largest page calendarView gives ($top). */
    public const PAGE = 1000;

    private static ?UsagePolicy $policy = null;

    public function __construct(private readonly Graph $graph, private readonly int $availabilityInterval = 30)
    {
    }

    public function getName(): string
    {
        return 'microsoft';
    }

    public function getTitle(): string
    {
        return 'Microsoft Graph';
    }

    public function usagePolicy(): UsagePolicy
    {
        return self::$policy ??= UsagePolicy::fromFile(self::POLICY);
    }

    public function capabilities(): Capabilities
    {
        return new Capabilities(calendarEvents: true, write: true, expand: true, attendees: true, freeBusy: true, people: true, policy: $this->usagePolicy());
    }

    public function calendars(): array
    {
        $calendars = [];
        foreach ($this->pages('/me/calendars', ['$top' => 100], \PHP_INT_MAX) as $item) {
            $id = $item->string('id');
            if (null === $id) {
                continue;
            }
            $calendars[] = new Calendar(
                $id,
                $item->string('name') ?? $id,
                'microsoft',
                color: $item->string('hexColor'),
                readOnly: false === $item->bool('canEdit'),
                primary: true === $item->bool('isDefaultCalendar'),
                owner: $item->string('owner.address'),
                version: $item->string('changeKey'),
            );
        }
        usort($calendars, static fn (Calendar $a, Calendar $b) => (int) $b->primary <=> (int) $a->primary);

        return $calendars;
    }

    public function events(EventQuery $query): array
    {
        $from = self::utc($query->from);
        $to = self::utc($query->to);
        $calendar = '/me/calendars/'.rawurlencode($query->calendarId);
        $params = $query->expand
            ? ['startDateTime' => $from, 'endDateTime' => $to, '$top' => min(self::PAGE, $query->limit), '$orderby' => 'start/dateTime']
            : ['$filter' => \sprintf("start/dateTime lt '%s' and end/dateTime gt '%s'", substr($to, 0, -1), substr($from, 0, -1)), '$top' => min(self::PAGE, $query->limit), '$orderby' => 'start/dateTime'];
        $events = [];
        foreach ($this->pages($calendar.($query->expand ? '/calendarView' : '/events'), $params, $query->limit, null !== $query->search) as $item) {
            $event = $this->read($item, $query->calendarId);
            if (null !== $event && (null === $query->search || str_contains(mb_strtolower($event->title.' '.$event->description.' '.$event->location), mb_strtolower($query->search)))) {
                $events[] = $event;
            }
            if (\count($events) >= $query->limit) {
                break;
            }
        }

        return $events;
    }

    public function event(string $calendarId, string $eventId): ?Event
    {
        $data = $this->graph->json('GET', '/me/events/'.rawurlencode($eventId));

        return null === $data ? null : $this->read($data, $calendarId);
    }

    public function create(Event $event): Event
    {
        $calendar = $event->calendarId ?? throw new InvalidArgumentException('An event is created in a calendar: its calendarId.');
        $data = $this->graph->json('POST', '/me/calendars/'.rawurlencode($calendar).'/events', ['json' => self::write($event)]);

        return (null === $data ? null : $this->read($data, $calendar)) ?? throw new InvalidArgumentException('Microsoft Graph gave back no event.');
    }

    public function update(Event $event): Event
    {
        $id = $event->id ?? throw new InvalidArgumentException('An event is updated by its id.');
        $calendar = $event->calendarId ?? throw new InvalidArgumentException('An event is updated in its calendar: its calendarId.');
        $this->unchanged($calendar, $id, $event->etag);
        $data = $this->graph->json('PATCH', '/me/events/'.rawurlencode($id), ['json' => self::write($event)]);

        return (null === $data ? null : $this->read($data, $calendar)) ?? throw new InvalidArgumentException(\sprintf('Microsoft Graph has no event "%s" to update.', $id));
    }

    public function delete(string $calendarId, string $eventId, ?string $etag = null): void
    {
        if (null !== $etag && !$this->unchanged($calendarId, $eventId, $etag, false)) {
            return;
        }
        $this->graph->raw('DELETE', '/me/events/'.rawurlencode($eventId));
    }

    public function freeBusy(FreeBusyQuery $query): array
    {
        $people = array_values(array_filter($query->calendars, static fn (string $c) => str_contains($c, '@')));
        $answers = [];
        if ([] !== $people) {
            $data = $this->graph->json('POST', '/me/calendar/getSchedule', ['json' => [
                'schedules' => $people,
                'startTime' => ['dateTime' => substr(self::utc($query->from), 0, -1), 'timeZone' => 'UTC'],
                'endTime' => ['dateTime' => substr(self::utc($query->to), 0, -1), 'timeZone' => 'UTC'],
                'availabilityViewInterval' => $this->availabilityInterval,
            ]]);
            foreach ($data?->list('value') ?? [] as $schedule) {
                $answers[strtolower((string) $schedule->string('scheduleId'))] = $schedule;
            }
        }
        $result = [];
        foreach ($query->calendars as $calendar) {
            $schedule = $answers[strtolower($calendar)] ?? null;
            if (!str_contains($calendar, '@')) {
                $result[] = new FreeBusy($calendar, [], 'Microsoft Graph tells busy time by address (getSchedule): give the person\'s address.', 'microsoft');
                continue;
            }
            $slots = [];
            foreach ($schedule?->list('scheduleItems') ?? [] as $item) {
                $start = self::moment($item->data('start'));
                $end = self::moment($item->data('end'));
                if (null !== $start && null !== $end && $end >= $start) {
                    $slots[] = new BusySlot($start, $end, FreeBusyStatus::read($item->string('status')));
                }
            }
            $error = null === $schedule ? 'Microsoft Graph gave no schedule for this address.' : $schedule->string('error.message');
            $result[] = new FreeBusy($calendar, $slots, $error, 'microsoft');
        }

        return $result;
    }

    /**
     * An event as Graph sends it, read into the model.
     */
    public function read(Data $item, string $calendarId): ?Event
    {
        $id = $item->string('id');
        $start = self::moment($item->data('start'));
        $end = self::moment($item->data('end'));
        if (null === $id || null === $start || null === $end) {
            return null;
        }
        $allDay = true === $item->bool('isAllDay');
        $zone = Zones::toIana($item->string('originalStartTimeZone')) ?? Zones::toIana($item->string('start.timeZone')) ?? 'UTC';
        $in = new \DateTimeZone($zone);
        if ($allDay) {
            $start = new \DateTimeImmutable($start->format('Y-m-d'), $in);
            $end = new \DateTimeImmutable($end->format('Y-m-d'), $in);
        } else {
            $start = $start->setTimezone($in);
            $end = $end->setTimezone($in);
        }
        $organizer = null === ($address = $item->string('organizer.emailAddress.address')) ? null
            : new Attendee($address, $item->string('organizer.emailAddress.name'), ParticipationStatus::Accepted, AttendeeRole::Chair, organizer: true);
        $attendees = [];
        foreach ($item->list('attendees') as $attendee) {
            $email = $attendee->string('emailAddress.address');
            if (null === $email) {
                continue;
            }
            $type = $attendee->string('type');
            $attendees[] = new Attendee(
                $email,
                $attendee->string('emailAddress.name'),
                ParticipationStatus::read($attendee->string('status.response')),
                'optional' === $type ? AttendeeRole::Optional : ('resource' === $type ? AttendeeRole::NonParticipant : AttendeeRole::Required),
                resource: 'resource' === $type,
            );
        }
        $recurrence = $item->data('recurrence');
        $type = $item->string('type');
        $showAs = $item->string('showAs');

        return new Event(
            $item->string('subject') ?? '',
            $start,
            $end,
            $allDay,
            $id,
            $calendarId,
            $zone,
            $item->string('body.content') ?? $item->string('bodyPreview'),
            $item->string('location.displayName'),
            true === $item->bool('isCancelled') ? EventStatus::Cancelled : ('tentative' === $showAs ? EventStatus::Tentative : EventStatus::Confirmed),
            'free' === $showAs ? Transparency::Transparent : Transparency::Opaque,
            $organizer,
            $attendees,
            [] === $recurrence->toArray() ? null : PatternedRecurrence::toRecurrence($recurrence, $allDay, $zone),
            \in_array($type, ['occurrence', 'exception'], true) ? $item->string('seriesMasterId') : null,
            $item->dateTime('originalStart'),
            $item->string('iCalUId'),
            $item->string('webLink'),
            $item->string('changeKey'),
            $item->dateTime('lastModifiedDateTime'),
            'microsoft',
        );
    }

    /**
     * An event written as Graph's resource.
     *
     * @return array<string, mixed>
     */
    public static function write(Event $event): array
    {
        $zone = $event->timeZone ?? 'UTC';
        $in = new \DateTimeZone($zone);
        $moment = static fn (\DateTimeImmutable $m) => ['dateTime' => $event->allDay ? $m->format('Y-m-d').'T00:00:00' : $m->setTimezone($in)->format('Y-m-d\TH:i:s'), 'timeZone' => $zone];
        $body = [
            'subject' => $event->title,
            'body' => ['contentType' => 'text', 'content' => $event->description ?? ''],
            'start' => $moment($event->start),
            'end' => $moment($event->end),
            'isAllDay' => $event->allDay,
            'location' => ['displayName' => $event->location ?? ''],
            'showAs' => Transparency::Transparent === $event->transparency ? 'free' : (EventStatus::Tentative === $event->status ? 'tentative' : 'busy'),
            'attendees' => array_values(array_map(static fn (Attendee $a) => [
                'emailAddress' => array_filter(['address' => $a->email, 'name' => $a->name], static fn (?string $v) => null !== $v),
                'type' => $a->resource ? 'resource' : (AttendeeRole::Optional === $a->role ? 'optional' : 'required'),
            ], array_filter($event->attendees, static fn (Attendee $a) => !$a->organizer))),
        ];
        if (null !== $event->recurrence && !$event->isOccurrence()) {
            $body['recurrence'] = PatternedRecurrence::fromRecurrence($event->recurrence, $event);
        }

        return $body;
    }

    /**
     * Throws ConflictException when the event's changeKey is not the one
     * given; false when the event is gone and $required is false.
     */
    private function unchanged(string $calendarId, string $eventId, ?string $etag, bool $required = true): bool
    {
        if (null === $etag) {
            return true;
        }
        $current = $this->event($calendarId, $eventId);
        if (null === $current) {
            return $required ? throw new ConflictException(\sprintf('Microsoft Graph has no event "%s" any more: it was deleted since it was read.', $eventId), 412) : false;
        }
        if ($current->etag !== $etag) {
            throw new ConflictException(\sprintf('Microsoft Graph\'s event "%s" changed since it was read (its changeKey): read it again before writing.', $eventId), 412);
        }

        return true;
    }

    /**
     * Every item of every page, up to $limit, following @odata.nextLink on
     * Graph's own host.
     *
     * @param array<string, string|int> $query
     *
     * @return \Generator<int, Data>
     */
    private function pages(string $path, array $query, int $limit, bool $all = false): \Generator
    {
        $seen = 0;
        $data = $this->graph->json('GET', $path, ['query' => $query]);
        while (null !== $data) {
            foreach ($data->list('value') as $item) {
                yield $item;
                if (!$all && ++$seen >= $limit) {
                    return;
                }
            }
            $next = $data->string('@odata.nextLink');
            if (null === $next || !$this->graph->owns($next)) {
                return;
            }
            $data = $this->graph->json('GET', $next);
        }
    }

    /** A dateTimeTimeZone ("2026-11-16T09:00:00.0000000", "UTC") as a moment. */
    private static function moment(Data $value): ?\DateTimeImmutable
    {
        $text = $value->string('dateTime');
        if (null === $text || !preg_match('~^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2})?)~', $text, $m)) {
            return null;
        }
        $zone = new \DateTimeZone(Zones::toIana($value->string('timeZone')) ?? 'UTC');

        return \DateTimeImmutable::createFromFormat(16 === \strlen($m[1]) ? 'Y-m-d\TH:i' : 'Y-m-d\TH:i:s', $m[1], $zone) ?: null;
    }

    private static function utc(\DateTimeImmutable $moment): string
    {
        return $moment->setTimezone(new \DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z');
    }
}
