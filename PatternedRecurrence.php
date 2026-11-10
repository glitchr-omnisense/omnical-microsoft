<?php

namespace Omnical\Microsoft;

use Omnical\Exception\NotSupportedException;
use Omnical\Http\Data;
use Omnical\Model\Event;
use Omnical\Model\Recurrence;

/**
 * Graph's patternedRecurrence (a pattern and a range) and iCalendar's RRULE,
 * both ways:
 *
 *   daily               FREQ=DAILY
 *   weekly              FREQ=WEEKLY;BYDAY=MO,WE (firstDayOfWeek as WKST, unless Monday)
 *   absoluteMonthly     FREQ=MONTHLY;BYMONTHDAY=15
 *   relativeMonthly     FREQ=MONTHLY;BYDAY=2TU (one day), or BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1 (several)
 *   absoluteYearly      FREQ=YEARLY;BYMONTH=3;BYMONTHDAY=8
 *   relativeYearly      FREQ=YEARLY;BYMONTH=11;BYDAY=4TH
 *   range endDate       UNTIL (the last day's end in the recurrence's zone, in UTC; a date for an all-day event)
 *   range numbered      COUNT
 *   range noEnd         neither
 *
 * index first, second, third, fourth, last is 1, 2, 3, 4, -1. A rule Graph
 * cannot say - another frequency, BYHOUR and the like, several rules,
 * EXRULE, RDATE or EXDATE (Graph takes an occurrence out by deleting it) -
 * is a NotSupportedException rather than another rule written in its place.
 */
final class PatternedRecurrence
{
    private const DAYS = ['MO' => 'monday', 'TU' => 'tuesday', 'WE' => 'wednesday', 'TH' => 'thursday', 'FR' => 'friday', 'SA' => 'saturday', 'SU' => 'sunday'];
    private const INDEX = ['first' => 1, 'second' => 2, 'third' => 3, 'fourth' => 4, 'last' => -1];
    private const PARTS = ['FREQ', 'INTERVAL', 'BYDAY', 'BYMONTHDAY', 'BYMONTH', 'BYSETPOS', 'WKST', 'UNTIL', 'COUNT'];

    /**
     * Graph's recurrence read as a Recurrence; null for a pattern it does
     * not know.
     *
     * @param string $zone the event's IANA zone, for a range read in no zone of its own
     */
    public static function toRecurrence(Data $recurrence, bool $allDay, string $zone): ?Recurrence
    {
        $pattern = $recurrence->data('pattern');
        $range = $recurrence->data('range');
        $days = array_values(array_filter(array_map(static fn (string $d) => array_search(strtolower($d), self::DAYS, true) ?: null, $pattern->strings('daysOfWeek'))));
        $index = self::INDEX[$pattern->string('index') ?? 'first'] ?? 1;
        $parts = match ($pattern->string('type')) {
            'daily' => ['FREQ' => 'DAILY'],
            'weekly' => ['FREQ' => 'WEEKLY', 'BYDAY' => implode(',', $days)],
            'absoluteMonthly' => ['FREQ' => 'MONTHLY', 'BYMONTHDAY' => (string) ($pattern->int('dayOfMonth') ?? 1)],
            'relativeMonthly' => ['FREQ' => 'MONTHLY'] + self::relative($days, $index),
            'absoluteYearly' => ['FREQ' => 'YEARLY', 'BYMONTH' => (string) ($pattern->int('month') ?? 1), 'BYMONTHDAY' => (string) ($pattern->int('dayOfMonth') ?? 1)],
            'relativeYearly' => ['FREQ' => 'YEARLY', 'BYMONTH' => (string) ($pattern->int('month') ?? 1)] + self::relative($days, $index),
            default => null,
        };
        if (null === $parts) {
            return null;
        }
        $interval = $pattern->int('interval') ?? 1;
        if ($interval > 1) {
            $parts = ['FREQ' => $parts['FREQ'], 'INTERVAL' => (string) $interval] + $parts;
        }
        $first = strtolower((string) $pattern->string('firstDayOfWeek'));
        if ('WEEKLY' === $parts['FREQ'] && '' !== $first && 'monday' !== $first && false !== $wkst = array_search($first, self::DAYS, true)) {
            $parts['WKST'] = $wkst;
        }
        if ('numbered' === $range->string('type') && null !== $count = $range->int('numberOfOccurrences')) {
            $parts['COUNT'] = (string) $count;
        } elseif ('endDate' === $range->string('type') && null !== $end = $range->string('endDate')) {
            $in = new \DateTimeZone(Zones::toIana($range->string('recurrenceTimeZone')) ?? $zone);
            $until = \DateTimeImmutable::createFromFormat('!Y-m-d H:i:s', substr($end, 0, 10).' 23:59:59', $in);
            if (false !== $until) {
                $parts['UNTIL'] = $allDay ? $until->format('Ymd') : $until->setTimezone(new \DateTimeZone('UTC'))->format('Ymd\THis\Z');
            }
        }

        $parts = array_filter($parts, static fn (string $v) => '' !== $v);

        return Recurrence::rule(implode(';', array_map(static fn (string $k, string $v) => $k.'='.$v, array_keys($parts), $parts)));
    }

    /**
     * A Recurrence written as Graph's recurrence, from the event it repeats
     * (its start gives the range's start, its zone the recurrence's).
     *
     * @return array{pattern: array<string, mixed>, range: array<string, mixed>}
     *
     * @throws NotSupportedException
     */
    public static function fromRecurrence(Recurrence $recurrence, Event $event): array
    {
        if (1 !== \count($recurrence->rules) || [] !== $recurrence->exrules || [] !== $recurrence->rdates || [] !== $recurrence->exdates) {
            throw self::cannot('one rule alone (no EXRULE, RDATE or EXDATE: delete an occurrence to take it out)');
        }
        $parts = $recurrence->parts();
        if ([] !== $other = array_diff(array_keys($parts), self::PARTS)) {
            throw self::cannot(implode(', ', $other));
        }
        $zone = $event->timeZone ?? 'UTC';
        $start = $event->start->setTimezone(new \DateTimeZone($zone));
        $interval = $recurrence->interval();
        $days = array_map(self::day(...), $recurrence->days());
        $ordinals = array_filter($days, static fn (array $d) => null !== $d[0]);
        $setpos = isset($parts['BYSETPOS']) ? self::integer($parts['BYSETPOS'], 'BYSETPOS') : null;
        $weekday = array_search(strtolower($start->format('l')), self::DAYS, true) ?: 'MO';
        $pattern = match ($parts['FREQ'] ?? null) {
            'DAILY' => match (true) {
                [] === $days => ['type' => 'daily', 'interval' => $interval],
                1 === $interval && [] === $ordinals => ['type' => 'weekly', 'interval' => 1, 'daysOfWeek' => self::names($days), 'firstDayOfWeek' => 'monday'],
                default => throw self::cannot('a daily rule on given days every '.$interval.' days'),
            },
            'WEEKLY' => [] === $ordinals && null === $setpos ? ['type' => 'weekly', 'interval' => $interval, 'daysOfWeek' => self::names($days ?: [[null, $weekday]]), 'firstDayOfWeek' => self::DAYS[$parts['WKST'] ?? 'MO'] ?? 'monday'] : throw self::cannot('a weekly rule with ordinal days'),
            'MONTHLY' => isset($parts['BYMONTHDAY'])
                ? ([] === $days ? ['type' => 'absoluteMonthly', 'interval' => $interval, 'dayOfMonth' => self::monthDay($parts['BYMONTHDAY'])] : throw self::cannot('BYMONTHDAY with BYDAY'))
                : ([] === $days ? ['type' => 'absoluteMonthly', 'interval' => $interval, 'dayOfMonth' => (int) $start->format('j')] : ['type' => 'relativeMonthly', 'interval' => $interval] + self::relativePattern($days, $setpos)),
            'YEARLY' => ['month' => isset($parts['BYMONTH']) ? self::month($parts['BYMONTH']) : (int) $start->format('n')] + (isset($parts['BYMONTHDAY'])
                ? ([] === $days ? ['type' => 'absoluteYearly', 'interval' => $interval, 'dayOfMonth' => self::monthDay($parts['BYMONTHDAY'])] : throw self::cannot('BYMONTHDAY with BYDAY'))
                : ([] === $days ? ['type' => 'absoluteYearly', 'interval' => $interval, 'dayOfMonth' => (int) $start->format('j')] : ['type' => 'relativeYearly', 'interval' => $interval] + self::relativePattern($days, $setpos))),
            default => throw self::cannot('FREQ='.($parts['FREQ'] ?? '?')),
        };
        $range = ['type' => 'noEnd', 'startDate' => $start->format('Y-m-d'), 'recurrenceTimeZone' => $zone];
        if (null !== $count = $recurrence->count()) {
            $range = ['type' => 'numbered', 'numberOfOccurrences' => $count] + $range;
        } elseif (null !== $until = $recurrence->until()) {
            $range = ['type' => 'endDate', 'endDate' => (isset($parts['UNTIL']) && 8 === \strlen($parts['UNTIL']) ? $until : $until->setTimezone(new \DateTimeZone($zone)))->format('Y-m-d')] + $range;
        }

        return ['pattern' => $pattern, 'range' => $range];
    }

    /**
     * @param list<string> $days
     *
     * @return array<string, string>
     */
    private static function relative(array $days, int $index): array
    {
        $days = $days ?: ['MO'];

        return 1 === \count($days) ? ['BYDAY' => $index.$days[0]] : ['BYDAY' => implode(',', $days), 'BYSETPOS' => (string) $index];
    }

    /**
     * @param list<array{0: ?int, 1: string}> $days
     *
     * @return array{daysOfWeek: list<string>, index: string}
     */
    private static function relativePattern(array $days, ?int $setpos): array
    {
        $ordinals = array_values(array_unique(array_filter(array_map(static fn (array $d) => $d[0], $days), static fn (?int $o) => null !== $o)));
        $index = match (true) {
            1 === \count($days) && 1 === \count($ordinals) && null === $setpos => $ordinals[0],
            [] === $ordinals && null !== $setpos => $setpos,
            default => throw self::cannot('this BYDAY and BYSETPOS'),
        };
        $name = array_search($index, self::INDEX, true);
        if (false === $name) {
            throw self::cannot('the occurrence '.$index.' of a month (first to fourth, or last)');
        }

        return ['daysOfWeek' => array_map(static fn (array $d) => self::DAYS[$d[1]], $days), 'index' => $name];
    }

    /**
     * @param list<array{0: ?int, 1: string}> $days
     *
     * @return list<string>
     */
    private static function names(array $days): array
    {
        return array_map(static fn (array $d) => self::DAYS[$d[1]], $days);
    }

    /** @return array{0: ?int, 1: string} the ordinal and the day: "-1FR" is [-1, "FR"] */
    private static function day(string $day): array
    {
        if (!preg_match('~^([+-]?\d{1,2})?(MO|TU|WE|TH|FR|SA|SU)$~', $day, $m)) {
            throw self::cannot('BYDAY='.$day);
        }

        return ['' === $m[1] ? null : (int) $m[1], $m[2]];
    }

    private static function monthDay(string $value): int
    {
        $day = self::integer($value, 'BYMONTHDAY');

        return $day >= 1 && $day <= 31 ? $day : throw self::cannot('BYMONTHDAY='.$value);
    }

    private static function month(string $value): int
    {
        $month = self::integer($value, 'BYMONTH');

        return $month >= 1 && $month <= 12 ? $month : throw self::cannot('BYMONTH='.$value);
    }

    private static function integer(string $value, string $part): int
    {
        return preg_match('~^[+-]?\d+$~', $value) ? (int) $value : throw self::cannot($part.'='.$value);
    }

    private static function cannot(string $what): NotSupportedException
    {
        return NotSupportedException::call('Microsoft Graph', 'say '.$what.' in a recurrence pattern');
    }
}
