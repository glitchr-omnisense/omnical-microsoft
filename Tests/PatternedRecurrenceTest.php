<?php

namespace Omnical\Microsoft\Tests;

use Omnical\Exception\NotSupportedException;
use Omnical\Http\Data;
use Omnical\Microsoft\PatternedRecurrence;
use Omnical\Model\Event;
use Omnical\Model\Recurrence;
use PHPUnit\Framework\TestCase;

/** Graph's patternedRecurrence and RRULE, both ways, on their own. */
final class PatternedRecurrenceTest extends TestCase
{
    /** @return iterable<string, array{0: array<string, mixed>, 1: string}> */
    public static function patterns(): iterable
    {
        $range = ['type' => 'noEnd', 'startDate' => '2026-11-16', 'recurrenceTimeZone' => 'Europe/Paris'];
        yield 'daily every 2 days, 10 times' => [['pattern' => ['type' => 'daily', 'interval' => 2], 'range' => ['type' => 'numbered', 'numberOfOccurrences' => 10] + $range], 'FREQ=DAILY;INTERVAL=2;COUNT=10'];
        yield 'weekly, Monday and Wednesday, until a day' => [['pattern' => ['type' => 'weekly', 'interval' => 1, 'daysOfWeek' => ['monday', 'wednesday'], 'firstDayOfWeek' => 'monday'], 'range' => ['type' => 'endDate', 'endDate' => '2027-06-30'] + $range], 'FREQ=WEEKLY;BYDAY=MO,WE;UNTIL=20270630T215959Z'];
        yield 'every other week from Sunday' => [['pattern' => ['type' => 'weekly', 'interval' => 2, 'daysOfWeek' => ['friday'], 'firstDayOfWeek' => 'sunday'], 'range' => $range], 'FREQ=WEEKLY;INTERVAL=2;BYDAY=FR;WKST=SU'];
        yield 'the 15th of each month' => [['pattern' => ['type' => 'absoluteMonthly', 'interval' => 1, 'dayOfMonth' => 15], 'range' => $range], 'FREQ=MONTHLY;BYMONTHDAY=15'];
        yield 'the second Tuesday' => [['pattern' => ['type' => 'relativeMonthly', 'interval' => 1, 'daysOfWeek' => ['tuesday'], 'index' => 'second'], 'range' => $range], 'FREQ=MONTHLY;BYDAY=2TU'];
        yield 'the last weekday' => [['pattern' => ['type' => 'relativeMonthly', 'interval' => 1, 'daysOfWeek' => ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'], 'index' => 'last'], 'range' => $range], 'FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-1'];
        yield 'every 8 March' => [['pattern' => ['type' => 'absoluteYearly', 'interval' => 1, 'month' => 3, 'dayOfMonth' => 8], 'range' => $range], 'FREQ=YEARLY;BYMONTH=3;BYMONTHDAY=8'];
        yield 'the fourth Thursday of November' => [['pattern' => ['type' => 'relativeYearly', 'interval' => 1, 'month' => 11, 'daysOfWeek' => ['thursday'], 'index' => 'fourth'], 'range' => $range], 'FREQ=YEARLY;BYMONTH=11;BYDAY=4TH'];
    }

    /** @param array<string, mixed> $graph */
    #[\PHPUnit\Framework\Attributes\DataProvider('patterns')]
    public function testAPatternIsARuleAndTheRuleThePatternAgain(array $graph, string $rule): void
    {
        $recurrence = PatternedRecurrence::toRecurrence(Data::of($graph), false, 'Europe/Paris');
        self::assertSame([$rule], $recurrence?->rules);

        $event = new Event('x', new \DateTimeImmutable('2026-11-16T09:00', new \DateTimeZone('Europe/Paris')), new \DateTimeImmutable('2026-11-16T10:00', new \DateTimeZone('Europe/Paris')));
        $back = PatternedRecurrence::fromRecurrence($recurrence, $event);
        $pattern = \is_array($graph['pattern']) ? $graph['pattern'] : [];
        if ('weekly' === $pattern['type'] && !isset($pattern['firstDayOfWeek'])) {
            $pattern['firstDayOfWeek'] = 'monday';
        }
        ksort($pattern);
        $written = $back['pattern'];
        ksort($written);
        self::assertSame($pattern, $written);
        self::assertSame(\is_array($graph['range']) ? $graph['range']['type'] : null, $back['range']['type']);
        if ('endDate' === $back['range']['type']) {
            self::assertSame('2027-06-30', $back['range']['endDate']);
        }
        self::assertSame(['2026-11-16', 'Europe/Paris'], [$back['range']['startDate'], $back['range']['recurrenceTimeZone']]);
    }

    public function testAnAllDaySeriesEndsOnADateAndWeekdaysDailyAreWeekly(): void
    {
        $graph = ['pattern' => ['type' => 'daily', 'interval' => 1], 'range' => ['type' => 'endDate', 'startDate' => '2026-11-16', 'endDate' => '2026-12-18', 'recurrenceTimeZone' => 'Romance Standard Time']];
        self::assertSame(['FREQ=DAILY;UNTIL=20261218'], PatternedRecurrence::toRecurrence(Data::of($graph), true, 'Europe/Paris')?->rules);

        $event = new Event('x', new \DateTimeImmutable('2026-11-16'), new \DateTimeImmutable('2026-11-17'), allDay: true, timeZone: 'Europe/Paris');
        $back = PatternedRecurrence::fromRecurrence(Recurrence::rule('FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;UNTIL=20261218'), $event);
        self::assertSame(['type' => 'weekly', 'interval' => 1, 'daysOfWeek' => ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'], 'firstDayOfWeek' => 'monday'], $back['pattern']);
        self::assertSame(['type' => 'endDate', 'endDate' => '2026-12-18', 'startDate' => '2026-11-16', 'recurrenceTimeZone' => 'Europe/Paris'], $back['range']);
        self::assertNull(PatternedRecurrence::toRecurrence(Data::of(['pattern' => ['type' => 'hourly']]), false, 'UTC'));
    }

    /** @return iterable<string, array{0: Recurrence, 1: string}> */
    public static function unsaid(): iterable
    {
        yield 'hours' => [Recurrence::rule('FREQ=DAILY;BYHOUR=9,14'), 'BYHOUR'];
        yield 'minutely' => [Recurrence::rule('FREQ=MINUTELY;INTERVAL=30'), 'FREQ=MINUTELY'];
        yield 'two rules' => [new Recurrence(['FREQ=DAILY', 'FREQ=WEEKLY']), 'one rule alone'];
        yield 'an exception date' => [Recurrence::rule('FREQ=WEEKLY', [new \DateTimeImmutable('2026-11-23')]), 'EXDATE'];
        yield 'the fifth Monday' => [Recurrence::rule('FREQ=MONTHLY;BYDAY=5MO'), 'the occurrence 5'];
        yield 'two month days' => [Recurrence::rule('FREQ=MONTHLY;BYMONTHDAY=1,15'), 'BYMONTHDAY=1,15'];
        yield 'the last day of the month' => [Recurrence::rule('FREQ=MONTHLY;BYMONTHDAY=-1'), 'BYMONTHDAY=-1'];
        yield 'a weekly rule with ordinals' => [Recurrence::rule('FREQ=WEEKLY;BYDAY=1MO'), 'ordinal'];
    }

    #[\PHPUnit\Framework\Attributes\DataProvider('unsaid')]
    public function testARuleGraphCannotSayIsRefused(Recurrence $recurrence, string $message): void
    {
        $this->expectException(NotSupportedException::class);
        $this->expectExceptionMessage($message);
        PatternedRecurrence::fromRecurrence($recurrence, new Event('x', new \DateTimeImmutable('2026-11-16T09:00Z'), new \DateTimeImmutable('2026-11-16T10:00Z')));
    }
}
