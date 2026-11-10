<?php

namespace Omnical\Microsoft\Tests;

use Omnical\Auth\ArraySecrets;
use Omnical\Exception\UsageLimitReachedException;
use Omnical\Microsoft\MicrosoftFactory;
use Omnical\Registry;
use Omnical\Testing\RecordedHttpClient;
use Omnisense\Contracts\Policy\InvalidPolicyException;
use Omnisense\Contracts\Policy\Permission;
use Omnisense\Contracts\Policy\Scope;
use PHPUnit\Framework\TestCase;

/** Microsoft Graph's usage policy, declared, counted per user (mailbox), tightened by a site, never loosened. */
final class PolicyTest extends TestCase
{
    public function testThePolicyIsTheOneMicrosoftPublishes(): void
    {
        $policy = self::registry([])->policy('outlook');

        self::assertSame('microsoft', $policy->provider);
        self::assertSame(['1000/minute/user/*'], array_map(static fn ($l) => $l->id(), $policy->rateLimits ?? []));
        self::assertSame(Scope::User, ($policy->rateLimits ?? [])[0]->scope);
        self::assertStringContainsString('10,000 API requests in any 10 minutes', (string) ($policy->rateLimits ?? [])[0]->note);
        self::assertSame([Permission::Allowed, Permission::Forbidden, false], [$policy->rights->commercial, $policy->rights->redistribution, $policy->rights->attribution->required]);
        self::assertTrue($policy->requirements->key);
        self::assertSame(['https://learn.microsoft.com/en-us/legal/microsoft-apis/terms-of-use', '2026-11-10'], [$policy->provenance->terms, $policy->provenance->verifiedAt]);
    }

    public function testEachUserCountsApart(): void
    {
        $http = RecordedHttpClient::routes(['~/me/calendars~' => [200, '{"value":[]}']]);
        $registry = self::registry(['rateLimits' => [['max' => 2, 'window' => 'minute', 'scope' => 'user']]], $http);

        $registry->calendarEvents('outlook', 'user-1')->calendars();
        $registry->calendarEvents('outlook', 'user-1')->calendars();
        $registry->calendarEvents('outlook', 'user-2')->calendars();
        try {
            $registry->calendarEvents('outlook', 'user-1')->calendars();
            self::fail('A third call of the minute went out for user-1.');
        } catch (UsageLimitReachedException $e) {
            self::assertSame('2/minute/user/*', $e->limit?->id());
        }
        self::assertCount(3, $http->requests, 'the refused call reached nothing');
    }

    public function testASiteCannotRaiseTheLimit(): void
    {
        $this->expectException(InvalidPolicyException::class);
        self::registry(['rateLimits' => [['max' => 5000, 'window' => 'minute', 'scope' => 'user']]])->calendarEvents('outlook');
    }

    /** @param array<string, mixed> $policy */
    private static function registry(array $policy, ?RecordedHttpClient $http = null): Registry
    {
        return new Registry([new MicrosoftFactory($http ?? RecordedHttpClient::routes([]), null, new ArraySecrets(['MS_TOKEN' => 'not-a-real-token']))], ['outlook' => ['factory' => 'microsoft', 'options' => ['access_token' => 'MS_TOKEN'], 'policy' => $policy]]);
    }
}
