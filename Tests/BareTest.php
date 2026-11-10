<?php

namespace Omnical\Microsoft\Tests;

use Omnical\Microsoft\MicrosoftProvider;
use Omnical\Testing\BareTestCase;

/** omnical/microsoft's classes load with omnisense/omnical, omnisense/contracts and Symfony's HTTP client contracts alone (no ext-intl needed). */
final class BareTest extends BareTestCase
{
    protected static function entryClass(): string
    {
        return MicrosoftProvider::class;
    }

    protected static function requiredPackages(): array
    {
        return ['symfony/http-client-contracts'];
    }
}
