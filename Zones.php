<?php

namespace Omnical\Microsoft;

/**
 * Time zones as Graph names them, read as IANA names: Windows names
 * ("Romance Standard Time") through ext-intl's IntlTimeZone when it is
 * there, through a table of the common ones (CLDR's windowsZones, territory
 * 001) otherwise; an IANA name as it is; Outlook's custom zones
 * ("tzone://Microsoft/Custom") and anything else as unknown (null).
 */
final class Zones
{
    /** @var array<string, string> */
    public const WINDOWS = [
        'UTC' => 'UTC',
        'GMT Standard Time' => 'Europe/London',
        'Greenwich Standard Time' => 'Atlantic/Reykjavik',
        'W. Europe Standard Time' => 'Europe/Berlin',
        'Romance Standard Time' => 'Europe/Paris',
        'Central Europe Standard Time' => 'Europe/Budapest',
        'Central European Standard Time' => 'Europe/Warsaw',
        'E. Europe Standard Time' => 'Europe/Chisinau',
        'FLE Standard Time' => 'Europe/Kiev',
        'GTB Standard Time' => 'Europe/Bucharest',
        'Turkey Standard Time' => 'Europe/Istanbul',
        'Russian Standard Time' => 'Europe/Moscow',
        'Israel Standard Time' => 'Asia/Jerusalem',
        'Egypt Standard Time' => 'Africa/Cairo',
        'South Africa Standard Time' => 'Africa/Johannesburg',
        'Morocco Standard Time' => 'Africa/Casablanca',
        'W. Central Africa Standard Time' => 'Africa/Lagos',
        'Arab Standard Time' => 'Asia/Riyadh',
        'Arabian Standard Time' => 'Asia/Dubai',
        'India Standard Time' => 'Asia/Kolkata',
        'SE Asia Standard Time' => 'Asia/Bangkok',
        'Singapore Standard Time' => 'Asia/Singapore',
        'China Standard Time' => 'Asia/Shanghai',
        'Taipei Standard Time' => 'Asia/Taipei',
        'Korea Standard Time' => 'Asia/Seoul',
        'Tokyo Standard Time' => 'Asia/Tokyo',
        'AUS Eastern Standard Time' => 'Australia/Sydney',
        'New Zealand Standard Time' => 'Pacific/Auckland',
        'Hawaiian Standard Time' => 'Pacific/Honolulu',
        'Alaskan Standard Time' => 'America/Anchorage',
        'Pacific Standard Time' => 'America/Los_Angeles',
        'US Mountain Standard Time' => 'America/Phoenix',
        'Mountain Standard Time' => 'America/Denver',
        'Central Standard Time' => 'America/Chicago',
        'Central Standard Time (Mexico)' => 'America/Mexico_City',
        'Canada Central Standard Time' => 'America/Regina',
        'Eastern Standard Time' => 'America/New_York',
        'SA Pacific Standard Time' => 'America/Bogota',
        'Atlantic Standard Time' => 'America/Halifax',
        'Newfoundland Standard Time' => 'America/St_Johns',
        'E. South America Standard Time' => 'America/Sao_Paulo',
        'Argentina Standard Time' => 'America/Buenos_Aires',
        'Pacific SA Standard Time' => 'America/Santiago',
    ];

    /** The IANA name of a zone Graph sends; null when it cannot be told. */
    public static function toIana(?string $name): ?string
    {
        $name = trim((string) $name);
        if ('' === $name || str_starts_with($name, 'tzone://')) {
            return null;
        }
        if (\in_array($name, \DateTimeZone::listIdentifiers(\DateTimeZone::ALL_WITH_BC), true) || 'UTC' === $name) {
            return $name;
        }
        if (class_exists(\IntlTimeZone::class)) {
            $id = (string) \IntlTimeZone::getIDForWindowsID($name);
            if ('' !== $id && 'Etc/Unknown' !== $id) {
                return $id;
            }
        }

        return self::WINDOWS[$name] ?? null;
    }
}
