<?php

namespace Omnical\Microsoft;

use Omnical\Auth\AccessTokenInterface;
use Omnical\Auth\EnvSecrets;
use Omnical\Auth\OAuth2Client;
use Omnical\Auth\RefreshingAccessToken;
use Omnical\Auth\SecretsInterface;
use Omnical\Auth\StaticAccessToken;
use Omnical\Auth\TokenStoreInterface;
use Omnical\Config;
use Omnical\Exception\InvalidConfigException;
use Omnical\Http\JsonClient;
use Omnical\Http\RawClient;
use Omnical\ProviderFactory;
use Omnical\ProviderInterface;
use Symfony\Component\HttpClient\HttpClient;
use Symfony\Contracts\HttpClient\HttpClientInterface;

/**
 * "microsoft": the signed-in user's Outlook calendars through Microsoft
 * Graph. Options:
 *
 * - token: the key the user's OAuth token is kept under in the factory's token store
 *   ("user:42:microsoft"); refreshed when it expires or Graph refuses it (Microsoft sends a new
 *   refresh token each time: the store keeps it). Needs client_id, and client_secret for a
 *   confidential (web) application.
 * - access_token: instead, the reference of a token given as it is (never refreshed).
 * - client_id: the application (client) id of its Entra registration.
 * - client_secret: the REFERENCE of its client secret ("MICROSOFT_CLIENT_SECRET": an
 *   environment variable by default, or the factory's secrets), never the secret.
 * - tenant: "common" (work, school and personal accounts; the default), "organizations",
 *   "consumers", or a tenant id or domain.
 * - base_uri: https://graph.microsoft.com/v1.0 (a national cloud's Graph otherwise).
 * - timeout: seconds; availability_interval: getSchedule's minutes per slot (5 to 1440, 30).
 * - policy, usage: the site's tightening of the policy, and what the site says of itself.
 *
 * oauth() gives the OAuth2Client of the consent step (authorization code with PKCE).
 */
final class MicrosoftFactory extends ProviderFactory
{
    public const SCOPE_OFFLINE = 'offline_access';
    public const SCOPE_READ_BASIC = 'https://graph.microsoft.com/Calendars.ReadBasic';
    public const SCOPE_READ = 'https://graph.microsoft.com/Calendars.Read';
    public const SCOPE_READ_WRITE = 'https://graph.microsoft.com/Calendars.ReadWrite';

    /** @param ?\Closure(): int $now the clock tokens expire by (time() by default) */
    public function __construct(
        private readonly ?HttpClientInterface $http = null,
        private readonly ?TokenStoreInterface $tokens = null,
        private readonly ?SecretsInterface $secrets = null,
        private readonly ?\Closure $now = null,
    ) {
    }

    /**
     * The scopes to ask the user for: reading (Calendars.Read, which
     * getSchedule needs too) or writing (Calendars.ReadWrite), with a refresh
     * token (offline_access).
     *
     * @return list<string>
     */
    public static function scopes(bool $write = true): array
    {
        return [self::SCOPE_OFFLINE, $write ? self::SCOPE_READ_WRITE : self::SCOPE_READ];
    }

    /**
     * The OAuth client of the consent step and of the refreshes, from the
     * same options as create(): client_id, client_secret, tenant.
     *
     * @param array<string, mixed> $options
     */
    public function oauth(array $options): OAuth2Client
    {
        return $this->client($this->createConfig($options));
    }

    protected function populate(Config $c): void
    {
        $c->defaults([
            'omnical.factory_name' => 'microsoft',
            'omnical.factory_title' => 'Microsoft Graph',
            'omnical.required_options' => [],
            'base_uri' => 'https://graph.microsoft.com/v1.0',
            'tenant' => 'common',
            'client_id' => null,
            'client_secret' => null,
            'token' => null,
            'access_token' => null,
            'timeout' => null,
            'availability_interval' => 30,
        ]);
    }

    protected function build(Config $c): ProviderInterface
    {
        $base = rtrim((string) $c->string('base_uri'), '/');
        if (!preg_match('~^https://[^/\s]+~', $base)) {
            throw new InvalidConfigException(\sprintf('The "microsoft" provider needs an https base_uri such as https://graph.microsoft.com/v1.0, not "%s".', $base));
        }
        $interval = $c->int('availability_interval', 30);
        if ($interval < 5 || $interval > 1440) {
            throw new InvalidConfigException('getSchedule\'s availability_interval is 5 to 1440 minutes.');
        }
        $http = $this->httpClient($c);
        $graph = new Graph(
            new JsonClient($http, 'Microsoft Graph', ['Accept' => 'application/json'], $c->backoff()),
            new RawClient($http, 'Microsoft Graph', [], $c->backoff()),
            $this->token($c),
            $base,
        );

        return new MicrosoftProvider($graph, $interval);
    }

    private function token(Config $c): AccessTokenInterface
    {
        $secrets = $this->secrets ?? new EnvSecrets();
        if (null !== $reference = $c->string('access_token')) {
            return new StaticAccessToken($secrets, $reference, 'Microsoft Graph');
        }
        $key = $c->string('token') ?? throw new InvalidConfigException('The "microsoft" provider needs the key of the user\'s token in the token store ("token"), or the reference of an access token ("access_token").');
        $store = $this->tokens ?? throw new InvalidConfigException('The "microsoft" provider reads the user\'s token from a token store: give one to its factory.');

        return new RefreshingAccessToken($this->client($c), $store, $key, 60, $this->now);
    }

    private function client(Config $c): OAuth2Client
    {
        $tenant = (string) $c->string('tenant');
        if (!preg_match('~^[A-Za-z0-9][A-Za-z0-9.-]*$~', $tenant)) {
            throw new InvalidConfigException(\sprintf('"%s" is no tenant: common, organizations, consumers, a tenant id or domain.', $tenant));
        }
        $authority = 'https://login.microsoftonline.com/'.$tenant.'/oauth2/v2.0';

        return new OAuth2Client(
            $this->httpClient($c),
            'Microsoft identity platform',
            $authority.'/authorize',
            $authority.'/token',
            $c->string('client_id') ?? throw new InvalidConfigException('The "microsoft" provider needs its application\'s client_id to refresh a token.'),
            $this->secrets ?? new EnvSecrets(),
            $c->string('client_secret'),
            null,
            $this->now,
        );
    }

    private function httpClient(Config $c): HttpClientInterface
    {
        $http = $this->http ?? HttpClient::create(['timeout' => 20]);

        return null === ($timeout = $c->string('timeout')) ? $http : $http->withOptions(['timeout' => (float) $timeout]);
    }
}
