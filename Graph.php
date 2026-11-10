<?php

namespace Omnical\Microsoft;

use Omnical\Auth\AccessTokenInterface;
use Omnical\Exception\AuthenticationException;
use Omnical\Exception\ProviderException;
use Omnical\Http\Answer;
use Omnical\Http\Data;
use Omnical\Http\JsonClient;
use Omnical\Http\RawClient;

/**
 * Graph's calls with the user's token: "Authorization: Bearer" from the
 * access token on every request, one refresh and one retry when Graph
 * answers 401, the preferences every call states (RFC 7240, in one Prefer
 * field: times in UTC, immutable ids, bodies as text). A next page's link
 * (@odata.nextLink) is followed only on Graph's own host: the token is
 * never sent elsewhere.
 */
final class Graph
{
    public const PREFER = 'outlook.timezone="UTC", IdType="ImmutableId", outlook.body-content-type="text"';

    public function __construct(
        private readonly JsonClient $json,
        private readonly RawClient $raw,
        private readonly AccessTokenInterface $token,
        private readonly string $baseUri = 'https://graph.microsoft.com/v1.0',
    ) {
    }

    /**
     * A JSON call: a path under the base URI ("/me/calendars"), or a next
     * page's absolute link.
     *
     * @param array<string, mixed> $options query or json
     *
     * @return ?Data null on a 404
     */
    public function json(string $method, string $path, array $options = []): ?Data
    {
        $url = $this->url($path);

        return $this->authorized(fn (string $token) => $this->json->request($method, $url, $options, ['Authorization' => 'Bearer '.$token, 'Prefer' => self::PREFER]));
    }

    /** A call whose answer is empty (a DELETE's 204); a 404 is given back. */
    public function raw(string $method, string $path): Answer
    {
        $url = $this->url($path);

        return $this->authorized(fn (string $token) => $this->raw->request($method, $url, [], ['Authorization' => 'Bearer '.$token, 'Prefer' => self::PREFER]));
    }

    /** Whether a next page's link points at Graph itself. */
    public function owns(string $url): bool
    {
        $base = parse_url($this->baseUri);
        $link = parse_url($url);

        return \is_array($base) && \is_array($link) && ($link['scheme'] ?? null) === ($base['scheme'] ?? null) && strtolower($link['host'] ?? '') === strtolower($base['host'] ?? '');
    }

    /**
     * @template T
     *
     * @param \Closure(string): T $call
     *
     * @return T
     */
    private function authorized(\Closure $call): mixed
    {
        try {
            return $call($this->token->token());
        } catch (AuthenticationException) {
            return $call($this->token->refresh());
        }
    }

    private function url(string $path): string
    {
        if (preg_match('~^https?://~', $path)) {
            return $this->owns($path) ? $path : throw new ProviderException(\sprintf('Microsoft Graph pointed to %s for the next page: the token is sent to Graph only.', (string) parse_url($path, \PHP_URL_HOST)));
        }

        return rtrim($this->baseUri, '/').$path;
    }
}
