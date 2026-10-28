# I Hate Spam

A free, open-source inbox cleanup tool by Ethia for Gmail and Outlook.

Website: https://ihatespam.site/ · Contact: ethia.studio@gmail.com

Repository: https://github.com/ethia-ai/ihatespam

This source release matches the application source deployed as Sites version 70
(source commit `c1b407f890abafb5368c9128b55f340b1d5454c0`). Hosting configuration
and internal operational notes are excluded. Public documentation and licensing
are maintained separately from the deployment checkout.

## Important: Gmail verification

This application is **not verified by Google** and has not completed the CASA
assessment required by Google's review team. Gmail access may be limited or
blocked. Open-source publication is not a security certification or a waiver
of Google's requirements. Do not treat this repository as evidence of approval.
Only grant access if you understand the permissions and trust the project.

## Features

- Analyse recent inbox message metadata and group related sender addresses.
- Select individual addresses and review actions before applying them.
- Archive, move to Trash, and manage supported spam actions.
- Gmail future-sender routing uses exact-address filters to Trash, not the
  native blocked-senders list. Outlook future routing uses Inbox rules;
  native Outlook blocking is a manual step in Microsoft's settings.
- Assisted unsubscribe opens the sender's page. Optional server-side automatic
  unsubscribe supports a conservative set of verified Gmail headers/destinations.
- No permanent-delete or empty-trash action is implemented. Provider retention
  policies still apply. A successful unsubscribe request does not prove removal.

## Architecture and privacy

`dist/` contains the hand-written browser application, **not disposable build
output**. Browser code calls Gmail and Microsoft Graph directly for analysis
and cleanup. Message bodies and attachments are not requested by analysis.

The automatic-unsubscribe backend in `server/` temporarily receives the Gmail
access token and retrieves headers/account identity. It uses encrypted plans,
signature and destination checks, and a durable D1 ledger for quotas/replay
protection. The ledger contains keyed identifiers, counters and outcomes, not
raw mailbox tokens or email content. See the website's privacy policy for
retention and provider-response limitations.

Access tokens live in page memory. Language and treated-sender history
(including account/sender addresses) use local browser storage. Clearing site
data removes that local history. Hosting and identity providers also process
connection information. We do **not** claim that the application processes no
personal data or is guaranteed secure.

## Run locally

Use Node.js 22.13 or later (tests use `node:sqlite`).

```sh
npm test
npm run build
python3 -m http.server 8080 --directory dist
```

Open http://localhost:8080 and use the demo inbox. Static preview does not run
the unsubscribe API. Do not test real mailbox mutations without explicit consent.
Tests use synthetic fixtures; they do not access live mailboxes.

`npm ci` installs the pinned development tools if you need to generate database
migrations (`npm run db:generate`). Build and tests use Node built-ins.

## Deploying your own instance

1. Register your own Google and Microsoft OAuth applications. Replace the
   **public** client IDs in `dist/app.js` and configure your own approved origins
   and redirect URI (the Microsoft callback is `/auth-redirect.html`). Never add
   a client secret to browser code. Follow each provider's verification rules.
2. Replace the canonical/site URLs, operator contact, sharing links and
   `UNSUBSCRIBE_ORIGIN` in `server/unsubscribe-api.mjs`. Review legal copy.
3. Build the Worker with `npm run build`. Its entrypoint is
   `dist/server/index.js`, with static assets embedded by an explicit allowlist.
4. If enabling automatic unsubscribe, bind your own D1 database as `DB` and
   apply `drizzle/0000_useful_lady_deathstrike.sql`. Store a strong, independent
   `UNSUBSCRIBE_SECRET` in your hosting secret manager, and set
   `UNSUBSCRIBE_ENABLED=true` only after reviewing the security requirements.
   Never reuse another deployment's secret or publish it in source control.
5. Without those bindings the automatic API fails closed; it is not emulated
   by the static preview. Review CSP, HTTPS, logs, quotas and mailbox consent
   before offering a live deployment.

Production hosting project identifiers, deployment credentials, operational
databases, internal review correspondence and prior private Git history are
intentionally excluded. OAuth client IDs in the shipped browser code are
public identifiers, **not** client secrets or mailbox credentials.

## Security and licensing

See [SECURITY.md](SECURITY.md) for private vulnerability reporting. This source
release is not an independent security audit. MIT license for project code;
vendored Microsoft libraries retain their license in
`dist/vendor/MSAL-LICENSE.txt`. Google UI screenshots are illustrative third-party
interface material, not an endorsement or grant of rights to Google's branding.
