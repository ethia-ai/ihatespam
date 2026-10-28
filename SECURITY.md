# Security policy

Report suspected vulnerabilities privately to **ethia.studio@gmail.com**.
Do not create a public issue containing mailbox information, access tokens,
personal unsubscribe links, encryption keys, or exploitable production details.
Include reproduction steps using a synthetic account/fixture where possible.
Never test against another person's mailbox without their explicit permission.

## Security boundaries

- Browser tokens must not be persisted, logged, or committed. Public OAuth
  client IDs are not secrets; restrict origins in the provider configuration.
- Automatic unsubscribe temporarily processes tokens/headers on the backend.
  Server encryption secrets belong only in the hosting secret manager.
- Backend request origin checks are not authentication. Google validates the
  credential; mailbox binding, encrypted expiring plans, durable replay checks,
  quotas and destination restrictions provide additional controls.
- Never broaden the destination allowlist without reviewing SSRF, DNS and
  redirect behaviour. Never forward OAuth tokens or cookies to senders.
- A successful HTTP request is not proof that a subscription was cancelled.
- Local treated-sender history includes addresses and must be disclosed.
- This project has not completed Google's required CASA assessment and does
  not claim Google verification, independent certification or absolute safety.

Before deployment, review dependencies (including vendored libraries), OAuth
configuration, hosting access/log retention and privacy documentation. The
repository's automated tests and secret scanning are not a substitute for an
independent assessment. Do not publish runtime data or private Git history.
