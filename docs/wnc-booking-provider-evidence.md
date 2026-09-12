# WNC legacy booking evidence

`GET /api/wnc/booking-evidence?uid=<uid>` supplies read-only evidence for WNC's
independently reviewed booking-provider repair. It does not repair, create,
cancel, reschedule or send invitations. Keep it separate from the WNC frontend
change and do not deploy it without release authorization.

Authentication uses the existing `x-lia-internal-secret` header and a constant-time
comparison with `LIA_INTERNAL_SECRET`. Only sales event types with metadata
`clientSlug: "self"` and a nonempty `clientConfigId` are eligible. Responses disable
caching and never include credentials, attendee addresses or provider exceptions.

Successful responses contain `uid`, `referenceFingerprint`, `status`
(`scheduled` or `cancelled`), canonical UTC `start`/`end`, and `calendarState`
(`active` or `absent`). The fingerprint hashes the sorted calendar-reference
identities, including provider, event ID, credential ID and external calendar ID.
It deliberately ignores the local `deleted` flag.

Each reference is checked through its original Google credential. Delegated,
invalid, missing and unsupported credentials are refused. The referenced calendar
must be readable with its exact ID before an event 404/410 can establish absence.
An active event must retain the Cal booking's exact start/end; a provider-cancelled
tombstone is absent. Empty references, mixed states, unknown booking states,
credential failure and provider errors return 503. Cal cancellation by itself is
not evidence that the Google invitation was removed.

The implementation reuses the existing Google calendar adapter; its normal OAuth
refresh behavior still applies. No API credential or new configuration is added.
This scope supports legacy Google calendar references with explicit calendar IDs.
Other providers require their own reviewed exact-reference reader.

Verification on September 12, 2026:

```powershell
yarn test --run apps/web/app/api/wnc/booking-evidence/route.test.ts
yarn biome check apps/web/app/api/wnc/booking-evidence/route.ts apps/web/app/api/wnc/booking-evidence/evidence.ts apps/web/app/api/wnc/booking-evidence/route.test.ts
yarn workspace @calcom/web type-check
```

All six focused tests, Biome and the direct web workspace type-check passed.
The required `yarn type-check:ci --force --filter=@calcom/web` was also attempted;
its Windows `tsc-absolute` child could not spawn `tsc` (`ENOENT`). No error was
reported against the evidence code. Tests mock the provider adapter and do not
prove live calendar access or delivery. New functions have cyclomatic complexity
at most 9 under a consistent TypeScript AST decision count.
