import { describe, expect, test, vi } from "vitest";
import { isRetryableGoogleCalendarError, runBestEffortGoogleCalendarMutation } from "./google-calendar-retry";

describe("Google Calendar mutation recovery", () => {
  test("retries the PATCH-specific 403 rate limit response", async () => {
    const operation = vi
      .fn()
      .mockRejectedValueOnce({
        response: { status: 403, data: { error: { errors: [{ reason: "rateLimitExceeded" }] } } },
      })
      .mockResolvedValueOnce(undefined);
    const sleeps: number[] = [];

    await expect(
      runBestEffortGoogleCalendarMutation(operation, {
        sleep: async (delay) => {
          sleeps.push(delay);
        },
      })
    ).resolves.toEqual({ ok: true });
    expect(operation).toHaveBeenCalledTimes(2);
    expect(sleeps).toEqual([250]);
  });

  test("returns a failure after bounded retries so callers can preserve an already-created event", async () => {
    const error = { code: 429 };
    const operation = vi.fn().mockRejectedValue(error);
    const result = await runBestEffortGoogleCalendarMutation(operation, {
      attempts: 3,
      sleep: async () => undefined,
    });

    expect(result).toEqual({ ok: false, error });
    expect(operation).toHaveBeenCalledTimes(3);
  });

  test("does not retry permanent permission errors", () => {
    expect(
      isRetryableGoogleCalendarError({
        response: {
          status: 403,
          data: {
            error: {
              errors: [{ reason: "forbidden" }],
            },
          },
        },
      })
    ).toBe(false);
  });
});
