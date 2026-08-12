type RetryResult = { ok: true } | { ok: false; error: unknown };

export async function runBestEffortGoogleCalendarMutation(
  operation: () => Promise<unknown>,
  options: {
    attempts?: number;
    sleep?: (delayMs: number) => Promise<void>;
  } = {}
): Promise<RetryResult> {
  const attempts = Math.max(1, options.attempts ?? 3);
  const sleep = options.sleep ?? ((delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)));
  let lastError: unknown;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      await operation();
      return { ok: true };
    } catch (error) {
      lastError = error;
      if (!isRetryableGoogleCalendarError(error) || attempt === attempts) break;
      await sleep(250 * 2 ** (attempt - 1));
    }
  }
  return { ok: false, error: lastError };
}

export function isRetryableGoogleCalendarError(error: unknown): boolean {
  const record = object(error);
  const response = object(record?.response);
  const data = object(response?.data);
  const nestedError = object(data?.error);
  const status = number(record?.code) ?? number(response?.status) ?? number(nestedError?.code);
  if (status === 429 || (status !== undefined && status >= 500 && status <= 599)) return true;
  if (status !== 403) return false;
  const reasons = [record, nestedError]
    .flatMap((value) => (Array.isArray(value?.errors) ? value.errors : []))
    .map((value) => object(value)?.reason)
    .filter((value): value is string => typeof value === "string");
  return reasons.some((reason) => reason === "rateLimitExceeded" || reason === "userRateLimitExceeded");
}

function object(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
}

function number(value: unknown): number | undefined {
  return typeof value === "number" ? value : undefined;
}
