// @vitest-environment node
import { afterEach, expect, test, vi } from "vitest";

const find = vi.hoisted(() => vi.fn());
vi.mock("@calcom/prisma", () => ({ default: { booking: { findUnique: find } } }));

import { isWncNativeBooking } from "./wnc-confirmation";

afterEach(() => vi.resetAllMocks());

test("native mail routing survives Cal replacing or clearing the slot key", async () => {
  for (const idempotencyKey of ["cal-slot-key", null]) {
    find.mockResolvedValue({
      idempotencyKey,
      metadata: { wncRequestId: "a".repeat(64) },
      fromReschedule: null,
    });
    expect(await isWncNativeBooking("original")).toBe(true);
  }
});

test("a reschedule inherits native mail routing through its original booking", async () => {
  find
    .mockResolvedValueOnce({ idempotencyKey: null, metadata: {}, fromReschedule: "original" })
    .mockResolvedValueOnce({
      idempotencyKey: null,
      metadata: { wncRequestId: "a".repeat(64) },
      fromReschedule: null,
    });
  expect(await isWncNativeBooking("rescheduled")).toBe(true);
  expect(find).toHaveBeenLastCalledWith(expect.objectContaining({ where: { uid: "original" } }));
});

test("unrelated or malformed metadata does not select native mail routing", async () => {
  for (const metadata of [null, {}, { wncRequestId: "invalid" }]) {
    find.mockResolvedValue({ idempotencyKey: "cal-slot-key", metadata, fromReschedule: null });
    expect(await isWncNativeBooking("ordinary")).toBe(false);
  }
});
