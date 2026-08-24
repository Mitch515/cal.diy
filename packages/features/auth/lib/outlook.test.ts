import process from "node:process";
import { afterEach, describe, expect, it, vi } from "vitest";

describe("outlook", () => {
  const originalEnv = { ...process.env };

  afterEach(() => {
    // Restore original env completely
    for (const key of [
      "MS_GRAPH_CLIENT_ID",
      "MS_GRAPH_CLIENT_SECRET",
      "MS_GRAPH_TENANT_ID",
      "OUTLOOK_LOGIN_ENABLED",
    ]) {
      if (originalEnv[key] === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = originalEnv[key];
      }
    }
    vi.resetModules();
  });

  it("reads OUTLOOK_CLIENT_ID from MS_GRAPH_CLIENT_ID", async () => {
    process.env.MS_GRAPH_CLIENT_ID = "test-client-id";
    const { OUTLOOK_CLIENT_ID } = await import("./outlook");
    expect(OUTLOOK_CLIENT_ID).toBe("test-client-id");
  });

  it("reads OUTLOOK_CLIENT_SECRET from MS_GRAPH_CLIENT_SECRET", async () => {
    process.env.MS_GRAPH_CLIENT_SECRET = "test-secret";
    const { OUTLOOK_CLIENT_SECRET } = await import("./outlook");
    expect(OUTLOOK_CLIENT_SECRET).toBe("test-secret");
  });

  it("uses a configured Microsoft tenant for single-tenant applications", async () => {
    process.env.MS_GRAPH_TENANT_ID = "f1a52f59-393f-43d3-813f-fa1197512059";
    const { OUTLOOK_TENANT_ID } = await import("./outlook");
    expect(OUTLOOK_TENANT_ID).toBe("f1a52f59-393f-43d3-813f-fa1197512059");
  });

  it("falls back to the Wealth Navigator tenant when the tenant value is unsafe", async () => {
    process.env.MS_GRAPH_TENANT_ID = "tenant/../../common";
    const { OUTLOOK_TENANT_ID } = await import("./outlook");
    expect(OUTLOOK_TENANT_ID).toBe("f1a52f59-393f-43d3-813f-fa1197512059");
  });

  it("uses the Wealth Navigator tenant when no tenant is configured", async () => {
    delete process.env.MS_GRAPH_TENANT_ID;
    const { OUTLOOK_TENANT_ID } = await import("./outlook");
    expect(OUTLOOK_TENANT_ID).toBe("f1a52f59-393f-43d3-813f-fa1197512059");
  });

  it("OUTLOOK_LOGIN_ENABLED is true only for exact string 'true'", async () => {
    process.env.OUTLOOK_LOGIN_ENABLED = "true";
    const mod = await import("./outlook");
    expect(mod.OUTLOOK_LOGIN_ENABLED).toBe(true);
  });

  it("OUTLOOK_LOGIN_ENABLED is false for other values", async () => {
    process.env.OUTLOOK_LOGIN_ENABLED = "1";
    const mod = await import("./outlook");
    expect(mod.OUTLOOK_LOGIN_ENABLED).toBe(false);
  });

  it("IS_OUTLOOK_LOGIN_ENABLED is true when all vars are set", async () => {
    process.env.MS_GRAPH_CLIENT_ID = "client-id";
    process.env.MS_GRAPH_CLIENT_SECRET = "client-secret";
    process.env.OUTLOOK_LOGIN_ENABLED = "true";
    const { IS_OUTLOOK_LOGIN_ENABLED } = await import("./outlook");
    expect(IS_OUTLOOK_LOGIN_ENABLED).toBe(true);
  });

  it("IS_OUTLOOK_LOGIN_ENABLED is false when client ID is missing", async () => {
    delete process.env.MS_GRAPH_CLIENT_ID;
    process.env.MS_GRAPH_CLIENT_SECRET = "client-secret";
    process.env.OUTLOOK_LOGIN_ENABLED = "true";
    const { IS_OUTLOOK_LOGIN_ENABLED } = await import("./outlook");
    expect(IS_OUTLOOK_LOGIN_ENABLED).toBe(false);
  });

  it("IS_OUTLOOK_LOGIN_ENABLED is false when client secret is missing", async () => {
    process.env.MS_GRAPH_CLIENT_ID = "client-id";
    delete process.env.MS_GRAPH_CLIENT_SECRET;
    process.env.OUTLOOK_LOGIN_ENABLED = "true";
    const { IS_OUTLOOK_LOGIN_ENABLED } = await import("./outlook");
    expect(IS_OUTLOOK_LOGIN_ENABLED).toBe(false);
  });
});
