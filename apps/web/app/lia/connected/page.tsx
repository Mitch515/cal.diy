// Minimal standalone "calendar connected" page for the LIA host-onboarding flow.
// The Microsoft OAuth callback redirects here (returnTo) so the host lands on a
// clean confirmation instead of the Cal.diy app shell. No app chrome, no data,
// no auth dependency — just a self-contained confirmation.

export const dynamic = "force-static";

export const metadata = {
  title: "Calendar connected",
  robots: { index: false, follow: false },
};

export default function LiaCalendarConnectedPage() {
  return (
    <div
      style={{
        minHeight: "100vh",
        display: "grid",
        placeItems: "center",
        padding: "32px 20px",
        background: "#f4f7f6",
        fontFamily:
          "Inter, ui-sans-serif, -apple-system, BlinkMacSystemFont, 'Segoe UI', system-ui, sans-serif",
      }}>
      <div
        style={{
          width: "100%",
          maxWidth: 420,
          background: "#ffffff",
          border: "1px solid #e3e8e6",
          borderRadius: 14,
          padding: "36px 30px",
          textAlign: "center",
          boxShadow: "0 4px 12px rgba(0, 70, 67, 0.10)",
        }}>
        <div
          style={{
            width: 52,
            height: 52,
            margin: "0 auto 18px",
            borderRadius: "50%",
            background: "#def0ea",
            color: "#00665f",
            display: "grid",
            placeItems: "center",
            fontSize: 26,
            fontWeight: 700,
          }}>
          ✓
        </div>
        <h1 style={{ margin: "0 0 10px", fontSize: 22, color: "#004643" }}>Your calendar is connected</h1>
        <p style={{ margin: 0, fontSize: 15, lineHeight: 1.55, color: "#45555c" }}>
          You&rsquo;re all set — you can close this tab. Booked discovery calls will now land directly on your
          calendar, and your availability is kept in sync automatically.
        </p>
      </div>
    </div>
  );
}
