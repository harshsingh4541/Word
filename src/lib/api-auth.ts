import { NextRequest, NextResponse } from "next/server";

// Demo key store. Swap for a real table (apiKeys: {hash, userId, plan, createdAt})
// once auth/billing is wired up — keys here reset whenever the server restarts.
const DEMO_KEYS = new Set(["dk_test_51H7x9pQwErTyUiOpAsDfGh"]);

export function requireApiKey(req: NextRequest) {
  const header = req.headers.get("authorization") ?? "";
  const key = header.startsWith("Bearer ") ? header.slice(7) : null;

  if (!key || !DEMO_KEYS.has(key)) {
    return NextResponse.json(
      {
        error: {
          code: "unauthorized",
          message:
            "Missing or invalid API key. Pass it as 'Authorization: Bearer dk_live_...'.",
        },
      },
      { status: 401 },
    );
  }

  return null;
}
