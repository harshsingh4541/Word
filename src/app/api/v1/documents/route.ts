import { NextRequest, NextResponse } from "next/server";
import { requireApiKey } from "@/lib/api-auth";
import { createDocument, listDocuments } from "@/lib/document-store";

export async function GET(req: NextRequest) {
  const authError = requireApiKey(req);
  if (authError) return authError;

  return NextResponse.json({ data: listDocuments() });
}

export async function POST(req: NextRequest) {
  const authError = requireApiKey(req);
  if (authError) return authError;

  const body = await req.json().catch(() => null);

  if (!body || typeof body.title !== "string" || !["docx", "xlsx"].includes(body.type)) {
    return NextResponse.json(
      {
        error: {
          code: "invalid_request",
          message: "Expected { type: 'docx' | 'xlsx', title: string }.",
        },
      },
      { status: 400 },
    );
  }

  const doc = createDocument({ type: body.type, title: body.title });
  return NextResponse.json(doc, { status: 201 });
}
