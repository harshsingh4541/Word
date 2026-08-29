import { NextRequest, NextResponse } from "next/server";
import { requireApiKey } from "@/lib/api-auth";
import { deleteDocument, getDocument, updateDocument } from "@/lib/document-store";

export async function GET(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const authError = requireApiKey(req);
  if (authError) return authError;

  const { id } = await params;
  const doc = getDocument(id);
  if (!doc) {
    return NextResponse.json({ error: { code: "not_found", message: "No such document." } }, { status: 404 });
  }
  return NextResponse.json(doc);
}

export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const authError = requireApiKey(req);
  if (authError) return authError;

  const { id } = await params;
  const body = await req.json().catch(() => null);
  const doc = updateDocument(id, { title: body?.title });
  if (!doc) {
    return NextResponse.json({ error: { code: "not_found", message: "No such document." } }, { status: 404 });
  }
  return NextResponse.json(doc);
}

export async function DELETE(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const authError = requireApiKey(req);
  if (authError) return authError;

  const { id } = await params;
  const ok = deleteDocument(id);
  if (!ok) {
    return NextResponse.json({ error: { code: "not_found", message: "No such document." } }, { status: 404 });
  }
  return new NextResponse(null, { status: 204 });
}
