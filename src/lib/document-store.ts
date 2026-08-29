export type DocKaroDocument = {
  id: string;
  type: "docx" | "xlsx";
  title: string;
  createdAt: string;
  updatedAt: string;
  editUrl: string;
};

// In-memory store for the API skeleton — replace with a real database
// (Postgres/Mongo) before taking traffic. Resets on every server restart.
const documents = new Map<string, DocKaroDocument>();

export function createDocument(input: { type: "docx" | "xlsx"; title: string }) {
  const id = `doc_${crypto.randomUUID().replace(/-/g, "").slice(0, 20)}`;
  const now = new Date().toISOString();
  const doc: DocKaroDocument = {
    id,
    type: input.type,
    title: input.title,
    createdAt: now,
    updatedAt: now,
    editUrl: `https://dockaro.com/e/${id}`,
  };
  documents.set(id, doc);
  return doc;
}

export function listDocuments() {
  return Array.from(documents.values()).sort((a, b) =>
    b.createdAt.localeCompare(a.createdAt),
  );
}

export function getDocument(id: string) {
  return documents.get(id) ?? null;
}

export function updateDocument(id: string, patch: Partial<Pick<DocKaroDocument, "title">>) {
  const existing = documents.get(id);
  if (!existing) return null;
  const updated = { ...existing, ...patch, updatedAt: new Date().toISOString() };
  documents.set(id, updated);
  return updated;
}

export function deleteDocument(id: string) {
  return documents.delete(id);
}
