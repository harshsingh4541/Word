const PREFIX = "dockaro:";

// Every editor session starts from a blank document: a refresh (or reopening
// the URL) is a fresh start, not a resume. Autosave still runs during the
// session so an accidental reload can't wedge the editor mid-edit, but the
// snapshot is never read back — it's discarded on the next mount.
//
// Flip this to true to restore the previous resume-where-you-left-off
// behavior; nothing else needs to change, since both editors seed their
// document from loadSnapshot().
const RESTORE_LAST_SESSION = false;

export function loadSnapshot<T>(key: string): T | null {
  try {
    if (!RESTORE_LAST_SESSION) {
      // Drop the stale snapshot rather than leaving it to sit in storage,
      // otherwise it lingers forever and quietly eats the origin's quota.
      localStorage.removeItem(PREFIX + key);
      return null;
    }
    const raw = localStorage.getItem(PREFIX + key);
    if (!raw) return null;
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

export function saveSnapshot(key: string, data: unknown) {
  try {
    localStorage.setItem(PREFIX + key, JSON.stringify(data));
  } catch {
    // Private browsing / quota exceeded — editing still works, it just
    // won't survive a refresh. Not worth surfacing to the user.
  }
}

export function clearSnapshot(key: string) {
  try {
    localStorage.removeItem(PREFIX + key);
  } catch {
    // ignore
  }
}
