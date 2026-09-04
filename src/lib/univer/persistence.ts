const PREFIX = "dockaro:";

export function loadSnapshot<T>(key: string): T | null {
  try {
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
