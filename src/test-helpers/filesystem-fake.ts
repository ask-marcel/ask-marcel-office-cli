import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

export type FileSystemFake = FileSystem & {
  readonly seed: (path: string, content: string) => void;
  readonly seedBytes: (path: string, bytes: Uint8Array) => void;
  readonly snapshot: (path: string) => string | undefined;
  readonly snapshotBytes: (path: string) => Uint8Array | undefined;
  readonly snapshotMode: (path: string) => number | undefined;
  readonly has: (path: string) => boolean;
};

export const createFileSystemFake = (): FileSystemFake => {
  const store = new Map<string, string>();
  const bytesStore = new Map<string, Uint8Array>();
  const modes = new Map<string, number>();

  return {
    readJson: async <T>(path: string) => {
      const raw = store.get(path);
      if (raw === undefined) return err({ type: 'not_found' });
      try {
        return ok(JSON.parse(raw) as T);
      } catch (e) {
        return err({ type: 'parse_failed', message: e instanceof Error ? e.message : String(e) });
      }
    },
    readBytes: async (path) => {
      const bytes = bytesStore.get(path);
      if (bytes !== undefined) return ok(bytes);
      const text = store.get(path);
      if (text !== undefined) return ok(new TextEncoder().encode(text));
      return err({ type: 'not_found' });
    },
    writeText: async (path, content) => {
      store.set(path, content);
      bytesStore.delete(path);
      return ok(undefined);
    },
    writeBytes: async (path, bytes) => {
      bytesStore.set(path, bytes);
      store.delete(path);
      return ok(undefined);
    },
    chmod: async (path, mode) => {
      if (!store.has(path) && !bytesStore.has(path)) return err({ type: 'io_failed', message: `ENOENT: chmod target missing: ${path}` });
      modes.set(path, mode);
      return ok(undefined);
    },
    deleteIfExists: async (path) => {
      store.delete(path);
      bytesStore.delete(path);
      return ok(undefined);
    },
    // Login-fix round-1: directories aren't modeled (the fake is flat),
    // so wipe every key that has `path/` as a prefix to approximate a
    // recursive delete. Returns ok regardless — matches port contract.
    deleteDirIfExists: async (path) => {
      const prefix = path.endsWith('/') ? path : `${path}/`;
      for (const key of [...store.keys()]) {
        if (key === path || key.startsWith(prefix)) store.delete(key);
      }
      for (const key of [...bytesStore.keys()]) {
        if (key === path || key.startsWith(prefix)) bytesStore.delete(key);
      }
      return ok(undefined);
    },
    seed: (path, content) => {
      store.set(path, content);
    },
    seedBytes: (path, bytes) => {
      bytesStore.set(path, bytes);
    },
    snapshot: (path) => store.get(path),
    snapshotBytes: (path) => bytesStore.get(path),
    snapshotMode: (path) => modes.get(path),
    has: (path) => store.has(path) || bytesStore.has(path),
  };
};
