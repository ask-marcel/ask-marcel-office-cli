import { err, ok } from '../domain/result.ts';
import type { FileSystem } from '../use-cases/ports/filesystem.ts';

export type FileSystemFake = FileSystem & {
  readonly seed: (path: string, content: string) => void;
  readonly snapshot: (path: string) => string | undefined;
  readonly snapshotBytes: (path: string) => Uint8Array | undefined;
  readonly has: (path: string) => boolean;
};

export const createFileSystemFake = (): FileSystemFake => {
  const store = new Map<string, string>();
  const bytesStore = new Map<string, Uint8Array>();

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
    deleteIfExists: async (path) => {
      store.delete(path);
      bytesStore.delete(path);
      return ok(undefined);
    },
    seed: (path, content) => {
      store.set(path, content);
    },
    snapshot: (path) => store.get(path),
    snapshotBytes: (path) => bytesStore.get(path),
    has: (path) => store.has(path) || bytesStore.has(path),
  };
};
