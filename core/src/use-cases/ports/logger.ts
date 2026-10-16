export type LogMeta = Readonly<Record<string, unknown>>;

export type Logger = {
  readonly info: (event: string, meta?: LogMeta) => void;
  readonly warn: (event: string, meta?: LogMeta) => void;
  readonly error: (event: string, meta?: LogMeta) => void;
};
