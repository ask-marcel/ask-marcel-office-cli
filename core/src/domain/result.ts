type Ok<T> = { readonly ok: true; readonly value: T };
type Err<E> = { readonly ok: false; readonly error: E };

export type Result<T, E> = Ok<T> | Err<E>;

export const ok = <T>(value: T): Result<T, never> => ({ ok: true, value });
export const err = <E>(error: E): Result<never, E> => ({ ok: false, error });
export const isOk = <T, E>(r: Result<T, E>): r is Ok<T> => r.ok;
export const isErr = <T, E>(r: Result<T, E>): r is Err<E> => !r.ok;

export const unwrap = <T, E>(r: Result<T, E>): T => {
  if (!r.ok) throw new Error(`unwrap called on Err: ${JSON.stringify(r.error)}`);
  return r.value;
};

export const unwrapOr = <T, E>(r: Result<T, E>, fallback: T): T => (r.ok ? r.value : fallback);

export const map = <T, E, U>(r: Result<T, E>, fn: (value: T) => U): Result<U, E> => (r.ok ? ok(fn(r.value)) : r);

export const mapErr = <T, E, F>(r: Result<T, E>, fn: (error: E) => F): Result<T, F> => (r.ok ? r : err(fn(r.error)));

export const andThen = <T, E, U>(r: Result<T, E>, fn: (value: T) => Result<U, E>): Result<U, E> => (r.ok ? fn(r.value) : r);
