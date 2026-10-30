/** The one error shape of every Graph call: the read graph, the write graph and the commands built on them. */
export type GraphError =
  | { type: 'api_error'; status: number; message: string; code?: string; retryAfterSeconds?: number }
  | { type: 'auth_failed'; message: string; code?: string }
  | { type: 'network_error'; message: string; code?: string }
  | { type: 'validation_error'; message: string; code?: string };
