export function describeOfficeError(error: unknown): string {
  const err = error as {
    code?: string;
    name?: string;
    message?: string;
    debugInfo?: unknown;
  };
  const parts = [
    err.code || err.name,
    err.message,
    err.debugInfo ? `debugInfo=${safeJson(err.debugInfo)}` : "",
  ].filter(Boolean);
  return parts.join(" ");
}

export function withOfficeErrorContext(error: unknown, context: string): Error {
  return new Error(`${context}: ${describeOfficeError(error) || String(error)}`);
}

function safeJson(value: unknown): string {
  try {
    return JSON.stringify(value);
  } catch {
    return String(value);
  }
}
