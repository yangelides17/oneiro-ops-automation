import type { Request, Response, NextFunction, RequestHandler } from 'express';

export class AppError extends Error {
  constructor(
    public statusCode: number,
    message: string,
  ) {
    super(message);
    this.name = 'AppError';
  }
}

export function errorHandler(err: Error, _req: Request, res: Response, _next: NextFunction) {
  if (err instanceof AppError) {
    return res.status(err.statusCode).json({ error: err.message });
  }

  // PostgreSQL constraint violations → 409 Conflict
  if ((err as any).code === '23505') {
    return res.status(409).json({ error: 'Already exists' });
  }

  // PostgreSQL invalid input (e.g., non-UUID in UUID column) → 400
  if ((err as any).code === '22P02') {
    return res.status(400).json({ error: 'Invalid input format' });
  }

  console.error('Unhandled error:', err);
  return res.status(500).json({ error: 'Internal server error' });
}
