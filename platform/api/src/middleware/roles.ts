import type { Request, Response, NextFunction } from 'express';

export type Role = 'owner' | 'admin' | 'foreman' | 'crew';

/**
 * Middleware factory that restricts access to specific roles.
 * Usage: requireRole('owner', 'admin')
 */
export function requireRole(...allowed: Role[]) {
  return (req: Request, res: Response, next: NextFunction) => {
    const userRole = req.user?.role as Role | undefined;
    if (!userRole || !allowed.includes(userRole)) {
      return res.status(403).json({ error: 'Forbidden' });
    }
    next();
  };
}
