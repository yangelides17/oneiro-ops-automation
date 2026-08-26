import type { Request, Response, NextFunction } from 'express';

/**
 * Ensures req.user is set and has an orgId.
 * Should be used after authMiddleware.
 * Provides a convenience getter for the org-scoped queries.
 */
export function tenantMiddleware(req: Request, res: Response, next: NextFunction) {
  if (!req.user?.orgId) {
    return res.status(403).json({ error: 'No organization context' });
  }
  next();
}

/** Helper to extract orgId from authenticated request. */
export function getOrgId(req: Request): string {
  return req.user!.orgId;
}

/** Helper to extract userId from authenticated request. */
export function getUserId(req: Request): string {
  return req.user!.userId;
}
