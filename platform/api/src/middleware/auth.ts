import type { Request, Response, NextFunction } from 'express';
import { verifyJwt, type JwtPayload } from '../utils/jwt.js';

declare global {
  namespace Express {
    interface Request {
      user?: JwtPayload;
    }
  }
}

const SESSION_COOKIE = 'session';

export function authMiddleware(req: Request, res: Response, next: NextFunction) {
  const token = req.cookies?.[SESSION_COOKIE];
  if (!token) {
    return res.status(401).json({ error: 'Not authenticated' });
  }

  try {
    req.user = verifyJwt(token);
    next();
  } catch {
    return res.status(401).json({ error: 'Invalid or expired session' });
  }
}

export function optionalAuth(req: Request, _res: Response, next: NextFunction) {
  const token = req.cookies?.[SESSION_COOKIE];
  if (token) {
    try {
      req.user = verifyJwt(token);
    } catch {
      // Invalid token — continue without user
    }
  }
  next();
}
