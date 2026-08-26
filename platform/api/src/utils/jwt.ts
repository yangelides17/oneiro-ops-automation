import jwt from 'jsonwebtoken';
import { config } from '../config.js';

export interface JwtPayload {
  userId: string;
  orgId: string;
  role: string;
}

export function signJwt(payload: JwtPayload): string {
  return jwt.sign(payload, config.jwt.secret, {
    expiresIn: config.jwt.expiresIn as jwt.SignOptions['expiresIn'],
  });
}

export function verifyJwt(token: string): JwtPayload {
  return jwt.verify(token, config.jwt.secret) as JwtPayload;
}
