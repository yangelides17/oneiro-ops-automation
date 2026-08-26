import crypto from 'crypto';
import { config } from '../config.js';

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;
const TAG_LENGTH = 16;

function getKey(): Buffer {
  const raw = config.encryption.key;
  if (!raw) throw new Error('ENCRYPTION_KEY not configured');
  const key = Buffer.from(raw, 'base64');
  if (key.length !== 32) throw new Error(`ENCRYPTION_KEY must decode to 32 bytes (got ${key.length})`);
  return key;
}

export function encrypt(plaintext: string): string {
  const key = getKey();
  const iv = crypto.randomBytes(IV_LENGTH);
  const cipher = crypto.createCipheriv(ALGORITHM, key, iv);
  const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  // Format: enc:v1:<base64(iv + tag + ciphertext)>
  const combined = Buffer.concat([iv, tag, encrypted]);
  return `enc:v1:${combined.toString('base64')}`;
}

export function decrypt(encoded: string): string {
  if (!encoded.startsWith('enc:v1:')) {
    throw new Error('Unrecognized encryption format');
  }
  const key = getKey();
  const combined = Buffer.from(encoded.slice(7), 'base64');
  const iv = combined.subarray(0, IV_LENGTH);
  const tag = combined.subarray(IV_LENGTH, IV_LENGTH + TAG_LENGTH);
  const ciphertext = combined.subarray(IV_LENGTH + TAG_LENGTH);
  const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
  decipher.setAuthTag(tag);
  return decipher.update(ciphertext) + decipher.final('utf8');
}
