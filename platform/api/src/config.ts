import 'dotenv/config';

function required(key: string): string {
  const val = process.env[key];
  if (!val) throw new Error(`Missing required env var: ${key}`);
  return val;
}

function optional(key: string, fallback: string): string {
  return process.env[key] || fallback;
}

export const config = {
  port: parseInt(optional('PORT', '3001'), 10),
  nodeEnv: optional('NODE_ENV', 'development'),
  isProd: optional('NODE_ENV', 'development') === 'production',

  db: {
    url: required('DATABASE_URL'),
  },

  redis: {
    url: optional('REDIS_URL', 'redis://localhost:6379'),
  },

  jwt: {
    secret: required('JWT_SECRET'),
    expiresIn: optional('JWT_EXPIRES_IN', '7d'),
  },

  r2: {
    accountId: optional('R2_ACCOUNT_ID', ''),
    accessKeyId: optional('R2_ACCESS_KEY_ID', ''),
    secretAccessKey: optional('R2_SECRET_ACCESS_KEY', ''),
    bucketName: optional('R2_BUCKET_NAME', 'oneiro-platform'),
    publicUrl: optional('R2_PUBLIC_URL', ''),
  },

  resend: {
    apiKey: optional('RESEND_API_KEY', ''),
    fromEmail: optional('RESEND_FROM_EMAIL', 'noreply@oneiro.app'),
  },

  googleMaps: {
    apiKey: optional('GOOGLE_MAPS_API_KEY', ''),
  },

  anthropic: {
    apiKey: optional('ANTHROPIC_API_KEY', ''),
  },

  encryption: {
    key: optional('ENCRYPTION_KEY', ''),
  },
} as const;
