import { drizzle } from 'drizzle-orm/node-postgres';
import pg from 'pg';
import { config } from '../config.js';
import * as schema from './schema.js';

const pool = new pg.Pool({
  connectionString: config.db.url,
  max: 20,
});

export const db = drizzle(pool, { schema });
export type Db = typeof db;
