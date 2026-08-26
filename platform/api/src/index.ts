import { createApp } from './app.js';
import { config } from './config.js';
import { db } from './db/client.js';
import { sql } from 'drizzle-orm';

const app = createApp();

// Verify database connection before starting
try {
  await db.execute(sql`SELECT 1`);
  console.log('Database connected');
} catch (err) {
  console.error('Failed to connect to database:', err);
  process.exit(1);
}

// Catch unhandled promise rejections so the server doesn't crash
process.on('unhandledRejection', (reason) => {
  console.error('Unhandled rejection:', reason);
});

app.listen(config.port, () => {
  console.log(`API server running on port ${config.port} (${config.nodeEnv})`);
});
