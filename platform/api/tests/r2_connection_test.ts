/**
 * R2 Connection Test — verifies upload, download, list, delete
 * Usage: npx tsx tests/r2_connection_test.ts
 */
import 'dotenv/config';
import {
  S3Client, PutObjectCommand, GetObjectCommand,
  DeleteObjectCommand, ListObjectsV2Command,
} from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import { config } from '../src/config.js';

async function run() {
  console.log('\n═══ R2 CONNECTION TEST ═══\n');
  console.log(`  Account:  ${config.r2.accountId}`);
  console.log(`  Bucket:   ${config.r2.bucketName}`);
  console.log(`  Endpoint: https://${config.r2.accountId}.r2.cloudflarestorage.com\n`);

  const client = new S3Client({
    region: 'auto',
    endpoint: `https://${config.r2.accountId}.r2.cloudflarestorage.com`,
    credentials: {
      accessKeyId: config.r2.accessKeyId,
      secretAccessKey: config.r2.secretAccessKey,
    },
  });

  const BUCKET = config.r2.bucketName;
  const TEST_KEY = '_test/connection_check.txt';
  const TEST_CONTENT = `R2 connection test — ${new Date().toISOString()}`;

  try {
    // 1. Upload
    await client.send(new PutObjectCommand({
      Bucket: BUCKET,
      Key: TEST_KEY,
      Body: Buffer.from(TEST_CONTENT),
      ContentType: 'text/plain',
    }));
    console.log('  ✓ Upload works');

    // 2. Download
    const resp = await client.send(new GetObjectCommand({ Bucket: BUCKET, Key: TEST_KEY }));
    const chunks: Uint8Array[] = [];
    for await (const chunk of resp.Body as AsyncIterable<Uint8Array>) chunks.push(chunk);
    const content = Buffer.concat(chunks).toString();
    console.log(`  ✓ Download works (${content.length} bytes)`);

    if (content !== TEST_CONTENT) {
      console.error('  ✗ Content mismatch!');
      process.exit(1);
    }
    console.log('  ✓ Content matches');

    // 3. Signed URL
    const url = await getSignedUrl(client, new GetObjectCommand({ Bucket: BUCKET, Key: TEST_KEY }), { expiresIn: 60 });
    console.log(`  ✓ Signed URL generated (${url.slice(0, 80)}...)`);

    // 4. List
    const list = await client.send(new ListObjectsV2Command({ Bucket: BUCKET, Prefix: '_test/' }));
    console.log(`  ✓ List works (${list.Contents?.length || 0} objects under _test/)`);

    // 5. Delete
    await client.send(new DeleteObjectCommand({ Bucket: BUCKET, Key: TEST_KEY }));
    console.log('  ✓ Delete works');

    console.log('\n  ✅ R2 connection fully working!\n');
  } catch (err: any) {
    console.error(`\n  ✗ R2 connection failed: ${err.message}\n`);
    if (err.Code) console.error(`    Error code: ${err.Code}`);
    if (err.$metadata) console.error(`    HTTP status: ${err.$metadata.httpStatusCode}`);
    process.exit(1);
  }
}

run();
