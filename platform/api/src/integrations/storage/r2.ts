import {
  S3Client,
  PutObjectCommand,
  GetObjectCommand,
  DeleteObjectCommand,
  ListObjectsV2Command,
} from '@aws-sdk/client-s3';
import { getSignedUrl as awsGetSignedUrl } from '@aws-sdk/s3-request-presigner';
import { config } from '../../config.js';
import type { StorageConnector, FileMetadata } from './interface.js';

function createClient(): S3Client {
  if (!config.r2.accountId || !config.r2.accessKeyId || !config.r2.secretAccessKey) {
    throw new Error('R2 storage not configured: R2_ACCOUNT_ID, R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY required');
  }
  return new S3Client({
    region: 'auto',
    endpoint: `https://${config.r2.accountId}.r2.cloudflarestorage.com`,
    credentials: {
      accessKeyId: config.r2.accessKeyId,
      secretAccessKey: config.r2.secretAccessKey,
    },
  });
}

let client: S3Client | null = null;
function getClient(): S3Client {
  if (!client) client = createClient();
  return client;
}

/**
 * All files are stored under the org's UUID prefix:
 *   {orgId}/{path}
 * This provides natural tenant isolation at the storage level.
 */
function fullKey(orgId: string, path: string): string {
  return `${orgId}/${path}`;
}

export const r2Storage: StorageConnector = {
  async upload(orgId, path, data, mime) {
    const key = fullKey(orgId, path);
    try {
      await getClient().send(new PutObjectCommand({
        Bucket: config.r2.bucketName,
        Key: key,
        Body: data,
        ContentType: mime,
      }));
      return key;
    } catch (err) {
      console.error(`[R2] Upload failed for ${key}:`, err);
      throw err;
    }
  },

  async download(orgId, storageKey) {
    try {
      const resp = await getClient().send(new GetObjectCommand({
        Bucket: config.r2.bucketName,
        Key: storageKey,
      }));
      const chunks: Uint8Array[] = [];
      for await (const chunk of resp.Body as AsyncIterable<Uint8Array>) {
        chunks.push(chunk);
      }
      return Buffer.concat(chunks);
    } catch (err) {
      console.error(`[R2] Download failed for ${storageKey}:`, err);
      throw err;
    }
  },

  async getSignedUrl(orgId, storageKey, expiresInSec) {
    const command = new GetObjectCommand({
      Bucket: config.r2.bucketName,
      Key: storageKey,
    });
    return awsGetSignedUrl(getClient(), command, { expiresIn: expiresInSec });
  },

  async delete(orgId, storageKey) {
    try {
      await getClient().send(new DeleteObjectCommand({
        Bucket: config.r2.bucketName,
        Key: storageKey,
      }));
    } catch (err) {
      console.error(`[R2] Delete failed for ${storageKey}:`, err);
      // Don't throw on delete failures — fire-and-forget is acceptable
    }
  },

  async list(orgId, prefix) {
    const fullPrefix = fullKey(orgId, prefix);
    const resp = await getClient().send(new ListObjectsV2Command({
      Bucket: config.r2.bucketName,
      Prefix: fullPrefix,
    }));

    return (resp.Contents || []).map((obj) => ({
      key: obj.Key!,
      filename: obj.Key!.split('/').pop()!,
      size: obj.Size || 0,
      mimeType: '',
      lastModified: obj.LastModified || new Date(),
    }));
  },
};
