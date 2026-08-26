export interface FileMetadata {
  key: string;
  filename: string;
  size: number;
  mimeType: string;
  lastModified: Date;
}

export interface StorageConnector {
  /** Upload a file and return its storage key. */
  upload(orgId: string, path: string, data: Buffer, mime: string): Promise<string>;

  /** Download a file by storage key. */
  download(orgId: string, storageKey: string): Promise<Buffer>;

  /** Get a time-limited signed URL for direct browser access. */
  getSignedUrl(orgId: string, storageKey: string, expiresInSec: number): Promise<string>;

  /** Delete a file by storage key. */
  delete(orgId: string, storageKey: string): Promise<void>;

  /** List files under a prefix. */
  list(orgId: string, prefix: string): Promise<FileMetadata[]>;
}
