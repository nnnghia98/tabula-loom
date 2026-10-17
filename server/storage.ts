import fs from 'node:fs/promises';
import path from 'node:path';
import { S3Client, PutObjectCommand, GetObjectCommand, DeleteObjectCommand } from '@aws-sdk/client-s3';

const S3_BUCKET =
  process.env.BUCKET ||
  process.env.S3_BUCKET ||
  process.env.AWS_S3_BUCKET_NAME ||
  '';
const S3_ENDPOINT =
  process.env.ENDPOINT ||
  process.env.S3_ENDPOINT ||
  process.env.AWS_ENDPOINT_URL ||
  '';
const S3_REGION =
  process.env.REGION ||
  process.env.S3_REGION ||
  process.env.AWS_DEFAULT_REGION ||
  'auto';
const S3_ACCESS_KEY_ID =
  process.env.ACCESS_KEY_ID ||
  process.env.S3_ACCESS_KEY_ID ||
  process.env.AWS_ACCESS_KEY_ID ||
  '';
const S3_SECRET_ACCESS_KEY =
  process.env.SECRET_ACCESS_KEY ||
  process.env.S3_SECRET_ACCESS_KEY ||
  process.env.AWS_SECRET_ACCESS_KEY ||
  '';

const LOCAL_STORAGE_DIR = process.env.RAILWAY_VOLUME_MOUNT_PATH
  ? path.join(process.env.RAILWAY_VOLUME_MOUNT_PATH, 'workbooks')
  : path.resolve(process.cwd(), 'storage', 'workbooks');

const s3Client =
  S3_BUCKET && S3_ACCESS_KEY_ID && S3_SECRET_ACCESS_KEY
    ? new S3Client({
        region: S3_REGION,
        endpoint: S3_ENDPOINT || undefined,
        forcePathStyle: Boolean(S3_ENDPOINT),
        credentials: {
          accessKeyId: S3_ACCESS_KEY_ID,
          secretAccessKey: S3_SECRET_ACCESS_KEY,
        },
      })
    : null;

export function getStorageMode(): 'railway-s3-bucket' | 'volume-or-disk' {
  return s3Client ? 'railway-s3-bucket' : 'volume-or-disk';
}

export async function saveOriginalXlsxBinary(
  workbookId: string,
  fileName: string,
  buffer: Buffer
): Promise<string> {
  const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
  const storageKey = `${workbookId}/${safeName}`;

  if (s3Client && S3_BUCKET) {
    await s3Client.send(
      new PutObjectCommand({
        Bucket: S3_BUCKET,
        Key: storageKey,
        Body: buffer,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
    );
    return `s3://${S3_BUCKET}/${storageKey}`;
  }

  const targetDir = path.join(LOCAL_STORAGE_DIR, workbookId);
  await fs.mkdir(targetDir, { recursive: true });
  const filePath = path.join(targetDir, safeName);
  await fs.writeFile(filePath, buffer);
  return filePath;
}

export async function readOriginalXlsxBinary(
  storageKey: string
): Promise<{ buffer: Buffer; fileName: string } | null> {
  if (!storageKey) return null;

  if (storageKey.startsWith('s3://') && s3Client && S3_BUCKET) {
    const key = storageKey.replace(`s3://${S3_BUCKET}/`, '');
    const res = await s3Client.send(
      new GetObjectCommand({
        Bucket: S3_BUCKET,
        Key: key,
      })
    );
    if (!res.Body) return null;
    const byteArray = await res.Body.transformToByteArray();
    return {
      buffer: Buffer.from(byteArray),
      fileName: path.basename(key),
    };
  }

  try {
    const buffer = await fs.readFile(storageKey);
    return {
      buffer,
      fileName: path.basename(storageKey),
    };
  } catch {
    return null;
  }
}

export async function deleteOriginalXlsxBinary(storageKey: string | null): Promise<void> {
  if (!storageKey) return;

  if (storageKey.startsWith('s3://') && s3Client && S3_BUCKET) {
    const key = storageKey.replace(`s3://${S3_BUCKET}/`, '');
    try {
      await s3Client.send(
        new DeleteObjectCommand({
          Bucket: S3_BUCKET,
          Key: key,
        })
      );
    } catch {
      // ignore delete errors
    }
    return;
  }

  try {
    await fs.unlink(storageKey);
  } catch {
    // ignore missing file
  }
}
