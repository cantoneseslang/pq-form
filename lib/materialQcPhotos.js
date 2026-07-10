import { getSupabaseAdmin, isSupabaseConfigured } from './supabase.js';

export const RAW_MATERIAL_QC_STORAGE_BUCKET = process.env.RAW_MATERIAL_QC_STORAGE_BUCKET || 'material-qc';

function stripBase64Prefix(data) {
  return String(data ?? '').replace(/^data:image\/\w+;base64,/, '');
}

function sanitizeFilePart(value) {
  return String(value ?? 'unknown').trim().replace(/[^\w.-]+/g, '_').slice(0, 40) || 'unknown';
}

function normalizeMimeType(mimeType) {
  const raw = String(mimeType || 'image/jpeg').split(';')[0].trim().toLowerCase();
  if (raw === 'image/png') return 'image/png';
  if (raw === 'image/webp') return 'image/webp';
  return 'image/jpeg';
}

async function ensureBucket(supabase, bucket) {
  const { data: buckets, error: listError } = await supabase.storage.listBuckets();
  if (listError) throw new Error(`Storage listBuckets failed: ${listError.message}`);
  if ((buckets || []).some((b) => b.name === bucket)) return;

  const { error: createError } = await supabase.storage.createBucket(bucket, {
    public: true,
    fileSizeLimit: 5 * 1024 * 1024,
  });
  if (createError && !/already exists/i.test(createError.message)) {
    throw new Error(`Storage createBucket failed: ${createError.message}`);
  }
}

export async function uploadQcPhotoToSupabase({
  imageBase64,
  mimeType = 'image/jpeg',
  receiptDate,
  lotNo,
}) {
  if (!isSupabaseConfigured()) {
    throw new Error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are not set');
  }

  const data = stripBase64Prefix(imageBase64);
  if (!data) throw new Error('imageBase64 is empty');

  const supabase = getSupabaseAdmin();
  const bucket = RAW_MATERIAL_QC_STORAGE_BUCKET;
  await ensureBucket(supabase, bucket);

  const safeMime = normalizeMimeType(mimeType);
  const ext = safeMime === 'image/png' ? 'png' : safeMime === 'image/webp' ? 'webp' : 'jpg';
  const timestamp = Date.now();
  const objectPath = `${sanitizeFilePart(receiptDate || 'unknown')}/${sanitizeFilePart(lotNo || 'lot')}_${timestamp}.${ext}`;
  const buffer = Buffer.from(data, 'base64');

  const { error: uploadError } = await supabase.storage.from(bucket).upload(objectPath, buffer, {
    contentType: safeMime,
    upsert: false,
  });
  if (uploadError) throw new Error(`Storage upload failed: ${uploadError.message}`);

  const { data: publicData } = supabase.storage.from(bucket).getPublicUrl(objectPath);
  const photoUrl = publicData?.publicUrl;
  if (!photoUrl) throw new Error('Storage upload returned no public URL');

  return { photoUrl, objectPath, bucket };
}
