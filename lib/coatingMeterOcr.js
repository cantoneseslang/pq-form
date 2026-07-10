const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-2.5-flash';

const SCAN_PROMPT = `You are reading a handheld coating thickness meter LCD (e.g. DELIXI ELECTRIC 涂层测厚仪).
Extract ONLY what is visible on the meter display:
- substrate: the base material indicator, usually top-left (e.g. 铁, 鉄, 非铁, 非鉄, Fe)
- thickness: the main numeric measurement in the center (7-segment style digits, may include one decimal)
- unit: the unit shown beside the number (usually μm or um)

Return JSON with keys: substrate, thickness, unit, confidence ("high" | "medium" | "low"), notes (optional string).
If a field is unreadable, use null for that field and lower confidence.`;

/** 简体字 铁 → 鉄、非铁 → 非鉄 */
export function normalizeSubstrate(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  if (/^非\s*铁$/u.test(text) || /^非\s*鉄$/u.test(text) || /non.?ferrous/i.test(text)) return '非鉄';
  if (/^铁$/u.test(text) || /^鉄$/u.test(text) || /^Fe$/i.test(text) || /ferrous|steel/i.test(text)) return '鉄';
  return text;
}

export function parseThicknessUm(value) {
  if (value == null || value === '') return null;
  const num = typeof value === 'number'
    ? value
    : parseFloat(String(value).replace(/[^\d.]/g, ''));
  return Number.isFinite(num) ? num : null;
}

function stripBase64Prefix(data) {
  return String(data ?? '').replace(/^data:image\/\w+;base64,/, '');
}

function parseGeminiJson(text) {
  const raw = String(text ?? '').trim();
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/);
  const candidate = fenced ? fenced[1].trim() : raw;
  return JSON.parse(candidate);
}

export async function scanCoatingMeterImage({ imageBase64, mimeType = 'image/jpeg' }) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error('GEMINI_API_KEY is not set');

  const data = stripBase64Prefix(imageBase64);
  if (!data) throw new Error('imageBase64 is required');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${encodeURIComponent(apiKey)}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{
        parts: [
          { text: SCAN_PROMPT },
          { inline_data: { mime_type: mimeType, data } },
        ],
      }],
      generationConfig: {
        responseMimeType: 'application/json',
        temperature: 0.1,
      },
    }),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Gemini API error ${res.status}: ${errText.slice(0, 300)}`);
  }

  const payload = await res.json();
  const textPart = payload?.candidates?.[0]?.content?.parts?.find((p) => p.text)?.text;
  if (!textPart) {
    return {
      success: false,
      error: 'Gemini returned no text',
      substrate: null,
      thicknessUm: null,
      unit: null,
      confidence: 'low',
      rawText: '',
    };
  }

  let parsed;
  try {
    parsed = parseGeminiJson(textPart);
  } catch {
    return {
      success: false,
      error: 'Failed to parse Gemini JSON response',
      substrate: null,
      thicknessUm: null,
      unit: null,
      confidence: 'low',
      rawText: textPart,
    };
  }

  const substrate = normalizeSubstrate(parsed.substrate);
  const thicknessUm = parseThicknessUm(parsed.thickness);
  const unitRaw = String(parsed.unit ?? 'μm').trim().toLowerCase();
  const unit = unitRaw === 'um' || unitRaw === 'μm' || unitRaw === 'µm' ? 'μm' : (parsed.unit || 'μm');
  const confidence = ['high', 'medium', 'low'].includes(parsed.confidence)
    ? parsed.confidence
    : (substrate && thicknessUm != null ? 'medium' : 'low');

  const success = Boolean(substrate && thicknessUm != null);

  return {
    success,
    error: success ? undefined : 'Could not read substrate and thickness from meter display',
    substrate: substrate || null,
    thicknessUm,
    unit,
    confidence,
    rawText: textPart,
    notes: parsed.notes || undefined,
  };
}
