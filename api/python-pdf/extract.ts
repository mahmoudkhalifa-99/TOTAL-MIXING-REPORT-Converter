import type { IncomingMessage, ServerResponse } from 'node:http';

interface VercelRequest extends IncomingMessage {
  body?: unknown;
}

interface VercelResponse extends ServerResponse {
  status: (code: number) => VercelResponse;
  json: (body: unknown) => void;
}

async function readJsonBody(req: VercelRequest): Promise<Record<string, unknown>> {
  if (req.body && typeof req.body === 'object') {
    return req.body as Record<string, unknown>;
  }
  if (typeof req.body === 'string') {
    return JSON.parse(req.body) as Record<string, unknown>;
  }

  const chunks: Buffer[] = [];
  for await (const chunk of req) {
    chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
  }
  const raw = Buffer.concat(chunks).toString('utf8');
  return raw ? (JSON.parse(raw) as Record<string, unknown>) : {};
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed. Use POST.' });
    return;
  }

  const pythonApiUrl = (process.env.PYTHON_PDF_API_URL || '').trim().replace(/\/+$/, '');
  if (!pythonApiUrl) {
    res.status(503).json({
      error:
        'خادم Python FastAPI غير مربوط على Vercel (متغير البيئة PYTHON_PDF_API_URL غير مضبوط). استخدم محرك المتصفح المحلي (pdfjs-dist + Tesseract.js) الذي يعمل مجاناً وبشكل كامل مباشرة على Vercel.',
    });
    return;
  }

  try {
    const body = await readJsonBody(req);
    const fileBase64 = body.fileBase64;
    const fileName = typeof body.fileName === 'string' ? body.fileName : 'document.pdf';
    const pages = typeof body.pages === 'string' ? body.pages : 'all';
    const flavor = typeof body.flavor === 'string' ? body.flavor : 'auto';
    const mergeMultipage = body.mergeMultipage !== false;

    if (!fileBase64 || typeof fileBase64 !== 'string') {
      res.status(400).json({ error: 'لم يتم إرسال بيانات ملف PDF.' });
      return;
    }

    const cleanB64 = fileBase64.replace(/^data:[^;]+;base64,/, '');
    const buffer = Buffer.from(cleanB64, 'base64');

    const formData = new FormData();
    formData.append('file', new Blob([buffer], { type: 'application/pdf' }), fileName);
    formData.append('pages', pages || 'all');
    formData.append('flavor', flavor || 'auto');
    formData.append('merge_multipage', mergeMultipage ? 'true' : 'false');

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 55000);

    const response = await fetch(`${pythonApiUrl}/extract`, {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    }).finally(() => clearTimeout(timer));

    const text = await response.text();
    let payload: unknown = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = { detail: text };
    }

    if (!response.ok) {
      const errDetail =
        (payload as Record<string, unknown> | null)?.detail ||
        (payload as Record<string, unknown> | null)?.error ||
        'فشل استخراج الجداول من خادم Python FastAPI.';
      res.status(response.status).json({ error: String(errDetail) });
      return;
    }

    res.status(200).json(payload);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'تعذر الاتصال بخادم Python FastAPI.';
    res.status(503).json({
      error: `تعذر الوصول إلى خادم Python FastAPI: ${message}`,
    });
  }
}
