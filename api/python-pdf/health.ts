import type { IncomingMessage, ServerResponse } from 'node:http';

interface VercelResponse extends ServerResponse {
  status: (code: number) => VercelResponse;
  json: (body: unknown) => void;
}

export default async function handler(_req: IncomingMessage, res: VercelResponse) {
  const pythonApiUrl = (process.env.PYTHON_PDF_API_URL || '').trim().replace(/\/+$/, '');

  if (!pythonApiUrl) {
    res.status(200).json({
      available: false,
      camelot: false,
      pdfplumber: false,
      mode: 'vercel',
    });
    return;
  }

  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 3000);
    const response = await fetch(`${pythonApiUrl}/health`, {
      signal: controller.signal,
    }).finally(() => clearTimeout(timer));

    if (!response.ok) {
      res.status(200).json({ available: false, camelot: false, pdfplumber: false });
      return;
    }

    const data = (await response.json()) as Record<string, unknown>;
    const isPdfApi =
      typeof data?.camelot === 'boolean' && typeof data?.pdfplumber === 'boolean';

    if (!isPdfApi) {
      res.status(200).json({ available: false, camelot: false, pdfplumber: false });
      return;
    }

    res.status(200).json({ available: true, ...data });
  } catch {
    res.status(200).json({ available: false, camelot: false, pdfplumber: false });
  }
}
