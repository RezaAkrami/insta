import express, { Request, Response } from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';
import crypto from 'crypto';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DEFAULT_BOXAPI_TOKEN = process.env.BOXAPI_TOKEN || '';
const BOXAPI_COMMENTS_URL = 'https://boxapi.ir/api/instagram/media/get_comments';

export interface OnlyCommentItem {
  row: number;
  text: string;
}

interface StoredExcel {
  buffer: Buffer;
  filename: string;
  createdAt: number;
  count: number;
  identifier: string;
}

const excelFilesCache = new Map<string, StoredExcel>();

// Clean up old files after 2 hours
setInterval(() => {
  const now = Date.now();
  for (const [id, item] of excelFilesCache.entries()) {
    if (now - item.createdAt > 2 * 60 * 60 * 1000) {
      excelFilesCache.delete(id);
    }
  }
}, 15 * 60 * 1000);

// Extract shortcode or detect pure numeric media ID
export function parseInputIdentifier(input: string): { type: 'media_id' | 'shortcode'; value: string } | null {
  if (!input) return null;
  const cleaned = input.replace(/[\u200B-\u200F\u202A-\u202E\uFEFF]/g, '').trim();

  // If pure digits (e.g. 2967835229909923253)
  if (/^\d{12,25}$/.test(cleaned)) {
    return { type: 'media_id', value: cleaned };
  }

  // App Share links (instagram.com/share/p/XYZ, /share/reel/XYZ)
  const shareMatch = cleaned.match(/instagram\.com\/share\/(?:p|reel)\/([A-Za-z0-9_-]+)/i);
  if (shareMatch && shareMatch[1]) return { type: 'shortcode', value: shareMatch[1] };

  // Standard post/reel/tv links (instagram.com/p/XYZ, /reel/XYZ, /reels/XYZ, /tv/XYZ)
  const standardMatch = cleaned.match(/(?:instagram\.com|instagr\.am|ig\.me)\/(?:p|reel|reels|tv)\/([A-Za-z0-9_-]+)/i);
  if (standardMatch && standardMatch[1]) return { type: 'shortcode', value: standardMatch[1] };

  // Plain shortcode (7 to 15 alphanumeric characters)
  if (/^[A-Za-z0-9_-]{7,15}$/.test(cleaned)) {
    return { type: 'shortcode', value: cleaned };
  }

  // General URL capture
  const generalMatch = cleaned.match(/instagram\.com\/([A-Za-z0-9_-]{9,14})/i);
  if (generalMatch && generalMatch[1]) return { type: 'shortcode', value: generalMatch[1] };

  return null;
}

// Convert Instagram shortcode to numeric Media ID locally using BigInt (0 API calls, 0 cost)
export function shortcodeToMediaId(shortcode: string): string {
  try {
    let id = BigInt(0);
    const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
    for (let i = 0; i < shortcode.length; i++) {
      const char = shortcode[i];
      const index = BigInt(alphabet.indexOf(char));
      if (index === BigInt(-1)) continue;
      id = id * BigInt(64) + index;
    }
    return id.toString();
  } catch {
    return '';
  }
}

// Check BoxAPI account status & remaining requests
async function getBoxApiAccountInfo(token: string) {
  try {
    const res = await fetch('https://boxapi.ir/api/my', {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token.trim()}`,
        'Content-Type': 'application/json',
      },
    });
    if (res.ok) {
      return await res.json();
    }
  } catch (err) {
    console.error('BoxAPI account check error:', err);
  }
  return null;
}

// Extract ALL comments: pinned comments, top-level comments, child comments, replies, and pagination cursor
export function extractAllCommentsFromBoxApi(data: any): {
  comments: string[];
  nextCursor?: string;
  hasMore: boolean;
} {
  if (!data) return { comments: [], hasMore: false };

  let body = data.response?.body ?? data.body ?? data;
  if (typeof body === 'string') {
    try {
      body = JSON.parse(body);
    } catch {}
  }

  const comments: string[] = [];
  const addText = (txt: any) => {
    if (!txt) return;
    const str = String(txt).trim();
    if (str.length > 0 && !comments.includes(str)) {
      comments.push(str);
    }
  };

  const processItem = (item: any) => {
    if (!item) return;
    if (typeof item === 'string') {
      addText(item);
      return;
    }

    // Top-level comment text
    addText(item.text || item.comment || item.message);

    // Nested replies and child comments
    const childLists = [
      item.preview_child_comments,
      item.child_comments,
      item.edge_threaded_comments?.edges,
    ];

    for (const list of childLists) {
      if (Array.isArray(list)) {
        for (const child of list) {
          const childObj = child.node || child;
          processItem(childObj);
        }
      }
    }
  };

  // 1. Pinned comments (often separated from the main list)
  if (Array.isArray(body?.pinned_comments)) {
    body.pinned_comments.forEach(processItem);
  }

  // 2. Candidate main comment arrays
  const candidateArrays = [
    body?.comments,
    body?.preview_comments,
    body?.data?.comments,
    body?.items,
    body?.data?.items,
    data?.comments,
    data?.data?.comments,
  ];

  for (const arr of candidateArrays) {
    if (Array.isArray(arr) && arr.length > 0) {
      arr.forEach(processItem);
      break;
    }
  }

  // 3. GraphQL edges format
  const edges = body?.data?.shortcode_media?.edge_media_to_parent_comment?.edges 
    || body?.graphql?.shortcode_media?.edge_media_to_parent_comment?.edges;
  if (Array.isArray(edges)) {
    for (const edge of edges) {
      processItem(edge.node || edge);
    }
  }

  // Pagination detection:
  // Instagram comments API uses next_max_id or next_min_id or end_cursor
  const hasMore = Boolean(
    body?.has_more_comments ||
    body?.has_more_headload_comments ||
    body?.data?.shortcode_media?.edge_media_to_parent_comment?.page_info?.has_next_page
  );

  const rawCursor =
    body?.next_max_id ||
    body?.next_min_id ||
    body?.max_id ||
    body?.data?.shortcode_media?.edge_media_to_parent_comment?.page_info?.end_cursor;

  const nextCursor = rawCursor ? String(rawCursor) : undefined;

  return { comments, nextCursor, hasMore };
}

// Format payload for BoxAPI media/get_comments
function formatBoxApiPayload(mediaId: string, cursor?: string): string {
  const idPart = /^\d+$/.test(mediaId) ? `"id":${mediaId}` : `"id":${JSON.stringify(mediaId)}`;
  const parts = [idPart, `"can_support_threading":true`];
  if (cursor) {
    parts.push(`"max_id":${JSON.stringify(cursor)}`);
    parts.push(`"min_id":${JSON.stringify(cursor)}`);
  }
  return `{${parts.join(',')}}`;
}

// Fetch ALL comments with full pagination and reply retrieval
async function fetchAllCommentsFromBoxApi(
  mediaId: string,
  token: string,
  fetchAllPages: boolean = true
): Promise<{
  success: boolean;
  comments: string[];
  pagesFetched: number;
  hasMoreRemaining?: boolean;
  error?: string;
  isLimitReached?: boolean;
  partialDueToLimit?: boolean;
}> {
  const allComments: string[] = [];
  let currentCursor: string | undefined = undefined;
  let page = 0;
  // If user requests all pages, loop up to 25 pages (can gather 500-1000+ comments)
  const maxPages = fetchAllPages ? 25 : 1;

  while (page < maxPages) {
    page++;
    try {
      const payload = formatBoxApiPayload(mediaId, currentCursor);

      const res = await fetch(BOXAPI_COMMENTS_URL, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${token.trim()}`,
          'Content-Type': 'application/json',
        },
        body: payload,
        signal: AbortSignal.timeout(25000),
      });

      const data = await res.json();

      if (data.success === false) {
        const msg = data.data?.message || data.message || 'خطا در برقراری ارتباط با سرویس';
        const isLimit = msg.toLowerCase().includes('limit') || msg.toLowerCase().includes('no more api requests');

        // If we already collected comments on earlier pages, return what we have!
        if (allComments.length > 0) {
          return {
            success: true,
            comments: allComments,
            pagesFetched: page - 1,
            hasMoreRemaining: true,
            partialDueToLimit: isLimit,
          };
        }

        return {
          success: false,
          comments: [],
          pagesFetched: page,
          error: isLimit ? 'اعتبار درخواست‌های حساب شما به پایان رسیده است.' : msg,
          isLimitReached: isLimit,
        };
      }

      const { comments: newBatch, nextCursor, hasMore } = extractAllCommentsFromBoxApi(data);

      let addedCount = 0;
      for (const c of newBatch) {
        if (!allComments.includes(c)) {
          allComments.push(c);
          addedCount++;
        }
      }

      // If no next cursor, or no more comments, or cursor didn't advance, stop pagination
      if (!hasMore || !nextCursor || nextCursor === currentCursor) {
        break;
      }

      // If a page returned 0 new comments, break to avoid infinite loops
      if (addedCount === 0 && page > 1) {
        break;
      }

      currentCursor = nextCursor;
    } catch (err: any) {
      if (allComments.length > 0) {
        return {
          success: true,
          comments: allComments,
          pagesFetched: page - 1,
          hasMoreRemaining: true,
        };
      }
      return {
        success: false,
        comments: [],
        pagesFetched: page,
        error: err.message || 'خطا در اتصال به BoxAPI',
      };
    }
  }

  return {
    success: true,
    comments: allComments,
    pagesFetched: page,
    hasMoreRemaining: Boolean(currentCursor),
  };
}

// Generate Excel file with ONLY comments and row number
export function generateOnlyCommentsExcel(comments: string[]): Buffer {
  if (!comments || comments.length === 0) {
    throw new Error('هیچ کامنتی برای ایجاد فایل اکسل وجود ندارد.');
  }

  const rows = comments.map((text, index) => ({
    'ردیف': index + 1,
    'متن کامنت': text,
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet['!views'] = [{ RTL: true }];
  worksheet['!cols'] = [
    { wch: 10 }, // ردیف
    { wch: 90 }, // متن کامنت
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'کامنت‌ها');

  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
}

// Parse copied text for Smart Paste mode
function parsePastedCommentsText(text: string): string[] {
  if (!text) return [];
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const metaRegex = /^(?:reply|view replies|view all|\d+[smhdw]|like|likes|see translation|translated by instagram|پاسخ|پاسخ دادن|پسند|لایک|مشاهده همه|مشاهده پاسخ‌ها|ترجمه|ارسال|اشتراک|share)\b/i;
  const usernameRegex = /^@?[a-zA-Z0-9._]{3,30}$/;
  const timeRegex = /^\d+\s*(?:ثانیه|دقیقه|ساعت|روز|هفته|ماه|سال|s|m|h|d|w)\b/i;

  const results: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (metaRegex.test(line) || timeRegex.test(line)) continue;
    if (usernameRegex.test(line) && i + 1 < lines.length && !metaRegex.test(lines[i + 1])) continue;

    const cleaned = line.replace(/^@[\w.]+\s*/, '').trim();
    if (cleaned.length >= 2) {
      results.push(cleaned);
    }
  }
  return results;
}

async function startServer() {
  const app = express();
  const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;
  const isProd = process.env.NODE_ENV === 'production';

  app.use(express.json({ limit: '20mb' }));

  // API: Check token credits
  app.get('/api/boxapi-account', async (req: Request, res: Response) => {
    const token = (req.query.token as string)?.trim();
    if (!token) {
      res.status(400).json({ success: false, error: 'لطفاً توکن را وارد کنید.' });
      return;
    }
    const accountInfo = await getBoxApiAccountInfo(token);
    if (accountInfo) {
      res.json({ success: true, account: accountInfo });
    } else {
      res.status(400).json({ success: false, error: 'توکن نامعتبر است یا ارتباط برقرار نشد.' });
    }
  });

  // API: Extract comments (Fetches ALL comments across all pages)
  app.post('/api/extract-comments', async (req: Request, res: Response) => {
    try {
      const { url, rawCommentsText, boxapiToken, fetchAllPages } = req.body;
      const token = (boxapiToken && typeof boxapiToken === 'string' && boxapiToken.trim())
        ? boxapiToken.trim()
        : DEFAULT_BOXAPI_TOKEN;

      let commentsList: string[] = [];
      let identifierUsed = 'custom';
      let pagesFetched = 1;
      let partialDueToLimit = false;

      // Mode 1: Direct text paste
      if (rawCommentsText && typeof rawCommentsText === 'string' && rawCommentsText.trim().length > 0) {
        commentsList = parsePastedCommentsText(rawCommentsText);
        identifierUsed = 'pasted';
      } else {
        // Mode 2: Via URL, Shortcode, or Numeric ID
        if (!token) {
          res.status(400).json({ error: 'لطفاً ابتدا توکن خود را در کادر بالای صفحه وارد کنید.' });
          return;
        }

        if (!url || typeof url !== 'string') {
          res.status(400).json({ error: 'لطفاً لینک پست اینستاگرام، کد پست یا آیدی عددی پست را وارد کنید.' });
          return;
        }

        const parsed = parseInputIdentifier(url);
        if (!parsed) {
          res.status(400).json({
            error: 'فرمت لینک نامعتبر است. نمونه صحیح: https://www.instagram.com/p/DB1234567/ یا آیدی عددی ۲۹۶۷۸۳۵۲۲۹۹۰۹۹۲۳۲۵۳',
          });
          return;
        }

        // Local BigInt conversion (0 API requests wasted)
        let mediaId = '';
        if (parsed.type === 'media_id') {
          mediaId = parsed.value;
          identifierUsed = mediaId;
        } else {
          identifierUsed = parsed.value;
          mediaId = shortcodeToMediaId(parsed.value);
        }

        if (!mediaId) {
          res.status(400).json({ error: 'امکان محاسبه آیدی عددی پست وجود نداشت.' });
          return;
        }

        // Fetch ALL comments across pages (defaults to true for complete extraction)
        const shouldFetchAll = fetchAllPages !== false;
        const boxResult = await fetchAllCommentsFromBoxApi(mediaId, token, shouldFetchAll);

        if (!boxResult.success) {
          if (boxResult.isLimitReached) {
            res.status(429).json({
              error: 'اعتبار درخواست‌های حساب شما به پایان رسیده است. لطفاً توکن جدید وارد کنید یا از حالت جای‌گذاری متن کامنت‌ها استفاده کنید.',
              isLimitReached: true,
            });
            return;
          }

          res.status(400).json({
            error: boxResult.error || 'خطا در دریافت کامنت‌ها از سرویس.',
          });
          return;
        }

        commentsList = boxResult.comments;
        pagesFetched = boxResult.pagesFetched;
        partialDueToLimit = Boolean(boxResult.partialDueToLimit);

        if (commentsList.length === 0) {
          res.status(404).json({
            error: 'کامنتی برای این پست ثبت نشده است یا بخش کامنت‌ها توسط صاحب پیج بسته شده است.',
            mediaId,
          });
          return;
        }
      }

      // Generate strictly comment-only Excel file
      const buffer = generateOnlyCommentsExcel(commentsList);
      const downloadId = crypto.randomUUID();
      const filename = `instagram_comments_${identifierUsed}_only.xlsx`;

      excelFilesCache.set(downloadId, {
        buffer,
        filename,
        createdAt: Date.now(),
        count: commentsList.length,
        identifier: identifierUsed,
      });

      const formattedComments: OnlyCommentItem[] = commentsList.map((text, i) => ({
        row: i + 1,
        text,
      }));

      // Combined plain text of all comments (one per line) ready for 1-click clipboard copy
      const plainTextComments = commentsList.join('\n');

      res.json({
        success: true,
        downloadId,
        downloadUrl: `/api/download/${downloadId}`,
        filename,
        fileSizeBytes: buffer.byteLength,
        fileSizeFormatted: `${(buffer.byteLength / 1024).toFixed(1)} کیلوبایت`,
        totalCount: commentsList.length,
        pagesFetched,
        partialDueToLimit,
        identifier: identifierUsed,
        comments: formattedComments,
        plainText: plainTextComments,
      });
    } catch (err: any) {
      console.error('Error in /api/extract-comments:', err);
      res.status(500).json({ error: err.message || 'خطای غیرمنتظره در سرور رخ داد.' });
    }
  });

  // API: Download Excel
  app.get('/api/download/:fileId', (req: Request, res: Response) => {
    const fileId = req.params.fileId;
    const fileItem = excelFilesCache.get(fileId);

    if (!fileItem || !fileItem.buffer || fileItem.buffer.byteLength === 0) {
      res.status(404).send('فایل اکسل یافت نشد یا منقضی شده است.');
      return;
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(fileItem.filename)}"`);
    res.setHeader('Content-Length', fileItem.buffer.byteLength);
    res.setHeader('Cache-Control', 'no-cache');
    res.end(fileItem.buffer);
  });

  // API: Health
  app.get('/api/health', (_req: Request, res: Response) => {
    res.json({ status: 'healthy', timestamp: new Date().toISOString() });
  });

  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
