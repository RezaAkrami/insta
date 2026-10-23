import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Send,
  Link as LinkIcon,
  CheckCircle2,
  AlertCircle,
  Loader2,
  Trash2,
  Search,
  FileText,
  Copy,
  Check,
  Download
} from 'lucide-react';

interface OnlyCommentItem {
  row: number;
  text: string;
}

interface ExtractionResult {
  downloadId: string;
  downloadUrl: string;
  filename: string;
  fileSizeBytes: number;
  fileSizeFormatted: string;
  totalCount: number;
  pagesFetched: number;
  identifier: string;
  comments: OnlyCommentItem[];
  plainText: string;
}

interface ApiAccount {
  username?: string;
  request_used?: number;
  request_limit?: number;
  request_left?: number;
  plan_expiration?: string;
}

export default function App() {
  const [postUrl, setPostUrl] = useState('');
  const [apiToken, setApiToken] = useState('');
  const [accountInfo, setAccountInfo] = useState<ApiAccount | null>(null);
  const [isCheckingAccount, setIsCheckingAccount] = useState(false);
  const [tokenSavedNotice, setTokenSavedNotice] = useState(false);

  const [inputMode, setInputMode] = useState<'url' | 'rawText'>('url');
  const [rawText, setRawText] = useState('');
  const [fetchAllPages, setFetchAllPages] = useState(true);

  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [result, setResult] = useState<ExtractionResult | null>(null);
  const [filterQuery, setFilterQuery] = useState('');
  const [copiedAll, setCopiedAll] = useState(false);
  const [copiedWithRows, setCopiedWithRows] = useState(false);
  const [viewMode, setViewMode] = useState<'list' | 'textarea'>('list');

  // Check account balance
  const checkAccountBalance = async (tokenToCheck?: string) => {
    const t = (tokenToCheck !== undefined ? tokenToCheck : apiToken).trim();
    if (!t) {
      setError('لطفاً ابتدا توکن خود را در کادر وارد کنید.');
      return;
    }

    setIsCheckingAccount(true);
    try {
      const res = await fetch(`/api/boxapi-account?token=${encodeURIComponent(t)}`);
      const data = await res.json();
      if (data.success && data.account) {
        setAccountInfo(data.account);
        setTokenSavedNotice(true);
        setTimeout(() => setTokenSavedNotice(false), 3000);
        setError(null);
      } else {
        setError(data.error || 'توکن نامعتبر است یا ارتباط برقرار نشد.');
      }
    } catch {
      setError('خطا در استعلام وضعیت توکن.');
    } finally {
      setIsCheckingAccount(false);
    }
  };

  // Submit and Extract ALL comments
  const handleExtractComments = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();

    if (inputMode === 'url') {
      if (!apiToken.trim()) {
        setError('لطفاً ابتدا توکن خود را در کادر بالای صفحه وارد کنید.');
        return;
      }
      if (!postUrl.trim()) {
        setError('لطفاً لینک پست اینستاگرام یا آیدی عددی آن را وارد کنید.');
        return;
      }
    }

    if (inputMode === 'rawText' && !rawText.trim()) {
      setError('لطفاً متن کامنت‌ها را در کادر وارد کنید.');
      return;
    }

    setIsLoading(true);
    setError(null);
    setCopiedAll(false);
    setCopiedWithRows(false);

    try {
      const payload: any = {
        boxapiToken: apiToken.trim(),
        fetchAllPages,
      };

      if (inputMode === 'rawText') {
        payload.rawCommentsText = rawText.trim();
      } else {
        payload.url = postUrl.trim();
      }

      const response = await fetch('/api/extract-comments', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'خطایی در استخراج کامنت‌ها رخ داد.');
      }

      if (!data.comments || data.comments.length === 0) {
        throw new Error('هیچ کامنتی برای این پست یافت نشد.');
      }

      setResult(data);
      checkAccountBalance();
    } catch (err: any) {
      setError(err.message || 'خطا در ارتباط با سرور.');
    } finally {
      setIsLoading(false);
    }
  };

  // Copy ALL comments to clipboard
  const handleCopyAllComments = async () => {
    if (!result || !result.comments || result.comments.length === 0) return;

    const commentsToCopy = filterQuery
      ? result.comments.filter((c) => c.text.toLowerCase().includes(filterQuery.toLowerCase()))
      : result.comments;

    const text = commentsToCopy.map((c) => c.text).join('\n');

    try {
      await navigator.clipboard.writeText(text);
      setCopiedAll(true);
      setTimeout(() => setCopiedAll(false), 3500);
    } catch {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      setCopiedAll(true);
      setTimeout(() => setCopiedAll(false), 3500);
    }
  };

  // Copy with row numbers
  const handleCopyWithRowNumbers = async () => {
    if (!result || !result.comments || result.comments.length === 0) return;

    const commentsToCopy = filterQuery
      ? result.comments.filter((c) => c.text.toLowerCase().includes(filterQuery.toLowerCase()))
      : result.comments;

    const text = commentsToCopy.map((c, i) => `${i + 1}\t${c.text}`).join('\n');

    try {
      await navigator.clipboard.writeText(text);
      setCopiedWithRows(true);
      setTimeout(() => setCopiedWithRows(false), 3500);
    } catch {
      // Fallback
    }
  };

  // Download Excel (.xlsx)
  const handleDownloadExcel = () => {
    if (!result || !result.comments || result.comments.length === 0) return;

    const sheetData = result.comments.map((c, idx) => ({
      'ردیف': idx + 1,
      'متن کامنت': c.text,
    }));

    const worksheet = XLSX.utils.json_to_sheet(sheetData);
    worksheet['!views'] = [{ RTL: true }];
    worksheet['!cols'] = [{ wch: 10 }, { wch: 90 }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'کامنت‌ها');

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const downloadUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = result.filename || `instagram_comments_only.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(downloadUrl);
  };

  const filteredComments = (result?.comments || []).filter((item) =>
    item.text.toLowerCase().includes(filterQuery.toLowerCase())
  );

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col selection:bg-emerald-500 selection:text-white overflow-x-hidden">
      {/* Top Header - Mobile Responsive & No Horizontal Scroll */}
      <header className="sticky top-0 z-30 bg-white/95 backdrop-blur-md border-b border-slate-200 w-full">
        <div className="max-w-4xl mx-auto px-4 sm:px-6 h-14 sm:h-16 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-600 flex items-center justify-center text-white shadow-sm font-bold text-xs shrink-0">
              IG
            </div>
            <span className="text-sm sm:text-base font-bold tracking-tight text-slate-900 truncate">
              استخراج تمامی کامنت‌های اینستاگرام
            </span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-4xl mx-auto w-full px-4 sm:px-6 py-6 sm:py-8 flex-1">
        {/* Title */}
        <div className="text-center mb-6">
          <h1 className="text-xl sm:text-3xl font-extrabold text-slate-900 tracking-tight mb-2">
            استخراج تمامی کامنت‌های پست اینستاگرام
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 max-w-xl mx-auto leading-relaxed">
            توکن و لینک پست را در کادرهای زیر وارد کنید. تمامی کامنت‌ها دریافت شده و در کادر پایین با امکان کپی ۱-کلیکه جهت انتقال به اکسل قرار می‌گیرند.
          </p>
        </div>

        {/* TOKEN INPUT & BALANCE */}
        <div className="mb-6 p-4 sm:p-5 bg-white rounded-2xl border border-slate-200 shadow-sm">
          {/* Token Input and Check Button */}
          <div className="space-y-3">
            <div>
              <label htmlFor="token-input" className="block text-xs font-semibold text-slate-700 mb-1.5">
                توکن اختصاصی شما (API Token):
              </label>
              <div className="flex flex-col sm:flex-row gap-2">
                <div className="relative flex-1">
                  <input
                    id="token-input"
                    type="text"
                    dir="ltr"
                    value={apiToken}
                    onChange={(e) => {
                      setApiToken(e.target.value);
                      if (error) setError(null);
                    }}
                    placeholder="توکن خود را اینجا پیست کنید..."
                    className="w-full h-11 px-3.5 text-xs sm:text-sm font-mono text-slate-900 bg-slate-50/80 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-600 transition-all placeholder:font-sans placeholder:text-slate-400"
                  />
                </div>

                <button
                  type="button"
                  onClick={() => checkAccountBalance()}
                  disabled={isCheckingAccount}
                  className="h-11 px-5 bg-slate-900 hover:bg-slate-800 active:bg-slate-950 text-white font-bold text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer shrink-0"
                >
                  {isCheckingAccount ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin text-emerald-400" />
                      <span>در حال استعلام...</span>
                    </>
                  ) : (
                    <>
                      <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                      <span>ثبت و بررسی توکن</span>
                    </>
                  )}
                </button>
              </div>
            </div>

            {tokenSavedNotice && (
              <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-lg text-xs text-emerald-800 font-semibold flex items-center gap-1.5 animate-in fade-in duration-200">
                <Check className="w-4 h-4 text-emerald-600" />
                <span>وضعیت توکن با موفقیت استعلام شد.</span>
              </div>
            )}

            {/* Live Request Balance Cards */}
            {accountInfo && (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-1 animate-in fade-in duration-200">
                <div className={`p-2.5 sm:p-3 rounded-xl border ${
                  accountInfo.request_left === 0
                    ? 'bg-rose-50 border-rose-200'
                    : 'bg-emerald-50/80 border-emerald-200'
                }`}>
                  <div className="text-[11px] text-slate-600 mb-0.5">درخواست‌های باقی‌مانده</div>
                  <div className={`text-lg sm:text-xl font-extrabold font-mono tabular-nums ${
                    accountInfo.request_left === 0 ? 'text-rose-700' : 'text-emerald-700'
                  }`}>
                    {accountInfo.request_left}
                  </div>
                  <div className="text-[10px] text-slate-500 mt-0.5">
                    {accountInfo.request_left === 0 ? 'نیاز به شارژ یا توکن جدید' : 'آماده استفاده'}
                  </div>
                </div>

                <div className="p-2.5 sm:p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <div className="text-[11px] text-slate-600 mb-0.5">درخواست‌های مصرف‌شده</div>
                  <div className="text-lg sm:text-xl font-extrabold font-mono tabular-nums text-slate-900">
                    {accountInfo.request_used}
                  </div>
                  <div className="text-[10px] text-slate-400 mt-0.5">تعداد فراخوانی‌ها</div>
                </div>

                <div className="p-2.5 sm:p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <div className="text-[11px] text-slate-600 mb-0.5">سقف کل پلن</div>
                  <div className="text-lg sm:text-xl font-extrabold font-mono tabular-nums text-slate-900">
                    {accountInfo.request_limit}
                  </div>
                  <div className="text-[10px] text-slate-400 mt-0.5">مجموع اعتبار</div>
                </div>

                <div className="p-2.5 sm:p-3 bg-slate-50 rounded-xl border border-slate-200">
                  <div className="text-[11px] text-slate-600 mb-0.5">حساب کاربری</div>
                  <div className="text-xs sm:text-sm font-extrabold font-mono tabular-nums text-slate-900 truncate">
                    {accountInfo.username || '---'}
                  </div>
                  <div className="text-[10px] text-slate-400 mt-1 truncate">
                    {accountInfo.plan_expiration ? `انقضا: ${accountInfo.plan_expiration.substring(0, 10)}` : 'فعال'}
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Mode Switcher */}
        <div className="flex items-center justify-center gap-2 mb-4">
          <div className="p-1 bg-slate-200/80 rounded-xl flex items-center gap-1 text-xs sm:text-sm font-medium w-full sm:w-auto">
            <button
              type="button"
              onClick={() => {
                setInputMode('url');
                setError(null);
              }}
              className={`flex-1 sm:flex-initial px-3 sm:px-4 py-2 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
                inputMode === 'url'
                  ? 'bg-white text-slate-900 shadow-sm font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <LinkIcon className="w-4 h-4 text-emerald-600" />
              <span>استخراج از طریق لینک پست</span>
            </button>
            <button
              type="button"
              onClick={() => {
                setInputMode('rawText');
                setError(null);
              }}
              className={`flex-1 sm:flex-initial px-3 sm:px-4 py-2 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer ${
                inputMode === 'rawText'
                  ? 'bg-white text-slate-900 shadow-sm font-semibold'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileText className="w-4 h-4 text-emerald-600" />
              <span>جای‌گذاری متن کامنت‌ها</span>
            </button>
          </div>
        </div>

        {/* Input Card with Text Box and Submit Button */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 sm:p-6 mb-8">
          <form onSubmit={handleExtractComments} className="space-y-4">
            {inputMode === 'url' ? (
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label htmlFor="instagram-url" className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-1.5">
                    <LinkIcon className="w-4 h-4 text-emerald-600" />
                    <span>لینک پست یا ریلز اینستاگرام (یا آیدی عددی پست)</span>
                  </label>
                  {postUrl && (
                    <button
                      type="button"
                      onClick={() => {
                        setPostUrl('');
                        setError(null);
                      }}
                      className="text-xs text-slate-400 hover:text-rose-600 transition-colors flex items-center gap-1 cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                      پاک کردن
                    </button>
                  )}
                </div>

                <div className="relative flex items-center">
                  <input
                    id="instagram-url"
                    type="text"
                    dir="ltr"
                    placeholder="https://www.instagram.com/p/... یا آیدی عددی پست"
                    value={postUrl}
                    onChange={(e) => {
                      setPostUrl(e.target.value);
                      if (error) setError(null);
                    }}
                    className="w-full h-12 sm:h-14 px-3.5 sm:px-4 text-left font-mono text-xs sm:text-sm text-slate-900 bg-slate-50/70 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-600 transition-all placeholder:text-slate-400 placeholder:font-sans"
                  />
                </div>

                {/* Option for pagination */}
                <div className="mt-3 p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-2 text-xs">
                  <div className="flex items-center gap-2">
                    <input
                      id="fetch-all-pages"
                      type="checkbox"
                      checked={fetchAllPages}
                      onChange={(e) => setFetchAllPages(e.target.checked)}
                      className="rounded border-slate-300 text-emerald-600 focus:ring-emerald-500 cursor-pointer w-4 h-4 shrink-0"
                    />
                    <label htmlFor="fetch-all-pages" className="cursor-pointer select-none font-bold text-slate-800">
                      پیمایش خودکار تمام صفحات برای استخراج ۱۰۰٪ کامنت‌ها
                    </label>
                  </div>
                  <p className="text-[11px] text-slate-500 pr-6 leading-relaxed">
                    {fetchAllPages ? (
                      <span className="text-amber-800">
                        ⚡ <strong>نحوه کسر اعتبار:</strong> اینستاگرام در هر صفحه بین ۲۰ تا ۵۰ کامنت برمی‌گرداند. با فعال بودن این گزینه، سیستم برای دریافت هر صفحه جدید ۱ درخواست ارسال می‌کند (مثلاً یک پست ۶۰ کامنتی حدود ۲ الی ۳ درخواست از سهمیه شما مصرف می‌کند).
                      </span>
                    ) : (
                      <span className="text-emerald-700">
                        ✓ <strong>صرفه‌جویی در اعتبار:</strong> فقط صفحه اول استخراج شده و <strong>دقیقاً ۱ درخواست</strong> از سهمیه شما کسر می‌شود.
                      </span>
                    )}
                  </p>
                </div>
              </div>
            ) : (
              <div>
                <div className="flex items-center justify-between mb-2">
                  <label htmlFor="raw-comments" className="text-xs sm:text-sm font-bold text-slate-800 flex items-center gap-1.5">
                    <FileText className="w-4 h-4 text-emerald-600" />
                    <span>متن کامنت‌ها را اینجا پیست کنید</span>
                  </label>
                </div>

                <textarea
                  id="raw-comments"
                  rows={6}
                  value={rawText}
                  onChange={(e) => {
                    setRawText(e.target.value);
                    if (error) setError(null);
                  }}
                  placeholder="متن کامنت‌های کپی‌شده از اینستاگرام را مستقیماً اینجا پیست کنید. سیستم تمام آیدی‌ها و متن‌های اضافی را حذف کرده و فقط کامنت‌ها را آماده می‌کند."
                  className="w-full p-3 text-xs sm:text-sm text-slate-900 bg-slate-50/70 border border-slate-300 rounded-xl focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-600 transition-all placeholder:text-slate-400 font-sans leading-relaxed"
                />
              </div>
            )}

            {/* Error Message */}
            {error && (
              <div className="p-3.5 bg-rose-50 border border-rose-200 rounded-xl flex items-start gap-2.5 text-rose-800 text-xs sm:text-sm animate-in fade-in duration-200">
                <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                <div className="flex-1 leading-relaxed">
                  <p className="font-bold mb-0.5">خطا</p>
                  <p>{error}</p>
                </div>
              </div>
            )}

            {/* Action Buttons */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
              <button
                type="submit"
                disabled={isLoading}
                className="w-full h-12 sm:h-13 px-4 bg-slate-900 hover:bg-slate-800 active:bg-slate-950 text-white font-bold text-xs sm:text-sm rounded-xl transition-all shadow-sm hover:shadow flex items-center justify-center gap-2 disabled:opacity-60 disabled:cursor-not-allowed cursor-pointer"
              >
                {isLoading ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin text-emerald-400" />
                    <span>در حال دریافت تمامی کامنت‌ها...</span>
                  </>
                ) : (
                  <>
                    <Send className="w-4 h-4 text-emerald-400" />
                    <span>ارسال و استخراج کامنت‌ها</span>
                  </>
                )}
              </button>

              <button
                type="button"
                onClick={handleCopyAllComments}
                disabled={!result || isLoading}
                className={`w-full h-12 sm:h-13 px-4 font-bold text-xs sm:text-sm rounded-xl transition-all flex items-center justify-center gap-2 shadow-sm ${
                  result
                    ? 'bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white shadow-emerald-600/25 hover:shadow-md cursor-pointer animate-in zoom-in-95 duration-200'
                    : 'bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed'
                }`}
              >
                {copiedAll ? (
                  <>
                    <Check className="w-4 h-4" />
                    <span>✓ تمام کامنت‌ها کپی شدند!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-4 h-4" />
                    <span>
                      {result
                        ? `کپی تمام کامنت‌ها (${result.totalCount} کامنت)`
                        : 'کپی تمام کامنت‌ها (در انتظار ارسال)'}
                    </span>
                  </>
                )}
              </button>
            </div>
          </form>
        </div>

        {/* Lower Section: Prominent Comments Display Box & Copy Actions */}
        {result && (
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden mb-8 animate-in fade-in slide-in-from-bottom-4 duration-300">
            {/* Box Header */}
            <div className="p-4 sm:p-5 bg-slate-900 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-emerald-500/20 text-emerald-400 flex items-center justify-center border border-emerald-500/30 shrink-0">
                  <CheckCircle2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-sm sm:text-base">
                    تمامی کامنت‌ها با موفقیت دریافت شدند
                  </h3>
                  <div className="flex items-center gap-2 text-xs text-slate-400 mt-0.5">
                    <span className="text-emerald-400 font-bold font-mono tabular-nums">
                      {result.totalCount} کامنت استخراج‌شده
                    </span>
                    <span aria-hidden="true">·</span>
                    <span>{result.pagesFetched} درخواست مصرف شد ({result.pagesFetched} صفحه)</span>
                  </div>
                </div>
              </div>

              {/* Main Copy Action Buttons */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleCopyAllComments}
                  className="inline-flex items-center justify-center gap-1.5 px-4 py-2 bg-emerald-500 hover:bg-emerald-400 active:bg-emerald-600 text-slate-950 font-bold text-xs rounded-xl transition-all shadow-sm cursor-pointer"
                >
                  {copiedAll ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedAll ? '✓ کپی شد!' : 'کپی تمام کامنت‌ها برای اکسل'}</span>
                </button>

                <button
                  type="button"
                  onClick={handleDownloadExcel}
                  className="inline-flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-semibold transition-colors cursor-pointer border border-slate-700"
                  title="دانلود فایل اکسل .xlsx"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>دانلود .xlsx</span>
                </button>
              </div>
            </div>

            {/* Instruction Banner for Excel */}
            <div className="bg-emerald-50 border-b border-emerald-100 px-4 sm:px-5 py-2.5 text-xs text-emerald-900 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center gap-1.5">
                <span className="font-bold text-emerald-800">راهنمای انتقال به اکسل:</span>
                <span>
                  روی دکمه <strong>«کپی تمام کامنت‌ها برای اکسل»</strong> کلیک کنید، سپس در اکسل <strong>Ctrl+V</strong> را بزنید؛ هر کامنت در یک سطر جداگانه قرار می‌گیرد.
                </span>
              </div>
              <button
                type="button"
                onClick={handleCopyWithRowNumbers}
                className="text-[11px] text-emerald-800 hover:text-emerald-950 font-bold underline shrink-0 cursor-pointer"
              >
                {copiedWithRows ? '✓ با شماره ردیف کپی شد' : 'کپی به همراه شماره ردیف (۲ ستونه)'}
              </button>
            </div>

            {/* Search & View Mode Switcher */}
            <div className="p-3.5 sm:p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-700">حالت نمایش:</span>
                <div className="p-0.5 bg-slate-100 rounded-lg flex items-center text-xs">
                  <button
                    type="button"
                    onClick={() => setViewMode('list')}
                    className={`px-2.5 py-1 rounded-md transition-all cursor-pointer ${
                      viewMode === 'list'
                        ? 'bg-white font-bold text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    سطر به سطر با شماره
                  </button>
                  <button
                    type="button"
                    onClick={() => setViewMode('textarea')}
                    className={`px-2.5 py-1 rounded-md transition-all cursor-pointer ${
                      viewMode === 'textarea'
                        ? 'bg-white font-bold text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    کادر متنی یکپارچه
                  </button>
                </div>
              </div>

              {/* Search in comments */}
              <div className="relative w-full sm:w-60">
                <input
                  type="text"
                  placeholder="جستجو در بین کامنت‌ها..."
                  value={filterQuery}
                  onChange={(e) => setFilterQuery(e.target.value)}
                  className="w-full h-8 pl-3 pr-8 text-xs bg-slate-50 border border-slate-200 rounded-lg focus:outline-none focus:bg-white focus:border-emerald-600"
                />
                <Search className="w-3.5 h-3.5 text-slate-400 absolute right-2.5 top-2.5" />
              </div>
            </div>

            {/* THE COMMENTS DISPLAY AREA */}
            <div className="p-3.5 sm:p-5">
              {viewMode === 'list' ? (
                <div className="border border-slate-200 rounded-xl max-h-[500px] overflow-y-auto divide-y divide-slate-100 bg-slate-50/30">
                  {filteredComments.map((comment) => (
                    <div
                      key={comment.row}
                      className="p-2.5 sm:p-3 text-xs flex items-start gap-2.5 hover:bg-slate-50 transition-colors"
                    >
                      <span className="w-7 shrink-0 text-center font-mono tabular-nums text-slate-400 font-semibold pt-0.5">
                        {comment.row}
                      </span>
                      <p className="text-slate-900 font-medium leading-relaxed flex-1 select-text">
                        {comment.text}
                      </p>
                    </div>
                  ))}
                  {filteredComments.length === 0 && (
                    <div className="py-8 text-center text-xs text-slate-400">
                      کامنتی مطابق با عبارت جستجو یافت نشد.
                    </div>
                  )}
                </div>
              ) : (
                <div>
                  <textarea
                    readOnly
                    rows={14}
                    value={filteredComments.map((c) => c.text).join('\n')}
                    className="w-full p-3 font-sans text-xs text-slate-900 bg-slate-50 border border-slate-300 rounded-xl leading-relaxed select-all focus:outline-none focus:bg-white focus:border-emerald-600"
                  />
                  <p className="mt-1 text-[11px] text-slate-500">
                    نکته: هر سطر یک کامنت مستقل است. می‌توانید تمام متن را انتخاب (Ctrl+A) کرده و مستقیماً در اکسل پیست کنید.
                  </p>
                </div>
              )}

              {/* Bottom Quick Bar */}
              <div className="mt-3 flex flex-col sm:flex-row items-center justify-between gap-2 text-xs text-slate-600 pt-2 border-t border-slate-100">
                <span>نمایش {filteredComments.length} از {result.totalCount} کامنت</span>
                <button
                  type="button"
                  onClick={handleCopyAllComments}
                  className="text-emerald-700 hover:text-emerald-800 font-bold flex items-center gap-1.5 transition-colors cursor-pointer"
                >
                  {copiedAll ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                  <span>{copiedAll ? 'کپی شد!' : 'کپی مجدد تمام کامنت‌ها'}</span>
                </button>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
