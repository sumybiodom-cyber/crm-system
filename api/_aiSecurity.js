const crypto = require('node:crypto');

const aiRateLimits = new Map();
const aiRequestAuditLog = [];
const MAX_AUDIT_RECORDS = 200;

const ACCESS_DENIED_MESSAGE = 'У вас нет доступа к этому действию. Обратитесь к администратору CRM.';
const UNAVAILABLE_MESSAGE = 'AI временно недоступен. Попробуйте позже.';

const ALLOWED_ROLES = new Set([
  'Директор',
  'Руководитель',
  'Адміністратор',
  'Менеджер',
  'Інженер',
  'Склад',
  'Комірник',
  'Бухгалтер',
  'Закупник',
]);

const BLOCKED_PATTERNS = [
  /system prompt/i,
  /hidden instructions?/i,
  /reveal.*prompt/i,
  /ignore (all|previous|earlier) instructions?/i,
  /developer message/i,
  /internal routes?/i,
  /api keys?/i,
  /secret/i,
  /backend architecture/i,
  /printenv/i,
  /process\.env/i,
  /\bselect\b.+\bfrom\b/i,
  /\bdrop\b.+\btable\b/i,
  /\binsert\b.+\binto\b/i,
  /\bupdate\b.+\bset\b/i,
  /\bdelete\b.+\bfrom\b/i,
  /\bexec\b/i,
  /\beval\b/i,
  /\bpowershell\b/i,
  /\bcmd\.exe\b/i,
  /\bbash\b/i,
  /\bshell\b.+\b(command|script)\b/i,
];

function normalizeText(value, maxLength = 4000) {
  return String(value ?? '')
    .replace(/\0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLength);
}

function nowIso() {
  return new Date().toISOString();
}

function getClientAddress(req) {
  const forwarded = req.headers['x-forwarded-for'];
  if (typeof forwarded === 'string' && forwarded.trim()) {
    return forwarded.split(',')[0].trim();
  }
  return req.socket?.remoteAddress || 'unknown';
}

function buildAuditEntry({ req, body, outcome, reason }) {
  const session = body?.context?.session && typeof body.context.session === 'object' ? body.context.session : {};
  return {
    id: crypto.randomUUID(),
    at: nowIso(),
    ip: getClientAddress(req),
    origin: req.headers.origin || req.headers.referer || '',
    userId: normalizeText(session.activeUserId || session.sessionUserId, 120),
    role: normalizeText(body?.context?.roleKey, 80),
    section: normalizeText(body?.context?.page, 80),
    outcome,
    reason: normalizeText(reason, 280),
  };
}

function logAiRequest(entry) {
  aiRequestAuditLog.unshift(entry);
  if (aiRequestAuditLog.length > MAX_AUDIT_RECORDS) {
    aiRequestAuditLog.length = MAX_AUDIT_RECORDS;
  }
}

function getAllowedOrigin() {
  return normalizeText(process.env.CRM_ALLOWED_ORIGIN, 300);
}

function validateOrigin(req) {
  const allowedOrigin = getAllowedOrigin();
  const origin = normalizeText(req.headers.origin || req.headers.referer, 300);
  if (!origin) {
    return { ok: false, reason: 'Missing origin' };
  }
  if (!allowedOrigin) {
    return { ok: true };
  }
  return origin.startsWith(allowedOrigin)
    ? { ok: true }
    : { ok: false, reason: 'Origin not allowed' };
}

function validateSessionContext(context) {
  const session = context?.session && typeof context.session === 'object' ? context.session : null;
  if (!session) {
    return { ok: false, reason: 'Missing session context' };
  }
  if (!normalizeText(session.sessionUserId, 120) || !normalizeText(session.activeUserId, 120)) {
    return { ok: false, reason: 'Missing session identifiers' };
  }
  if (session.isSessionLocked) {
    return { ok: false, reason: 'Session locked' };
  }
  if (normalizeText(session.userSessionStatus, 80) !== 'Активна') {
    return { ok: false, reason: 'Inactive session' };
  }
  const roleKey = normalizeText(context?.roleKey, 80);
  if (!ALLOWED_ROLES.has(roleKey)) {
    return { ok: false, reason: 'Role not allowed' };
  }
  return { ok: true };
}

function createRateKey(req, context) {
  const session = context?.session && typeof context.session === 'object' ? context.session : {};
  return `${getClientAddress(req)}:${normalizeText(session.activeUserId || 'guest', 120)}`;
}

function enforceRateLimit(req, context) {
  const limit = Math.max(Number(process.env.AI_RATE_LIMIT_PER_MINUTE || 18), 1);
  const windowMs = 60_000;
  const key = createRateKey(req, context);
  const now = Date.now();
  const current = aiRateLimits.get(key) || [];
  const recent = current.filter((ts) => now - ts < windowMs);
  if (recent.length >= limit) {
    aiRateLimits.set(key, recent);
    return { ok: false, reason: 'Rate limit exceeded' };
  }
  recent.push(now);
  aiRateLimits.set(key, recent);
  return { ok: true };
}

function sanitizeHistory(history) {
  if (!Array.isArray(history)) return [];
  return history.slice(-10).map((entry) => ({
    author: entry?.author === 'assistant' ? 'assistant' : 'user',
    text: normalizeText(entry?.text, 1200),
  }));
}

function sanitizeContext(context) {
  const safeContext = context && typeof context === 'object' ? context : {};
  const allowedPages = Array.isArray(safeContext.allowedPages)
    ? safeContext.allowedPages.map((page) => normalizeText(page, 80)).filter(Boolean).slice(0, 20)
    : [];
  const activeObject = safeContext.activeObject && typeof safeContext.activeObject === 'object'
    ? {
        type: normalizeText(safeContext.activeObject.type, 40),
        id: normalizeText(safeContext.activeObject.id, 80),
        title: normalizeText(safeContext.activeObject.title, 160),
        subtitle: normalizeText(safeContext.activeObject.subtitle, 200),
      }
    : null;
  return {
    role: normalizeText(safeContext.role, 80),
    roleKey: normalizeText(safeContext.roleKey, 80),
    section: normalizeText(safeContext.section, 120),
    page: normalizeText(safeContext.page, 80),
    allowedPages,
    activeObject,
    session: safeContext.session && typeof safeContext.session === 'object'
      ? {
          sessionUserId: normalizeText(safeContext.session.sessionUserId, 120),
          activeUserId: normalizeText(safeContext.session.activeUserId, 120),
          userSessionStatus: normalizeText(safeContext.session.userSessionStatus, 80),
          isSessionLocked: Boolean(safeContext.session.isSessionLocked),
        }
      : null,
    permissions: safeContext.permissions && typeof safeContext.permissions === 'object'
      ? {
          finance: Boolean(safeContext.permissions.finance),
          payments: Boolean(safeContext.permissions.payments),
          warehouse: Boolean(safeContext.permissions.warehouse),
          purchases: Boolean(safeContext.permissions.purchases),
          documents: Boolean(safeContext.permissions.documents),
          payroll: Boolean(safeContext.permissions.payroll),
          settings: Boolean(safeContext.permissions.settings),
          profit: Boolean(safeContext.permissions.profit),
          salary: Boolean(safeContext.permissions.salary),
          cost: Boolean(safeContext.permissions.cost),
        }
      : {},
  };
}

function detectPromptAttack(message, history) {
  const combined = [normalizeText(message, 3000), ...sanitizeHistory(history).map((entry) => entry.text)].join('\n');
  const matched = BLOCKED_PATTERNS.find((pattern) => pattern.test(combined));
  return matched ? { blocked: true, reason: 'Prompt injection or command-like request detected' } : { blocked: false };
}

function buildBlockedSecurityReply(kind = 'default') {
  if (kind === 'access') return ACCESS_DENIED_MESSAGE;
  return 'Этот запрос заблокирован политикой безопасности CRM.';
}

module.exports = {
  ACCESS_DENIED_MESSAGE,
  UNAVAILABLE_MESSAGE,
  buildAuditEntry,
  buildBlockedSecurityReply,
  detectPromptAttack,
  enforceRateLimit,
  logAiRequest,
  normalizeText,
  sanitizeContext,
  sanitizeHistory,
  validateOrigin,
  validateSessionContext,
};
