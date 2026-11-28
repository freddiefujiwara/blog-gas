import { describe, it, expect, vi, beforeEach } from 'vitest';
import * as Code from '../src/Code.js';

// Authentication gate tests. The gate is active only when the
// AVAILABLE_GMAILS script property is set (the staging project);
// production leaves it unset and every read stays public, which is
// what the rest of the suite exercises.

let store;
const mockProperties = {
  getProperty: vi.fn(),
  setProperty: vi.fn(),
  deleteProperty: vi.fn(),
};

const mockUrlFetchApp = {
  fetch: vi.fn(),
  fetchAll: vi.fn(),
};

let lastOutput;
const mockContentService = {
  MimeType: { JSON: 'JSON', XML: 'XML' },
  createTextOutput: vi.fn(),
};

global.PropertiesService = {
  getScriptProperties: vi.fn().mockReturnValue(mockProperties),
};
global.UrlFetchApp = mockUrlFetchApp;
global.ContentService = mockContentService;
global.CacheService = { getScriptCache: vi.fn().mockReturnValue({ get: vi.fn(), put: vi.fn() }) };
global.Utilities = {
  computeDigest: vi.fn().mockReturnValue([0, 15, 255, -1]),
  DigestAlgorithm: { MD5: 'MD5' },
};
global.console = { log: vi.fn(), error: vi.fn() };

const tokenInfoOk = (overrides = {}) => ({
  getResponseCode: () => 200,
  getContentText: () => JSON.stringify({
    iss: 'https://accounts.google.com',
    aud: 'test-client-id',
    exp: Math.floor(Date.now() / 1000) + 3600,
    email_verified: 'true',
    email: 'zegenvs@gmail.com',
    ...overrides,
  }),
});

describe('authentication gate', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    store = {
      AVAILABLE_GMAILS: 'zegenvs@gmail.com, fumikazu.fujiwara@rakuten.com',
      GOOGLE_OAUTH_CLIENT_ID: 'test-client-id',
    };
    mockProperties.getProperty.mockImplementation((key) => store[key] ?? null);
    mockProperties.setProperty.mockImplementation((key, value) => { store[key] = value; });
    mockProperties.deleteProperty.mockImplementation((key) => { delete store[key]; });
    mockContentService.createTextOutput.mockImplementation((s) => {
      lastOutput = s;
      return { setMimeType: vi.fn().mockReturnThis() };
    });
    global.Utilities.computeDigest.mockReturnValue([0, 15, 255, -1]);
  });

  it('is off when AVAILABLE_GMAILS is unset (production behaviour)', () => {
    delete store.AVAILABLE_GMAILS;
    expect(Code.authRequired_()).toBe(false);
    expect(Code.isAuthorizedRead_({})).toBe(true);
  });

  it('rejects reads with no token when the gate is on', () => {
    expect(Code.authRequired_()).toBe(true);
    expect(Code.isAuthorizedRead_({})).toBe(false);
    expect(Code.isAuthorizedRead_({ session_token: 'nope' })).toBe(false);
  });

  it('verifies a Google ID token for an allowed email', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk());
    expect(Code.verifyGoogleIdToken_('tok')).toBe('zegenvs@gmail.com');
    expect(Code.isAuthorizedRead_({ id_token: 'tok' })).toBe(true);
  });

  it('rejects a token for an email outside the allowlist', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk({ email: 'someone@example.com' }));
    expect(() => Code.verifyGoogleIdToken_('tok')).toThrow();
    expect(Code.isAuthorizedRead_({ id_token: 'tok' })).toBe(false);
  });

  it('rejects a token with the wrong audience, an expired token, and a failed lookup', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk({ aud: 'other-client' }));
    expect(() => Code.verifyGoogleIdToken_('tok')).toThrow();
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk({ exp: 1 }));
    expect(() => Code.verifyGoogleIdToken_('tok')).toThrow();
    mockUrlFetchApp.fetch.mockReturnValue({ getResponseCode: () => 400, getContentText: () => '{}' });
    expect(() => Code.verifyGoogleIdToken_('tok')).toThrow();
  });

  it('exchanges an ID token for a session that then authorizes reads', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk());
    const res = Code.exchangeSessionResponse_({ id_token: 'tok' });
    expect(res.error).toBeUndefined();
    expect(typeof res.session_token).toBe('string');
    expect(res.email).toBe('zegenvs@gmail.com');
    expect(Code.validateSession_(res.session_token)).toBe('zegenvs@gmail.com');
    expect(Code.isAuthorizedRead_({ session_token: res.session_token })).toBe(true);
  });

  it('exchange fails closed for a bad ID token', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk({ email: 'someone@example.com' }));
    expect(Code.exchangeSessionResponse_({ id_token: 'tok' })).toEqual({ error: 'unauthorized' });
    expect(Code.exchangeSessionResponse_({})).toEqual({ error: 'unauthorized' });
  });

  it('expired sessions and revoked sessions no longer authorize', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk());
    const { session_token: token } = Code.exchangeSessionResponse_({ id_token: 'tok' });
    const sessions = JSON.parse(store[Code.AUTH_SESSIONS_PROPERTY]);
    sessions[token].exp = Date.now() - 1;
    store[Code.AUTH_SESSIONS_PROPERTY] = JSON.stringify(sessions);
    expect(Code.validateSession_(token)).toBeNull();

    const { session_token: token2 } = Code.exchangeSessionResponse_({ id_token: 'tok' });
    Code.revokeSession_(token2);
    expect(Code.validateSession_(token2)).toBeNull();
  });

  it('authorizes a read carrying the share token when SHARE_TOKEN is set', () => {
    store.SHARE_TOKEN = 'a-long-random-share-token';
    expect(Code.validateShareToken_('a-long-random-share-token')).toBe(true);
    expect(Code.isAuthorizedRead_({ share_token: 'a-long-random-share-token' })).toBe(true);
  });

  it('rejects a wrong or empty share token', () => {
    store.SHARE_TOKEN = 'a-long-random-share-token';
    expect(Code.validateShareToken_('some-other-token')).toBe(false);
    expect(Code.validateShareToken_('')).toBe(false);
    expect(Code.isAuthorizedRead_({ share_token: 'some-other-token' })).toBe(false);
  });

  it('grants no share access when SHARE_TOKEN is unset', () => {
    expect(Code.validateShareToken_('a-long-random-share-token')).toBe(false);
    expect(Code.isAuthorizedRead_({ share_token: 'a-long-random-share-token' })).toBe(false);
  });

  it('doGet answers unauthorized before any read when the gate is on', () => {
    Code.doGet({ parameter: { o: 'titles' } });
    expect(JSON.parse(lastOutput)).toEqual({ error: 'unauthorized' });
    Code.doGet({ parameter: {} });
    expect(JSON.parse(lastOutput)).toEqual({ error: 'unauthorized' });
  });

  it('doGet routes exchange_session before the gate', () => {
    mockUrlFetchApp.fetch.mockReturnValue(tokenInfoOk());
    Code.doGet({ parameter: { action: 'exchange_session', id_token: 'tok' } });
    const parsed = JSON.parse(lastOutput);
    expect(parsed.error).toBeUndefined();
    expect(typeof parsed.session_token).toBe('string');
  });
});
