'use strict';

const { spFetch } = require('./sharepoint-fetch');
const { AUTH_FILE, readAuthFile } = require('./sharepoint-auth');

function normalizeEndpoint(endpoint) {
  return endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
}

function parseResponseBody(text) {
  const trimmed = text.trim();
  if (!trimmed) return null;
  try {
    return JSON.parse(trimmed);
  } catch {
    return trimmed;
  }
}

function loadAuth(authFile = AUTH_FILE) {
  const auth = readAuthFile(authFile);
  if (!auth?.SP_SITE) {
    throw new Error('No SharePoint site is authenticated. Run sp-api auth login --site <site>.');
  }
  if (!auth.SP_TOKEN && !auth.SP_COOKIES) {
    throw new Error('No SharePoint credentials are available. Run sp-api auth login --site <site>.');
  }
  return auth;
}

function authHeaders(auth) {
  if (auth.SP_TOKEN) return { Authorization: `Bearer ${auth.SP_TOKEN}` };
  return { Cookie: auth.SP_COOKIES };
}

async function fetchDigest(auth, deps = {}) {
  const res = await spFetch(`${auth.SP_SITE}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      ...authHeaders(auth),
      Accept: 'application/json;odata=nometadata',
      'Content-Length': '0',
    },
  }, deps);
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`Failed to fetch request digest. HTTP ${res.status}: ${text}`);
  }
  const json = parseResponseBody(text);
  if (!json?.FormDigestValue) {
    throw new Error('Could not parse request digest from contextinfo response.');
  }
  return json.FormDigestValue;
}

async function executeSharePointRequest(spec, endpoint, body, deps = {}) {
  const auth = deps.auth || loadAuth(deps.authFile);
  const url = `${auth.SP_SITE}${normalizeEndpoint(endpoint)}`;
  const headers = {
    ...authHeaders(auth),
    Accept: spec.accept || 'application/json;odata=nometadata',
  };
  const method = spec.method || 'GET';
  const fetchOptions = { method, headers };

  if (method !== 'GET') {
    const digest = await fetchDigest(auth, deps);
    fetchOptions.method = 'POST';
    fetchOptions.headers = {
      ...headers,
      'Content-Type': spec.contentType || 'application/json;odata=verbose',
      'X-RequestDigest': digest,
    };
    if (method !== 'POST') {
      fetchOptions.headers['X-HTTP-Method'] = method;
      if (['MERGE', 'PATCH', 'DELETE'].includes(method.toUpperCase())) {
        fetchOptions.headers['If-Match'] = '*';
      }
    }
    if (body) fetchOptions.body = body;
  }

  const res = await spFetch(url, fetchOptions, deps);
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} on ${method} ${url}: ${text}`);
  }
  return { data: parseResponseBody(text), endpoint: normalizeEndpoint(endpoint), method };
}

module.exports = {
  authHeaders,
  executeSharePointRequest,
  fetchDigest,
  loadAuth,
  normalizeEndpoint,
  parseResponseBody,
};
