const BASE = '/api';

function getPersonId() {
  return localStorage.getItem('cabin_person_id') || '';
}

async function req(method, url, body) {
  const headers = { 'Content-Type': 'application/json' };
  const pid = getPersonId();
  if (pid) headers['x-person-id'] = pid;
  const res = await fetch(BASE + url, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: res.statusText }));
    throw new Error(err.error || 'Request failed');
  }
  return res.json();
}

export const api = {
  get: (url) => req('GET', url),
  post: (url, body) => req('POST', url, body),
  patch: (url, body) => req('PATCH', url, body),
  put: (url, body) => req('PUT', url, body),
  del: (url) => req('DELETE', url),
};
