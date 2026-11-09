(() => {
  const VERSION = 4;
  if (window.__m365cli && window.__m365cli.version === VERSION) return;

  const bearers = new Map();
  const skypes = new Map();
  const people = new Map();
  const originalFetch = window.fetch.bind(window);

  const claimsOf = (bearer) => {
    try {
      const payload = bearer.split(' ')[1].split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      const bytes = Uint8Array.from(atob(payload), (c) => c.charCodeAt(0));
      return JSON.parse(new TextDecoder().decode(bytes));
    } catch {
      return null;
    }
  };

  const keepBearer = (value) => {
    if (!value || !/^Bearer /i.test(value)) return;
    const claims = claimsOf(value);
    if (!claims || !claims.aud || !claims.tid) return;
    const key = claims.tid + '|' + claims.aud;
    const exp = Number(claims.exp || 0) * 1000;
    const old = bearers.get(key);
    if (!old || old.exp <= exp) bearers.set(key, { value, exp });
    if (claims.oid && !people.has(claims.oid)) {
      people.set(claims.oid, {
        oid: claims.oid,
        tid: claims.tid,
        name: claims.name,
        upn: claims.upn || claims.unique_name || claims.preferred_username,
      });
    }
  };

  const headerOf = (headers, name) => {
    if (!headers) return null;
    try {
      return new Headers(headers).get(name);
    } catch {
      return null;
    }
  };

  const watchResponse = (url, response, authorization) => {
    if (/\/authsvc\/v1\.0\/authz/.test(url)) {
      const claims = authorization ? claimsOf(authorization) : null;
      response.clone().json().then((j) => {
        if (claims && claims.tid && j && j.tokens && j.tokens.skypeToken) {
          skypes.set(claims.tid, {
            value: j.tokens.skypeToken,
            exp: Date.now() + Number(j.tokens.expiresIn || 3600) * 1000,
          });
        }
      }).catch(() => {});
    } else if (/\/oauth2\/v2\.0\/token/.test(url)) {
      response.clone().json().then((j) => {
        if (j && j.access_token) keepBearer('Bearer ' + j.access_token);
      }).catch(() => {});
    }
  };

  window.fetch = async function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || String(input);
    const headers = (init && init.headers) || (input instanceof Request ? input.headers : null);
    const authorization = headerOf(headers, 'authorization');
    keepBearer(authorization);
    const response = await originalFetch(input, init);
    watchResponse(url, response, authorization);
    return response;
  };

  const xhrHeader = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (String(name).toLowerCase() === 'authorization') keepBearer(value);
    return xhrHeader.apply(this, arguments);
  };

  const stored = (suffix) => {
    const key = Object.keys(localStorage).find((k) => k.endsWith(suffix));
    if (!key) return null;
    try {
      const parsed = JSON.parse(localStorage.getItem(key));
      return parsed && parsed.item !== undefined ? parsed.item : parsed;
    } catch {
      return null;
    }
  };

  const primaryId = () => {
    const value = stored('.PrimaryUserId.PrimaryUserId');
    return typeof value === 'string' ? value : null;
  };

  const me = () => {
    const id = primaryId();
    return id ? people.get(id) || null : null;
  };

  const serviceMap = () => {
    const id = primaryId();
    if (!id) return {};
    const value = stored('.' + id + '.Discover.DISCOVER-REGION-GTM');
    return value && typeof value === 'object' ? value : {};
  };

  const FIXED = {
    substrate: 'https://substrate.office.com',
    presence: 'https://presence.teams.microsoft.com',
  };

  const baseOf = (service) => {
    if (FIXED[service]) return FIXED[service];
    const value = serviceMap()[service];
    return typeof value === 'string' && /^https:\/\//.test(value) ? value.replace(/\/$/, '') : null;
  };

  const status = () => {
    const now = Date.now();
    const user = me();
    const tid = user && user.tid;
    const own = {};
    for (const [key, t] of bearers) {
      const [tenant, aud] = key.split('|');
      if (tenant === tid) own[aud] = Math.round((t.exp - now) / 1000);
    }
    const skype = tid ? skypes.get(tid) : null;
    return {
      version: VERSION,
      skype: skype ? Math.round((skype.exp - now) / 1000) : null,
      bearers: own,
      accounts: skypes.size,
      chatService: baseOf('chatServiceAfd'),
      me: user,
      url: location.origin,
    };
  };

  const call = async (req) => {
    const user = me();
    if (!user) return { status: 0, error: 'no-token account' };
    const base = baseOf(req.service);
    if (!base) return { status: 0, error: 'unknown service ' + req.service };
    const headers = Object.assign({ Accept: 'application/json' }, req.headers || {});
    for (const auth of req.auth || []) {
      if (auth === 'skype') {
        const skype = skypes.get(user.tid);
        if (!skype) return { status: 0, error: 'no-skype-token' };
        headers['Authentication'] = 'skypetoken=' + skype.value;
        headers['x-skypetoken'] = skype.value;
      } else {
        const token = bearers.get(user.tid + '|' + auth);
        if (!token) return { status: 0, error: 'no-token ' + auth };
        headers['Authorization'] = token.value;
      }
    }
    const init = { method: req.method || 'GET', headers, credentials: 'omit' };
    if (req.body !== undefined && req.body !== null) {
      init.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
      if (!headers['Content-Type']) headers['Content-Type'] = 'application/json';
    }
    try {
      const response = await originalFetch(base + req.path, init);
      return {
        status: response.status,
        retryAfter: response.headers.get('Retry-After'),
        body: await response.text(),
      };
    } catch (e) {
      return { status: 0, error: String((e && e.message) || e) };
    }
  };

  Object.defineProperty(window, '__m365cli', {
    value: Object.freeze({ version: VERSION, status, call }),
    configurable: true,
  });
})();
