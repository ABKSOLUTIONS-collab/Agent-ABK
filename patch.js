const fs = require('fs');
let html = fs.readFileSync('/app/client/dist/index.html', 'utf8');

// Remove any previous ABK branding injection (idempotent)
html = html.replace(/<!-- ABK_START -->[\s\S]*?<!-- ABK_END -->\n?/g, '');
// Remove old-style (without markers) - cleanup legacy
html = html.replace(/<style>\s*footer \{ display: none !important[\s\S]*?<\/script>\n?/g, '');

// Set title
html = html.replace(/<title>[^<]*<\/title>/, '<title>ABK Assistant</title>');

// Inject new branding with markers
const inject = `<!-- ABK_START -->
<style>
  /* Hide LibreChat footer and links */
  footer { display: none !important; }
  a[href*="librechat.ai"] { display: none !important; }

  /* ── Hide Sign Up / registration links via CSS (reliable, no JS race) ── */
  a[href*="register"] { display: none !important; }
  p:has(a[href*="register"]),
  div:has(> a[href*="register"]),
  span:has(a[href*="register"]) { display: none !important; }

  /* ── ABK Solutions color overrides ── */

  /* Submit button: green → ABK blue */
  html, .light, .dark {
    --surface-submit: #0066CC !important;
    --surface-submit-hover: #0055B3 !important;
  }

  /* Input focus ring/border: green → ABK blue */
  input:focus,
  input:focus-visible {
    border-color: #0066CC !important;
    --tw-ring-color: rgba(0, 102, 204, 0.35) !important;
    box-shadow: 0 0 0 2px rgba(0, 102, 204, 0.25) !important;
    outline: none !important;
  }

  /* "Sign up" and other auth links: teal → ABK blue */
  a[href*="register"],
  a[href*="signup"],
  .login-form a,
  form ~ div a,
  form a {
    color: #0066CC !important;
  }

  /* Hide raven SVG logo (React-safe: hide in CSS, don't remove from DOM) */
  svg[data-abk-hidden="1"] {
    display: none !important;
  }

  /* ── ABK blue: ONLY for main app (not login/register) ── */
  /* body.abk-app is added by JS only on non-auth pages */
  body.abk-app { color: #0066CC !important; }
  body.abk-app h1, body.abk-app h2,
  body.abk-app h3, body.abk-app h4 { color: #0066CC !important; }
  /* Keep muted / secondary text gray */
  body.abk-app [class*="text-gray-4"],
  body.abk-app [class*="text-gray-5"],
  body.abk-app [class*="text-gray-6"],
  body.abk-app [class*="text-muted"],
  body.abk-app footer, body.abk-app footer * { color: #9ca3af !important; }
  /* Keep white text on dark backgrounds */
  body.abk-app [class*="bg-gray-9"] *,
  body.abk-app [class*="bg-black"] *,
  body.abk-app [style*="background-color: #0066"] *,
  body.abk-app [style*="background:#0066"] * { color: #fff !important; }
  /* Input typing text: keep dark */
  body.abk-app input, body.abk-app textarea { color: #111827 !important; }
  body.abk-app input::placeholder,
  body.abk-app textarea::placeholder { color: #9ca3af !important; }
  /* Green accents → ABK blue */
  body.abk-app [class*="text-green-"] { color: #0066CC !important; }
  body.abk-app [class*="bg-green-"] { background-color: rgba(0,102,204,0.12) !important; }
  body.abk-app [class*="border-green-"] { border-color: #0066CC !important; }
  /* Send button */
  body.abk-app button[class*="bg-black"],
  body.abk-app button[class*="bg-gray-900"] { background-color: #0066CC !important; }
  /* Greeting icon hidden by JS */
  [data-abk-greet-icon-hidden="1"] { display: none !important; }
</style>
<script>
// Unregister LibreChat's service worker so the latest index.html is always served.
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.getRegistrations().then(function(regs) {
    regs.forEach(function(r) { r.unregister(); });
    if (regs.length > 0) { console.log('[ABK] SW unregistered, reloading...'); window.location.reload(); }
  });
}
</script>
<script>
(function() {
  'use strict';
  // 0. Add/remove body.abk-app class — main app only, NOT on auth pages
  function updateAppClass() {
    if (!document.body) return;
    var path = window.location.pathname;
    var isAuth = (path === '/' || path.indexOf('/login') !== -1 || path.indexOf('/register') !== -1);
    if (isAuth) {
      document.body.classList.remove('abk-app');
      // Landing back on the login page only happens right after a logout.
      // The injected <script> in index.html runs once per HARD page load —
      // if LibreChat's logout/login is a client-side SPA transition (no full
      // reload), a stale detected email from the PREVIOUS user in the same
      // tab would otherwise leak into the next login (e.g. "Signed in as"
      // in Organization Settings showing the wrong person). Clear it here.
      if (typeof abkLastAppliedEmail !== 'undefined') { abkLastAppliedEmail = null; }
      localStorage.removeItem('abk_admin_email');
    } else {
      document.body.classList.add('abk-app');
    }
  }
  window.addEventListener('popstate', updateAppClass);
  // Intercept pushState/replaceState for SPA navigation
  try {
    var _origPush = history.pushState.bind(history);
    var _origReplace = history.replaceState.bind(history);
    history.pushState = function() { _origPush.apply(history, arguments); updateAppClass(); };
    history.replaceState = function() { _origReplace.apply(history, arguments); updateAppClass(); };
  } catch(e) {}

  // 1. Fix title
  var t = function() {
    if (document.title !== 'ABK Assistant') {
      document.title = 'ABK Assistant';
    }
  };
  t();
  new MutationObserver(t).observe(document.head, { childList: true, subtree: true });

  // 2. Patch login heading + subtitle
  function patchLoginText() {
    document.querySelectorAll('h1, h2, h3').forEach(function(el) {
      if (el.textContent.trim() === 'Welcome back' && !el.dataset.abkDone) {
        el.dataset.abkDone = '1';
        el.textContent = 'Welcome Back';

        // Hide LibreChat native appTitle shown above (avoid duplicate)
        var prev = el.previousElementSibling;
        while (prev) {
          if (prev.textContent.trim() === 'ABK Assistant') {
            prev.style.display = 'none';
          }
          prev = prev.previousElementSibling;
        }

        // Remove any previous subtitle we added
        var next = el.nextElementSibling;
        if (next && next.classList.contains('abk-subtitle')) {
          next.remove();
        }

        // Add subtitle
        var sub = document.createElement('p');
        sub.className = 'abk-subtitle';
        sub.textContent = 'ABK Assistant';
        sub.style.cssText = 'margin:6px 0 0;font-size:0.9rem;color:#6b7280;font-weight:400;text-align:center;letter-spacing:0.04em;';
        el.parentNode.insertBefore(sub, el.nextSibling);
      }
    });
  }

  // 3. Replace raven logo with ABK Solutions logo.
  //    The login logo is <img src="assets/logo.svg" class="h-full w-full object-contain">.
  //    We target it by its src (containing "logo.svg") OR by its Tailwind classes,
  //    and forcibly redirect the src to our ABK logo — bypassing any service-worker cache.
  function patchLogo() {
    var path = window.location.pathname;
    var onAuthPage = path === '/' || path.indexOf('login') !== -1 || path.indexOf('register') !== -1;
    if (!onAuthPage) return;

    // Primary: find the logo <img> by its src containing "logo.svg"
    var imgs = document.querySelectorAll('img');
    for (var i = 0; i < imgs.length; i++) {
      var el = imgs[i];
      var src = el.getAttribute('src') || '';
      if (src.indexOf('logo.svg') !== -1) {
        // Replace src — unique query string forces SW cache bypass
        el.setAttribute('src', '/assets/abk-logo.png?abk=1');
        el.style.height = '52px';
        el.style.width = 'auto';
        el.style.display = 'block';
        el.style.margin = '0 auto 4px';
        el.style.objectFit = 'contain';
        return;
      }
    }

    // Fallback: if JS bundle was already patched (src = "assets/abk-logo.png"),
    // check if the image is NOT loading (naturalWidth === 0) and force-reload it.
    var abkImgs = document.querySelectorAll('img[src*="abk-logo"]');
    for (var j = 0; j < abkImgs.length; j++) {
      var abkEl = abkImgs[j];
      if (!abkEl.dataset.abkChecked) {
        abkEl.dataset.abkChecked = '1';
        abkEl.setAttribute('src', '/assets/abk-logo.png?abk=1');
      }
    }
  }

  // 4. Intercept "Admin Settings" sidebar link → open org-admin as modal overlay
  function showOrgAdminModal(authParam) {
    var existing = document.getElementById('abk-org-overlay');
    if (existing) { existing.remove(); return; }

    var url = 'https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io/org-admin?embed=1' +
              (authParam ? '&' + authParam : '');

    var overlay = document.createElement('div');
    overlay.id = 'abk-org-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;';

    var modal = document.createElement('div');
    modal.style.cssText = 'position:relative;width:92%;max-width:1060px;height:88vh;background:#f9f9f9;border-radius:14px;overflow:hidden;box-shadow:0 24px 80px rgba(0,0,0,.35);display:flex;flex-direction:column;';

    var header = document.createElement('div');
    header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:14px 20px;background:#fff;border-bottom:1px solid #e8e8e8;flex-shrink:0;';
    var title = document.createElement('span');
    title.style.cssText = 'font-weight:600;font-size:14px;color:#0066cc;';
    title.textContent = 'Organization Settings';
    var closeBtn = document.createElement('button');
    closeBtn.innerHTML = '&#215;';
    closeBtn.style.cssText = 'background:none;border:none;font-size:22px;cursor:pointer;color:#888;padding:0 4px;line-height:1;';
    closeBtn.onclick = function() { overlay.remove(); };
    header.appendChild(title);
    header.appendChild(closeBtn);

    var iframe = document.createElement('iframe');
    iframe.src = url;
    iframe.style.cssText = 'flex:1;width:100%;border:none;';

    modal.appendChild(header);
    modal.appendChild(iframe);
    overlay.appendChild(modal);
    overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
  }

  function extractEmail(obj) {
    if (!obj || typeof obj !== 'object') return '';
    // Try every common nested path LibreChat might use
    return obj.email || (obj.user && obj.user.email) ||
           (obj.data && obj.data.email) || (obj.result && obj.result.email) || '';
  }

  var BRIDGE_URL = 'https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io';
  var abkLastAppliedEmail = null;
  localStorage.removeItem('abk_org_role');

  // Receive messages back from the org-admin iframe — it posts back the
  // email it actually resolved server-side, which corrects our cache once
  // the panel has successfully opened for the current user.
  window.addEventListener('message', function(evt) {
    if (!evt || !evt.data) return;
    if (evt.data.type === 'abk_admin_email' && evt.data.email) {
      abkLastAppliedEmail = evt.data.email;
      localStorage.setItem('abk_admin_email', evt.data.email);
    }
  }, false);

  // LibreChat keeps its access token in memory (not localStorage, not a
  // readable cookie) and attaches it as an Authorization header via its own
  // axios/fetch wrapper — confirmed live: a plain fetch('/api/user') from
  // this injected script gets 401 even with credentials:'include', while
  // every other guessed endpoint is a flat 404. So we cannot make our own
  // authenticated request; instead we snoop on the RESPONSE BODY of
  // LibreChat's own already-authenticated '/api/auth' or '/api/user' calls
  // to learn who is actually signed in, and keep localStorage's
  // abk_admin_email fresh so openOrgAdmin() below opens the panel as the
  // CURRENT user instead of a stale cached one from a previous login in
  // the same tab.
  (function autoDetectSignedInEmail() {
    function checkData(data) {
      if (!data || typeof data !== 'object') return;
      var email = (data.user && data.user.email) || data.email ||
                  (data.data && data.data.email);
      if (email && email.indexOf('@') !== -1 && email !== abkLastAppliedEmail) {
        abkLastAppliedEmail = email;
        localStorage.setItem('abk_admin_email', email);
      }
    }

    // Wrap window.fetch to intercept LibreChat's own API responses
    var _origFetch = window.fetch;
    window.fetch = function(resource, init) {
      var url = typeof resource === 'string' ? resource
              : (resource && typeof resource.url === 'string' ? resource.url : '');
      var p = _origFetch.apply(this, arguments);
      if (url && (url.indexOf('/api/auth') !== -1 || url.indexOf('/api/user') !== -1)) {
        p.then(function(resp) {
          if (resp && resp.ok) {
            resp.clone().json().then(checkData).catch(function() {});
          }
        }).catch(function() {});
      }
      return p;
    };

    // Also attempt directly — LibreChat may have already called refresh before our script ran
    _origFetch('/api/auth/refresh', { method: 'POST', credentials: 'include' })
      .then(function(r) { return r.ok ? r.json() : null; })
      .then(function(d) { if (d) checkData(d); }).catch(function() {});
    _origFetch('/api/auth/refresh', { credentials: 'include' })
      .then(function(r) { return r.ok ? r.json() : null; })
      .then(function(d) { if (d) checkData(d); }).catch(function() {});
  })();

  // 5. "Forgot password?" link → custom reset modal (no email required)
  function showResetPasswordModal() {
    if (document.getElementById('abk-reset-overlay')) return;
    var overlay = document.createElement('div');
    overlay.id = 'abk-reset-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;';

    var card = document.createElement('div');
    card.style.cssText = 'background:#fff;border-radius:14px;padding:32px 36px;width:340px;box-shadow:0 24px 80px rgba(0,0,0,.35);';
    card.innerHTML =
      '<h2 style="font-size:16px;font-weight:600;margin:0 0 6px;color:#111">Forgot Password</h2>' +
      '<p style="font-size:13px;color:#6b7280;margin:0 0 18px">Enter your email and we\\'ll send you a reset link.</p>' +
      '<div id="abk-reset-msg" style="display:none;margin-bottom:12px;font-size:13px;padding:10px 12px;border-radius:8px;"></div>' +
      '<input id="abk-reset-email" type="email" placeholder="Email address" style="width:100%;padding:10px 14px;border:1px solid #e3e3e3;border-radius:8px;font-size:14px;margin-bottom:16px;box-sizing:border-box;outline:none">' +
      '<button id="abk-reset-submit" style="width:100%;padding:11px;background:#0066cc;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;margin-bottom:8px">Send Reset Link</button>' +
      '<button id="abk-reset-cancel" style="width:100%;padding:8px;background:none;border:none;font-size:13px;color:#6b7280;cursor:pointer">Cancel</button>';

    overlay.appendChild(card);
    document.body.appendChild(overlay);

    function showMsg(txt, isErr) {
      var el = document.getElementById('abk-reset-msg');
      el.style.display = 'block';
      el.style.background = isErr ? '#fef2f2' : '#f0fdf4';
      el.style.color = isErr ? '#dc2626' : '#16a34a';
      el.textContent = txt;
    }

    document.getElementById('abk-reset-cancel').onclick = function() { overlay.remove(); };
    overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });

    document.getElementById('abk-reset-submit').onclick = function() {
      var email = document.getElementById('abk-reset-email').value.trim();
      if (!email || email.indexOf('@') === -1) { showMsg('Enter a valid email address.', true); return; }
      var btn = document.getElementById('abk-reset-submit');
      btn.disabled = true; btn.textContent = 'Sending...';
      fetch(BRIDGE_URL + '/org-admin/api/forgot-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: email })
      })
      .then(function(r) { return r.json(); })
      .then(function() {
        showMsg('If this email exists, a reset link has been sent. Check your inbox.', false);
        btn.style.display = 'none';
        document.getElementById('abk-reset-cancel').textContent = 'Back to Login';
      })
      .catch(function() {
        showMsg('Connection error. Please try again.', true);
        btn.disabled = false; btn.textContent = 'Send Reset Link';
      });
    };

    document.getElementById('abk-reset-email').addEventListener('keydown', function(e) {
      if (e.key === 'Enter') document.getElementById('abk-reset-submit').click();
    });
  }

  function patchForgotPassword() {
    var path = window.location.pathname;
    var onLoginPage = path === '/' || path.indexOf('login') !== -1;
    var existing = document.querySelector('[data-abk-forgot]');
    if (!onLoginPage) {
      if (existing) existing.remove();
      return;
    }
    if (existing && existing.isConnected) return;
    if (existing) existing.remove();
    var submitBtn = null;
    document.querySelectorAll('button[type="submit"]').forEach(function(btn) {
      var txt = (btn.textContent || '').toLowerCase().trim();
      if (!submitBtn && (txt.indexOf('sign in') !== -1 || txt.indexOf('log in') !== -1 ||
          txt.indexOf('login') !== -1 || txt.indexOf('continue') !== -1)) {
        submitBtn = btn;
      }
    });
    if (!submitBtn) return;
    var rect = submitBtn.getBoundingClientRect();
    if (!rect.width) return;
    var wrap = document.createElement('div');
    wrap.setAttribute('data-abk-forgot', '1');
    wrap.style.cssText = 'position:fixed;z-index:9999;text-align:center;pointer-events:auto;' +
      'left:' + rect.left + 'px;top:' + (rect.bottom + 8) + 'px;width:' + rect.width + 'px;';
    var a = document.createElement('a');
    a.href = '#';
    a.style.cssText = 'font-size:13px;color:#0066cc;text-decoration:none;';
    a.textContent = 'Forgot password?';
    a.addEventListener('click', function(e) { e.preventDefault(); showResetPasswordModal(); });
    wrap.appendChild(a);
    document.body.appendChild(wrap);
  }

  // 6. Rename "Admin Settings" → "Organization Settings" in the sidebar
  function renameAdminLink() {
    var els = document.querySelectorAll('a, button, li, [role="menuitem"]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      if ((el.innerText || '').trim() === 'Admin Settings' && !el.dataset.abkRenamed) {
        el.dataset.abkRenamed = '1';
        var walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null, false);
        var node;
        while ((node = walker.nextNode())) {
          if (node.textContent.trim() === 'Admin Settings') {
            node.textContent = node.textContent.replace('Admin Settings', 'Organization Settings');
            break;
          }
        }
      }
    }
  }

  // 6b. Move "Organization Settings" out of the Skills flyout and into the icon rail
  function moveOrgSettingsToRail() {
    var orgEl = null;
    document.querySelectorAll('button, a, li, [role="menuitem"]').forEach(function(el) {
      if (orgEl || el.hasAttribute('data-abk-rail-org')) return;
      if ((el.innerText || '').trim() === 'Organization Settings') orgEl = el;
    });
    if (orgEl) orgEl.style.display = 'none';

    if (document.querySelector('[data-abk-rail-org]')) return;
    if (!orgEl) return;

    var skillsBtn = null;
    document.querySelectorAll('button, a').forEach(function(el) {
      if (skillsBtn) return;
      var lbl = (el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
      if (lbl === 'Skills') skillsBtn = el;
    });
    if (!skillsBtn || !skillsBtn.parentElement) return;
    var rail = skillsBtn.parentElement;
    if (rail.children.length < 3) return;

    // Prefer cloning a non-active sibling's classes for clean (unselected) styling
    var template = null;
    for (var i = 0; i < rail.children.length; i++) {
      var sib = rail.children[i];
      if (sib !== skillsBtn && (sib.tagName === 'BUTTON' || sib.tagName === 'A')) { template = sib; break; }
    }
    if (!template) template = skillsBtn;

    var icon = orgEl.querySelector('svg');
    var newBtn = document.createElement(template.tagName);
    newBtn.className = template.className;
    newBtn.setAttribute('data-abk-rail-org', '1');
    newBtn.setAttribute('title', 'Organization Settings');
    newBtn.setAttribute('aria-label', 'Organization Settings');
    newBtn.innerHTML = icon ? icon.outerHTML : template.innerHTML;
    newBtn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      openOrgAdmin();
    });
    rail.appendChild(newBtn);
  }

  function openOrgAdmin() {
    // Use previously saved email (set via postMessage after first form submit)
    var savedEmail = localStorage.getItem('abk_admin_email');
    if (savedEmail && savedEmail.indexOf('@') !== -1) {
      console.log('[ABK] using saved email:', savedEmail);
      showOrgAdminModal('lc_email=' + encodeURIComponent(savedEmail));
      return;
    }

    // 1. Scan ALL localStorage keys for JWT or JSON containing email
    var lct = '';
    var emailFromStorage = '';
    for (var i = 0; i < localStorage.length; i++) {
      var k = localStorage.key(i);
      var v = localStorage.getItem(k) || '';

      // Check for JWT (3 dot-separated parts with valid base64 payload)
      if (!lct && v && v.split('.').length === 3) {
        try {
          var b64 = v.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');
          while (b64.length % 4) b64 += '=';
          var p = JSON.parse(atob(b64));
          if (p && (p.id || p.email || p.sub)) {
            lct = v;
            if (p.email) emailFromStorage = p.email;
            console.log('[ABK] JWT found in localStorage key:', k, '| has email:', !!p.email);
          }
        } catch(e) {}
      }

      // Check for JSON object that may contain email
      if (!emailFromStorage && v && v[0] === '{') {
        try {
          var parsed = JSON.parse(v);
          var fe = extractEmail(parsed);
          if (fe && fe.indexOf('@') !== -1) {
            emailFromStorage = fe;
            console.log('[ABK] email in localStorage key:', k, '→', fe);
          } else {
            // Handle persist:root style double-encoded JSON
            for (var field in parsed) {
              if (typeof parsed[field] === 'string' && parsed[field][0] === '{') {
                try {
                  var inner = JSON.parse(parsed[field]);
                  var ie = extractEmail(inner);
                  if (ie && ie.indexOf('@') !== -1) {
                    emailFromStorage = ie;
                    console.log('[ABK] email in nested localStorage key:', k + '.' + field, '→', ie);
                    break;
                  }
                } catch(e2) {}
              }
            }
          }
        } catch(e) {}
      }
    }

    if (emailFromStorage) {
      console.log('[ABK] using email from localStorage:', emailFromStorage);
      localStorage.setItem('abk_admin_email', emailFromStorage);
      showOrgAdminModal(lct ? 'lc_token=' + encodeURIComponent(lct) : 'lc_email=' + encodeURIComponent(emailFromStorage));
      return;
    }
    if (lct) {
      console.log('[ABK] using JWT (no email in payload)');
      showOrgAdminModal('lc_token=' + encodeURIComponent(lct));
      return;
    }

    // 2. Try LibreChat API endpoints (browser sends httpOnly cookie automatically)
    var endpoints = ['/api/user', '/api/auth/user', '/api/v1/user', '/api/auth/me', '/api/me'];
    var idx = 0;
    function tryNext() {
      if (idx >= endpoints.length) {
        console.warn('[ABK] all API endpoints failed — opening modal (email form will appear)');
        showOrgAdminModal('');
        return;
      }
      var ep = endpoints[idx++];
      fetch(ep, { credentials: 'include' })
        .then(function(r) {
          console.log('[ABK] ' + ep + ' → ' + r.status);
          return r.ok ? r.json() : Promise.reject(r.status);
        })
        .then(function(data) {
          console.log('[ABK] ' + ep + ' data:', JSON.stringify(data).slice(0, 200));
          var email = extractEmail(data);
          if (email && email.indexOf('@') !== -1) {
            console.log('[ABK] email from API:', email);
            localStorage.setItem('abk_admin_email', email);
            showOrgAdminModal('lc_email=' + encodeURIComponent(email));
          } else {
            tryNext();
          }
        })
        .catch(function(e) {
          console.warn('[ABK] ' + ep + ' failed:', e);
          tryNext();
        });
    }
    tryNext();
  }

  // Install ONE document-level capture listener — fires before React's root handler.
  // React (v17+) delegates events to div#root, so document capture = before React.
  var abkAdminInterceptorInstalled = false;
  function patchAdminLink() {
    if (abkAdminInterceptorInstalled) return;
    // Only install once the element actually exists in the DOM
    var found = false;
    var els = document.querySelectorAll('a, button, [role="menuitem"]');
    for (var i = 0; i < els.length; i++) {
      if ((els[i].innerText || '').trim() === 'Admin Settings') { found = true; break; }
    }
    if (!found) return;

    abkAdminInterceptorInstalled = true;
    document.addEventListener('click', function(e) {
      var el = e.target;
      for (var i = 0; i < 5 && el && el !== document.body; i++, el = el.parentElement) {
        var txt = (el.innerText || '').trim();
        if (txt === 'Admin Settings' || txt === 'Organization Settings') {
          e.preventDefault();
          e.stopPropagation();
          e.stopImmediatePropagation();
          openOrgAdmin();
          return;
        }
      }
    }, true);
  }

  // 7. Block registration for non-@abk.gr emails.
  // Disables the submit button in real-time as the user types.
  function patchRegistration() {
    var path = window.location.pathname;
    if (path.indexOf('register') === -1) return;
    var emailInput = document.querySelector(
      'input[type="email"], input[name="email"], input[autocomplete="email"], input[id*="email"]'
    );
    if (!emailInput || emailInput.getAttribute('data-abk-reg')) return;
    emailInput.setAttribute('data-abk-reg', '1');

    // Add a helper note below the field
    if (!document.getElementById('abk-reg-note')) {
      var note = document.createElement('p');
      note.id = 'abk-reg-note';
      note.style.cssText = 'font-size:12px;color:#9ca3af;margin:3px 0 0;';
      note.textContent = 'Only @abk.gr email addresses are accepted.';
      if (emailInput.parentNode) emailInput.parentNode.insertBefore(note, emailInput.nextSibling);
    }

    function validate() {
      var val = (emailInput.value || '').trim().toLowerCase();
      var submitBtn = document.querySelector('button[type="submit"]');
      var errEl = document.getElementById('abk-reg-err');
      var hasDomain = val.indexOf('@') !== -1;
      var allowed = !hasDomain || val.endsWith('@abk.gr');
      if (!allowed) {
        if (!errEl) {
          var err = document.createElement('p');
          err.id = 'abk-reg-err';
          err.style.cssText = 'font-size:12px;color:#dc2626;margin:4px 0 0;';
          err.textContent = 'Only @abk.gr addresses are allowed.';
          if (emailInput.parentNode) emailInput.parentNode.insertBefore(err, emailInput.nextSibling);
        }
        if (submitBtn) { submitBtn.disabled = true; submitBtn.style.opacity = '0.5'; }
      } else {
        if (errEl) errEl.remove();
        if (submitBtn) { submitBtn.disabled = false; submitBtn.style.opacity = ''; }
      }
    }

    emailInput.addEventListener('input', validate);
    emailInput.addEventListener('blur', validate);
  }

  // 8. Replace greeting icon (pencil/feather) with ABK logo
  function patchGreeting() {
    if (!document.body) return;
    try {
    // Use TreeWalker to find the exact text node with the greeting
    var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
    var node;
    while ((node = walker.nextNode())) {
      if (!/Good (morning|afternoon|evening)/i.test(node.textContent || '')) continue;
      var el = node.parentElement;
      if (!el || el.dataset.abkGreetDone) continue;
      el.dataset.abkGreetDone = '1';

      // Walk up until we find a container that also has an SVG sibling
      var container = el;
      for (var k = 0; k < 6; k++) {
        if (!container.parentElement) break;
        container = container.parentElement;
        if (container.querySelector('svg')) break;
      }

      // Hide ALL SVGs in the container (the feather/pencil icon)
      container.querySelectorAll('svg').forEach(function(s) {
        s.style.display = 'none';
        s.setAttribute('data-abk-greet-icon-hidden', '1');
      });

      // Hide the empty round avatar placeholder between the logo and the text
      container.querySelectorAll('[class*="rounded-full"]').forEach(function(s) {
        s.style.display = 'none';
        s.setAttribute('data-abk-greet-icon-hidden', '1');
      });

      // Insert ABK logo inline BEFORE the heading element (as a sibling)
      if (!document.querySelector('[data-abk-greet-logo]')) {
        var img = document.createElement('img');
        img.src = '/assets/abk-logo.png?abk=1';
        img.setAttribute('data-abk-greet-logo', '1');
        img.style.cssText = 'height:28px;width:auto;object-fit:contain;vertical-align:middle;margin-right:10px;display:inline-block;';
        // Insert at start of the parent container (same row as the text)
        var insertParent = el.parentElement || container;
        insertParent.insertBefore(img, insertParent.firstChild);
      }
      break;
    }
    } catch(e) {}
  }

  // 9. Hide Sign Up link and redirect /register → / (registration is admin-only via org settings)
  function hideSignUp() {
    // Redirect anyone who navigates directly to /register
    if (window.location.pathname.indexOf('register') !== -1) {
      window.location.replace('/');
      return;
    }
    // Hide "Don't have an account?" / "Sign up" links on login page
    var all = document.querySelectorAll('a, p, span, div');
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      var txt = (el.textContent || '').trim();
      if ((txt.indexOf('Sign up') !== -1 || txt.indexOf("Don't have an account") !== -1 ||
           txt.indexOf('Create an account') !== -1 || txt.indexOf('Register') !== -1) &&
          el.children.length <= 2) {
        el.style.display = 'none';
      }
    }
    // Also hide any <a> pointing to /register
    document.querySelectorAll('a[href*="register"]').forEach(function(a) {
      var p = a.parentElement;
      if (p && p !== document.body) p.style.display = 'none';
      else a.style.display = 'none';
    });
  }

  function tryPatch() {
    try { updateAppClass(); } catch(e) {}
    try { patchLoginText(); } catch(e) {}
    try { patchLogo(); } catch(e) {}
    try { patchForgotPassword(); } catch(e) {}
    try { hideSignUp(); } catch(e) {}
    try { patchAdminLink(); } catch(e) {}
    try { renameAdminLink(); } catch(e) {}
    try { moveOrgSettingsToRail(); } catch(e) {}
    try { patchGreeting(); } catch(e) {}
  }

  function setup() {
    new MutationObserver(tryPatch).observe(document.documentElement, { childList: true, subtree: true });
    tryPatch();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setup);
  } else {
    setup();
  }
})();
</script>
<!-- ABK_END -->`;

html = html.replace('</head>', inject + '\n</head>');
fs.writeFileSync('/app/client/dist/index.html', html);

// Replace logo.svg reference with ABK logo in all JS bundles
const assetsDir = '/app/client/dist/assets';
const jsFiles = fs.readdirSync(assetsDir).filter(f => f.endsWith('.js'));
let logoReplaced = 0;
jsFiles.forEach(f => {
  const p = `${assetsDir}/${f}`;
  const content = fs.readFileSync(p, 'utf8');
  if (content.includes('assets/logo.svg')) {
    fs.writeFileSync(p, content.replaceAll('assets/logo.svg', 'assets/abk-logo.png'));
    logoReplaced++;
    console.log(`  Logo reference replaced in: ${f}`);
  }
});

// ALSO overwrite logo.svg with an SVG wrapper containing the ABK logo as base64.
// This ensures ANY cached bundle version (old or new) always renders the ABK logo,
// bypassing service-worker cache issues entirely.
const pngData = fs.readFileSync('/app/client/dist/assets/abk-logo.png');
const base64 = pngData.toString('base64');
const svgWrapper = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     width="500" height="120" viewBox="0 0 500 120">
  <image xlink:href="data:image/png;base64,${base64}"
         x="0" y="0" width="500" height="120"
         preserveAspectRatio="xMidYMid meet"/>
</svg>`;
fs.writeFileSync(`${assetsDir}/logo.svg`, svgWrapper);
console.log('  logo.svg overwritten with ABK logo SVG wrapper.');

console.log(`ABK branding applied (idempotent)! Logo replaced in ${logoReplaced} JS file(s).`);
