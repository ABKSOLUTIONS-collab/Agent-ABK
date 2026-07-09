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
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

  /* Hide LibreChat footer and links */
  footer { display: none !important; }
  a[href*="librechat.ai"] { display: none !important; }

  /* ── Hide Sign Up / registration links via CSS (reliable, no JS race) ── */
  a[href*="register"] { display: none !important; }
  p:has(a[href*="register"]),
  div:has(> a[href*="register"]),
  span:has(a[href*="register"]) { display: none !important; }

  /* Hide raven SVG logo (React-safe: hide in CSS, don't remove from DOM) */
  svg[data-abk-hidden="1"] { display: none !important; }
  /* Greeting icon hidden by JS */
  [data-abk-greet-icon-hidden="1"] { display: none !important; }

  /* ══════════════════════════════════════════════════════════════
     ABK Solutions design tokens — mapped directly onto LibreChat's
     OWN CSS variables (html{} = light, .dark{} = dark), so every
     component that already consumes them re-themes automatically
     instead of guessing at individual Tailwind utility classes.
     ══════════════════════════════════════════════════════════════ */
  html, body, input, button, select, textarea {
    font-family: 'IBM Plex Sans', system-ui, sans-serif !important;
  }
  pre, code, kbd, samp, .font-mono {
    font-family: 'IBM Plex Mono', ui-monospace, monospace !important;
  }

  html {
    --text-primary: #14171c !important;
    --text-secondary: #586170 !important;
    --text-secondary-alt: #586170 !important;
    --text-tertiary: #8b94a1 !important;
    --text-destructive: #d64545 !important;
    --header-primary: #ffffff !important;
    --header-hover: #eef1f4 !important;
    --header-button-hover: #eef1f4 !important;
    --surface-active: #e6f1f9 !important;
    --surface-active-alt: #e6f1f9 !important;
    --surface-hover: #eef1f4 !important;
    --surface-hover-alt: #eef1f4 !important;
    --surface-primary: #ffffff !important;
    --surface-primary-alt: #f6f7f9 !important;
    --surface-primary-contrast: #fbfcfd !important;
    --surface-secondary: #fbfcfd !important;
    --surface-secondary-alt: #eef1f4 !important;
    --surface-tertiary: #eef1f4 !important;
    --surface-tertiary-alt: #ffffff !important;
    --surface-dialog: #ffffff !important;
    --surface-submit: #0071BC !important;
    --surface-submit-hover: #005a96 !important;
    --surface-destructive: #d64545 !important;
    --surface-destructive-hover: #b83a3a !important;
    --surface-chat: #ffffff !important;
    --border-light: #e6e9ee !important;
    --border-medium: #d5dae1 !important;
    --border-medium-alt: #d5dae1 !important;
    --border-heavy: #c3c9d1 !important;
    --border-xheavy: #a8afb9 !important;
    --border-destructive: #d64545 !important;
    --ring-primary: #0071BC !important;
    --ring: 204 100% 37% !important;
    --primary: 204 100% 37% !important;
    --primary-foreground: 0 0% 100% !important;
    --border: 218 19% 92% !important;
    --input: 215 17% 86% !important;
    --destructive: 0 64% 55% !important;
  }

  .dark {
    --text-primary: #eef1f5 !important;
    --text-secondary: #9aa5b1 !important;
    --text-secondary-alt: #9aa5b1 !important;
    --text-tertiary: #69737f !important;
    --text-destructive: #e46a6a !important;
    --header-primary: #101317 !important;
    --header-hover: #23282f !important;
    --header-button-hover: #23282f !important;
    --surface-active: #10293b !important;
    --surface-active-alt: #10293b !important;
    --surface-hover: #23282f !important;
    --surface-hover-alt: #23282f !important;
    --surface-primary: #101317 !important;
    --surface-primary-alt: #15181d !important;
    --surface-primary-contrast: #14171b !important;
    --surface-secondary: #14171b !important;
    --surface-secondary-alt: #23282f !important;
    --surface-tertiary: #1a1e24 !important;
    --surface-tertiary-alt: #1a1e24 !important;
    --surface-dialog: #1a1e24 !important;
    --surface-submit: #2e97e0 !important;
    --surface-submit-hover: #4aa8e8 !important;
    --surface-destructive: #e46a6a !important;
    --surface-destructive-hover: #ea8484 !important;
    --surface-chat: #101317 !important;
    --border-light: #282e35 !important;
    --border-medium: #353c44 !important;
    --border-medium-alt: #353c44 !important;
    --border-heavy: #454d57 !important;
    --border-xheavy: #5a636e !important;
    --border-destructive: #e46a6a !important;
    --ring-primary: #2e97e0 !important;
    --ring: 205 74% 53% !important;
    --primary: 205 74% 53% !important;
    --primary-foreground: 0 0% 100% !important;
    --border: 212 14% 18% !important;
    --input: 212 12% 24% !important;
    --destructive: 0 69% 65% !important;
  }

  /* Focus ring on plain (non-Radix) inputs, matching the brief's spec */
  input:focus, input:focus-visible, textarea:focus, textarea:focus-visible {
    border-color: #0071BC !important;
    box-shadow: 0 0 0 3px rgba(0,113,188,.15) !important;
    outline: none !important;
  }
  .dark input:focus, .dark input:focus-visible, .dark textarea:focus, .dark textarea:focus-visible {
    border-color: #2e97e0 !important;
    box-shadow: 0 0 0 3px rgba(46,151,224,.2) !important;
  }

  /* 3px brand accent bar pinned to the very top of the app (echoes the logo blue) */
  #abk-top-accent {
    position: fixed; top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #0071BC, #22a7e6);
    z-index: 6; pointer-events: none;
  }

  /* Home greeting subtitle, adapts to light/dark brand blue */
  [data-abk-greet-sub] { color: var(--ring-primary) !important; }

  .abk-scroll::-webkit-scrollbar { width: 10px; height: 10px; }
  .abk-scroll::-webkit-scrollbar-thumb { background: var(--border-medium); border-radius: 8px; border: 3px solid transparent; background-clip: padding-box; }
  .abk-scroll::-webkit-scrollbar-track { background: transparent; }
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
    // The greeting renders as a React Fragment with two children: an
    // sr-only span (screen readers) and this visually-animated sibling
    // (per-letter spans, aria-hidden="true") — confirmed directly from the
    // LibreChat bundle source. Target it directly instead of guessing at
    // sibling relationships from a matched text node.
    var candidates = document.querySelectorAll('.split-parent[aria-hidden="true"]');
    console.log('[ABK greet] candidates:', candidates.length);
    for (var ci = 0; ci < candidates.length; ci++) {
      var el = candidates[ci];
      console.log('[ABK greet] candidate', ci, 'done=', el.dataset.abkGreetDone, 'text=', JSON.stringify(el.textContent));
      if (el.dataset.abkGreetDone) continue;
      if (!/Good (morning|afternoon|evening)/i.test(el.textContent || '')) continue;
      el.dataset.abkGreetDone = '1';
      console.log('[ABK greet] MATCHED, wrapping element');

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

      // The heading's parent is a flex ROW (icon + text side by side), so a
      // plain sibling <p> would render BESIDE the heading, not under it.
      // Wrap the heading in its own column stack so the subtitle can sit
      // directly below "Good afternoon, {name}" as intended.
      var insertParent = el.parentElement || container;
      console.log('[ABK greet] insertParent found:', !!insertParent, insertParent === el.parentElement ? 'is-el-parent' : 'is-container');
      var stack = document.createElement('div');
      stack.setAttribute('data-abk-greet-stack', '1');
      stack.style.cssText = 'display:flex;flex-direction:column;';
      insertParent.insertBefore(stack, el);
      stack.appendChild(el);
      console.log('[ABK greet] stack created and attached:', document.body.contains(stack));

      // Insert ABK logo + vertical accent bar before the stack (same row)
      if (!document.querySelector('[data-abk-greet-logo]')) {
        var img = document.createElement('img');
        img.src = '/assets/abk-logo.png?abk=1';
        img.setAttribute('data-abk-greet-logo', '1');
        img.style.cssText = 'height:26px;width:auto;object-fit:contain;vertical-align:middle;margin-right:14px;display:inline-block;';
        insertParent.insertBefore(img, stack);

        var bar = document.createElement('span');
        bar.setAttribute('data-abk-greet-bar', '1');
        bar.style.cssText = 'display:inline-block;width:2px;height:42px;background:var(--ring-primary);margin-right:14px;vertical-align:middle;';
        insertParent.insertBefore(bar, stack);
      }

      // Primary-colored subtitle below the greeting heading
      var sub = document.createElement('p');
      sub.setAttribute('data-abk-greet-sub', '1');
      sub.textContent = 'Πώς μπορώ να βοηθήσω σήμερα;';
      sub.style.cssText = 'margin:4px 0 0;font-size:14px;font-weight:500;';
      stack.appendChild(sub);
      console.log('[ABK greet] DONE, stack still attached:', document.body.contains(stack));
      break;
    }
    } catch(e) { console.error('[ABK greet] EXCEPTION', e); }
  }

  // 8b. Lightweight toast for not-yet-implemented actions (UI only — no fake state/data)
  function showAbkToast(msg) {
    var old = document.getElementById('abk-toast');
    if (old) old.remove();
    var t = document.createElement('div');
    t.id = 'abk-toast';
    t.textContent = msg;
    t.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);' +
      'background:var(--text-primary,#14171c);color:#fff;padding:10px 18px;border-radius:8px;' +
      'font-size:13px;font-weight:500;z-index:100000;box-shadow:0 12px 44px rgba(16,24,40,.25);opacity:0;' +
      'transition:opacity .15s ease;';
    document.body.appendChild(t);
    requestAnimationFrame(function() { t.style.opacity = '1'; });
    setTimeout(function() {
      t.style.opacity = '0';
      setTimeout(function() { t.remove(); }, 200);
    }, 2200);
  }

  // 8c. Generic full-screen panel modal (same chrome as Organization Settings) for Projects / MCP
  function showAbkPanelModal(title, bodyHtml) {
    var existing = document.getElementById('abk-panel-overlay');
    if (existing) existing.remove();

    var overlay = document.createElement('div');
    overlay.id = 'abk-panel-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:99999;background:rgba(16,20,26,.5);display:flex;align-items:center;justify-content:center;';

    var modal = document.createElement('div');
    modal.style.cssText = 'position:relative;width:92%;max-width:960px;height:85vh;background:var(--surface-dialog,#fff);border-radius:14px;overflow:hidden;box-shadow:0 24px 80px rgba(0,0,0,.35);display:flex;flex-direction:column;';

    var header = document.createElement('div');
    header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:14px 20px;background:var(--surface-dialog,#fff);border-bottom:1px solid var(--border-light,#e6e9ee);flex-shrink:0;';
    var titleEl = document.createElement('span');
    titleEl.style.cssText = 'font-weight:600;font-size:14px;color:var(--ring-primary,#0071BC);';
    titleEl.textContent = title;
    var closeBtn = document.createElement('button');
    closeBtn.innerHTML = '&#215;';
    closeBtn.style.cssText = 'background:none;border:none;font-size:22px;cursor:pointer;color:var(--text-tertiary,#8b94a1);padding:0 4px;line-height:1;';
    closeBtn.onclick = function() { overlay.remove(); };
    header.appendChild(titleEl);
    header.appendChild(closeBtn);

    var body = document.createElement('div');
    body.className = 'abk-scroll';
    body.style.cssText = 'flex:1;overflow-y:auto;padding:24px 28px;color:var(--text-primary,#14171c);';
    body.innerHTML = bodyHtml;

    modal.appendChild(header);
    modal.appendChild(body);
    overlay.appendChild(modal);
    overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
    return body;
  }

  // 8d. Projects panel — UI shell only, genuinely empty (no seeded fake projects)
  function buildProjectsHtml() {
    return (
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:18px;">' +
        '<h1 style="font-size:26px;font-weight:700;letter-spacing:-.02em;margin:0;">Projects</h1>' +
        '<button id="abk-projects-new" style="height:38px;padding:0 16px;border:none;border-radius:8px;background:var(--ring-primary,#0071BC);color:#fff;font-size:14px;font-weight:600;cursor:pointer;white-space:nowrap;">+ New project</button>' +
      '</div>' +
      '<input type="text" placeholder="Search projects…" style="width:100%;height:44px;padding:0 14px;border:1px solid var(--border-medium,#d5dae1);border-radius:10px;background:var(--surface-primary,#fff);color:var(--text-primary,#14171c);font-size:14px;outline:none;box-sizing:border-box;margin-bottom:18px;">' +
      '<div style="display:flex;gap:20px;border-bottom:1px solid var(--border-light,#e6e9ee);margin-bottom:32px;">' +
        '<button class="abk-proj-tab" data-tab="your" style="background:none;border:none;padding:0 0 10px;font-size:14px;font-weight:600;color:var(--ring-primary,#0071BC);border-bottom:2px solid var(--ring-primary,#0071BC);cursor:pointer;">Your projects</button>' +
        '<button class="abk-proj-tab" data-tab="team" style="background:none;border:none;padding:0 0 10px;font-size:14px;font-weight:500;color:var(--text-tertiary,#8b94a1);border-bottom:2px solid transparent;cursor:pointer;">Team</button>' +
        '<button class="abk-proj-tab" data-tab="shared" style="background:none;border:none;padding:0 0 10px;font-size:14px;font-weight:500;color:var(--text-tertiary,#8b94a1);border-bottom:2px solid transparent;cursor:pointer;">Shared with you</button>' +
      '</div>' +
      '<div style="text-align:center;color:var(--text-tertiary,#8b94a1);font-size:13.5px;padding:60px 0;">No projects yet.</div>'
    );
  }
  function wireProjectsPanel(body) {
    var tabs = body.querySelectorAll('.abk-proj-tab');
    tabs.forEach(function(tab) {
      tab.addEventListener('click', function() {
        tabs.forEach(function(t) {
          t.style.color = 'var(--text-tertiary,#8b94a1)';
          t.style.fontWeight = '500';
          t.style.borderBottomColor = 'transparent';
        });
        tab.style.color = 'var(--ring-primary,#0071BC)';
        tab.style.fontWeight = '600';
        tab.style.borderBottomColor = 'var(--ring-primary,#0071BC)';
      });
    });
    var newBtn = body.querySelector('#abk-projects-new');
    if (newBtn) newBtn.addEventListener('click', function() { showAbkToast('Έρχεται σύντομα'); });
  }

  // 8e. MCP Settings panel — UI shell only, genuinely empty (no seeded fake servers)
  function buildMcpHtml() {
    return (
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">' +
        '<h1 style="font-size:26px;font-weight:700;letter-spacing:-.02em;margin:0;">MCP Servers</h1>' +
        '<button id="abk-mcp-add" style="height:38px;padding:0 16px;border:none;border-radius:8px;background:var(--ring-primary,#0071BC);color:#fff;font-size:14px;font-weight:600;cursor:pointer;white-space:nowrap;">+ Add MCP server</button>' +
      '</div>' +
      '<p style="color:var(--text-tertiary,#8b94a1);font-size:14px;margin:0 0 20px;">Σύνδεσε MCP servers για να δώσεις στο ABK Assistant και στους agents πρόσβαση σε εργαλεία και connectors.</p>' +
      '<input type="text" placeholder="Filter MCP servers by name…" style="width:100%;height:44px;padding:0 14px;border:1px solid var(--border-medium,#d5dae1);border-radius:10px;background:var(--surface-primary,#fff);color:var(--text-primary,#14171c);font-size:14px;outline:none;box-sizing:border-box;margin-bottom:28px;">' +
      '<div style="border:1.5px dashed var(--border-medium,#d5dae1);border-radius:12px;padding:52px 20px;text-align:center;">' +
        '<div style="width:52px;height:52px;border-radius:10px;background:var(--surface-active,#e6f1f9);display:flex;align-items:center;justify-content:center;margin:0 auto 16px;">' +
          '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" style="stroke:var(--ring-primary,#0071BC);" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="7" rx="1.5"/><rect x="3" y="13" width="18" height="7" rx="1.5"/><line x1="7" y1="7.5" x2="7.1" y2="7.5"/><line x1="7" y1="16.5" x2="7.1" y2="16.5"/></svg>' +
        '</div>' +
        '<div style="font-size:15px;font-weight:600;margin-bottom:6px;">No MCP servers yet</div>' +
        '<div style="font-size:13.5px;color:var(--text-tertiary,#8b94a1);margin-bottom:20px;">Create your first MCP server to get started</div>' +
        '<button id="abk-mcp-add-2" style="height:38px;padding:0 16px;border:none;border-radius:8px;background:var(--ring-primary,#0071BC);color:#fff;font-size:14px;font-weight:600;cursor:pointer;">+ Add MCP server</button>' +
      '</div>'
    );
  }
  function wireMcpPanel(body) {
    ['#abk-mcp-add', '#abk-mcp-add-2'].forEach(function(sel) {
      var btn = body.querySelector(sel);
      if (btn) btn.addEventListener('click', function() { showAbkToast('Έρχεται σύντομα'); });
    });
  }

  // 8f. Add Projects + MCP Settings icons to the left icon rail (cloning a
  // sibling's classes so they visually match, same technique as the
  // Organization Settings rail icon above).
  function addAbkRailIcon(dataAttr, titleText, svgHtml, onClick) {
    if (document.querySelector('[' + dataAttr + ']')) return;
    var skillsBtn = null;
    document.querySelectorAll('button, a').forEach(function(el) {
      if (skillsBtn) return;
      var lbl = (el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
      if (lbl === 'Skills') skillsBtn = el;
    });
    if (!skillsBtn || !skillsBtn.parentElement) return;
    var rail = skillsBtn.parentElement;
    if (rail.children.length < 3) return;

    var template = null;
    for (var i = 0; i < rail.children.length; i++) {
      var sib = rail.children[i];
      if (sib !== skillsBtn && (sib.tagName === 'BUTTON' || sib.tagName === 'A')) { template = sib; break; }
    }
    if (!template) template = skillsBtn;

    var newBtn = document.createElement(template.tagName);
    newBtn.className = template.className;
    newBtn.setAttribute(dataAttr, '1');
    newBtn.setAttribute('title', titleText);
    newBtn.setAttribute('aria-label', titleText);
    newBtn.innerHTML = svgHtml;
    newBtn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      onClick();
    });
    rail.appendChild(newBtn);
  }

  function patchProjectsAndMcpRail() {
    addAbkRailIcon('data-abk-rail-projects', 'Projects',
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"/></svg>',
      function() {
        var body = showAbkPanelModal('Projects', buildProjectsHtml());
        wireProjectsPanel(body);
      });
    addAbkRailIcon('data-abk-rail-mcp', 'MCP Settings',
      '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="7" rx="1.5"/><rect x="3" y="13" width="18" height="7" rx="1.5"/><line x1="7" y1="7.5" x2="7.1" y2="7.5"/><line x1="7" y1="16.5" x2="7.1" y2="16.5"/></svg>',
      function() {
        var body = showAbkPanelModal('MCP Servers', buildMcpHtml());
        wireMcpPanel(body);
      });
  }

  // 8g. Add a "Connectors" tab to the native Settings modal. UI shell only —
  // every connector is shown as NOT connected (the real, current state,
  // since there is no OAuth backend wired up yet) rather than seeded/fake data.
  var ABK_CONNECTORS = [
    { key: 'gdrive', name: 'Google Drive', desc: 'Access files from Google Drive.', color: '#1a73e8', initials: 'GD' },
    { key: 'gmail', name: 'Gmail', desc: 'Read and search your Gmail messages.', color: '#d64545', initials: 'GM' },
    { key: 'gcal', name: 'Google Calendar', desc: 'View and manage calendar events.', color: '#1a73e8', initials: 'GC' },
    { key: 'slack', name: 'Slack', desc: 'Search and post to Slack channels.', color: '#4a154b', initials: 'SL' },
    { key: 'github', name: 'GitHub', desc: 'Access repositories and issues.', color: '#24292f', initials: 'GH' },
    { key: 'notion', name: 'Notion', desc: 'Search and read Notion pages.', color: '#37352f', initials: 'NO' },
    { key: 'confluence', name: 'Confluence', desc: 'Search Confluence spaces and pages.', color: '#0071BC', initials: 'CF' }
  ];

  function buildConnectorsHtml() {
    var rows = ABK_CONNECTORS.map(function(c) {
      return (
        '<div style="display:flex;align-items:center;gap:14px;padding:14px 0;border-bottom:1px solid var(--border-light,#e6e9ee);">' +
          '<div style="width:36px;height:36px;border-radius:8px;background:' + c.color + ';color:#fff;font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;flex-shrink:0;">' + c.initials + '</div>' +
          '<div style="flex:1;min-width:0;">' +
            '<div style="font-size:14px;font-weight:600;">' + c.name + '</div>' +
            '<div style="font-size:12.5px;color:var(--text-tertiary,#8b94a1);margin-top:1px;">' + c.desc + '</div>' +
          '</div>' +
          '<button class="abk-connector-btn" data-key="' + c.key + '" style="height:34px;padding:0 16px;border:1px solid var(--ring-primary,#0071BC);border-radius:8px;background:none;color:var(--ring-primary,#0071BC);font-size:13px;font-weight:600;cursor:pointer;flex-shrink:0;white-space:nowrap;">Connect</button>' +
        '</div>'
      );
    }).join('');
    return (
      '<h2 style="font-size:16px;font-weight:600;margin:0 0 4px;">Connectors</h2>' +
      '<p style="font-size:13.5px;color:var(--text-tertiary,#8b94a1);margin:0 0 16px;">Connect third-party services to give ABK Assistant more context.</p>' +
      '<div style="display:inline-block;font-size:12.5px;font-weight:600;padding:4px 10px;border-radius:6px;background:var(--surface-active,#e6f1f9);color:var(--ring-primary,#0071BC);margin-bottom:8px;">0 / ' + ABK_CONNECTORS.length + ' connected</div>' +
      '<div>' + rows + '</div>'
    );
  }
  function wireConnectorsPanel(panel) {
    panel.querySelectorAll('.abk-connector-btn').forEach(function(btn) {
      btn.addEventListener('click', function() { showAbkToast('Έρχεται σύντομα'); });
    });
  }

  function findSettingsTabList() {
    var accountTab = null;
    document.querySelectorAll('button, [role="tab"], a').forEach(function(el) {
      if (accountTab) return;
      if ((el.textContent || '').trim() === 'Account') accountTab = el;
    });
    if (!accountTab || !accountTab.parentElement) return null;
    var tabList = accountTab.parentElement;
    if (tabList.children.length < 4) return null; // sanity: a real settings nav has several tabs
    return { tabList: tabList, accountTab: accountTab };
  }

  function injectConnectorsTab() {
    if (document.querySelector('[data-abk-connectors-tab]')) return;
    var found = findSettingsTabList();
    if (!found) return;
    var newTab = found.accountTab.cloneNode(true);
    newTab.removeAttribute('id');
    newTab.setAttribute('data-abk-connectors-tab', '1');
    var walker = document.createTreeWalker(newTab, NodeFilter.SHOW_TEXT, null, false);
    var node;
    while ((node = walker.nextNode())) {
      if (node.textContent.trim() === 'Account') { node.textContent = 'Connectors'; break; }
    }
    found.tabList.insertBefore(newTab, found.accountTab);

    newTab.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      showConnectorsPanel();
    }, true);
  }

  function showConnectorsPanel() {
    var tab = document.querySelector('[data-abk-connectors-tab]');
    if (!tab || !tab.parentElement) return;
    var tabList = tab.parentElement;
    var dialogBody = tabList.parentElement;
    if (!dialogBody) return;
    var contentPane = null;
    for (var i = 0; i < dialogBody.children.length; i++) {
      if (dialogBody.children[i] !== tabList) { contentPane = dialogBody.children[i]; break; }
    }
    if (!contentPane) return;

    var panel = document.getElementById('abk-connectors-panel');
    if (!panel) {
      panel = document.createElement('div');
      panel.id = 'abk-connectors-panel';
      panel.className = 'abk-scroll';
      panel.style.cssText = 'position:absolute;inset:0;background:var(--surface-dialog,#fff);overflow:auto;z-index:5;padding:20px 24px;';
      panel.innerHTML = buildConnectorsHtml();
      if (getComputedStyle(contentPane).position === 'static') contentPane.style.position = 'relative';
      contentPane.appendChild(panel);
      wireConnectorsPanel(panel);
    }
    panel.style.display = 'block';
    tab.style.background = 'var(--surface-active)';
    tab.style.color = 'var(--ring-primary)';
  }

  var abkConnectorsClickInstalled = false;
  function hideConnectorsPanelIfClickedElsewhere() {
    if (abkConnectorsClickInstalled) return;
    abkConnectorsClickInstalled = true;
    document.addEventListener('click', function(e) {
      var el = e.target;
      var isOurTab = false;
      for (var i = 0; i < 6 && el && el !== document.body; i++, el = el.parentElement) {
        if (el.hasAttribute && el.hasAttribute('data-abk-connectors-tab')) { isOurTab = true; break; }
      }
      if (isOurTab) return;
      var panel = document.getElementById('abk-connectors-panel');
      if (panel) panel.style.display = 'none';
      var tab = document.querySelector('[data-abk-connectors-tab]');
      if (tab) { tab.style.background = ''; tab.style.color = ''; }
    }, true);
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

  // 3px brand accent bar pinned to the very top of the viewport (added once, persists across routes)
  function patchTopAccentBar() {
    if (document.getElementById('abk-top-accent')) return;
    if (!document.body) return;
    var bar = document.createElement('div');
    bar.id = 'abk-top-accent';
    document.body.appendChild(bar);
  }

  function tryPatch() {
    try { updateAppClass(); } catch(e) {}
    try { patchTopAccentBar(); } catch(e) {}
    try { patchLoginText(); } catch(e) {}
    try { patchLogo(); } catch(e) {}
    try { patchForgotPassword(); } catch(e) {}
    try { hideSignUp(); } catch(e) {}
    try { patchAdminLink(); } catch(e) {}
    try { renameAdminLink(); } catch(e) {}
    try { moveOrgSettingsToRail(); } catch(e) {}
    try { patchProjectsAndMcpRail(); } catch(e) {}
    try { injectConnectorsTab(); } catch(e) {}
    try { hideConnectorsPanelIfClickedElsewhere(); } catch(e) {}
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
