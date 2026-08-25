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

  /* ── Sidebar restrictions ──────────────────────────────────────────
     Prompts / Memories / Bookmarks / Attach Files: removed for everyone.
     Agent Builder / Skills: hidden by default, unhidden by JS
     (applyAdminOnlyVisibility) only once the signed-in user is confirmed
     an org ADMIN/OWNER — CSS-first so there's no flash of a forbidden nav
     item for plain USERs while the role check is still in flight. */
  button[aria-label="Prompts"],
  button[aria-label="Memories"],
  button[aria-label="Bookmarks"],
  button[aria-label="Attach Files"] { display: none !important; }
  button[aria-label="Agent Builder"]:not([data-abk-admin-ok]),
  button[aria-label="Skills"]:not([data-abk-admin-ok]) { display: none !important; }
  /* Agent sharing is primary-owner-only — label is dynamic ("Share {agent name}") */
  button[aria-label^="Share "]:not([data-abk-owner-ok]) { display: none !important; }
  /* Remote Access (next to Delete in the Agent panel) is hidden for everyone */
  button[title="Remote Access"] { display: none !important; }

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

  /* SSO-only login: hide the native email/password form entirely */
  form[aria-label="Login form"] { display: none !important; }

  /* Style the OpenID SSO link as the primary blue CTA */
  a[data-testid="openid"] {
    background: var(--surface-submit) !important;
    border-color: var(--surface-submit) !important;
    color: #fff !important;
  }
  a[data-testid="openid"]:hover {
    background: var(--surface-submit-hover) !important;
  }
</style>
<script>
// Kill LibreChat's service worker outright — don't just unregister existing
// ones after the fact, actively block new registrations too. A stale SW
// precache is a known cause of "first login after logout fails and bounces
// back to /login, second attempt works" (see LibreChat issue #11534): the
// SW can keep serving an old cached index.html/bundle to the very page that
// needs to observe the freshly-set auth cookie, so the client-side session
// check runs against stale JS and fails once, then self-corrects the next
// load. Unregistering after the fact leaves a window where LibreChat's own
// bundle can re-register a SW before we get to it; overriding register()
// closes that gap, and clearing Cache Storage removes anything a past SW
// already cached, not just its registration.
function abkClearCaches() {
  if (!(window.caches && caches.keys)) return Promise.resolve();
  return caches.keys().then(function(keys) {
    return Promise.all(keys.map(function(k) { return caches.delete(k); }));
  }).catch(function() {});
}
abkClearCaches();
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register = function() {
    return Promise.reject(new Error('Service worker registration disabled (ABK patch)'));
  };
  navigator.serviceWorker.getRegistrations().then(function(regs) {
    var hadRegs = regs.length > 0;
    regs.forEach(function(r) { r.unregister(); });
    // A page already controlled by a SW stays controlled until it's
    // reloaded, even after unregister() — so force one reload to fully
    // release control, but only the one time there was actually something
    // to clean up (every load after that finds zero registrations, so this
    // can't loop).
    if (hadRegs) {
      abkClearCaches().then(function() {
        console.log('[ABK] SW unregistered + caches cleared, reloading...');
        window.location.reload();
      });
    }
  }).catch(function() {});
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
      // reload), a cached org-admin session token for the PREVIOUS user in
      // the same tab would otherwise carry over to whoever logs in next
      // (the token itself is still valid and still says the old user's
      // email — Organization Settings would silently open as them). Clear
      // it here so the next login always starts a fresh org-admin session.
      try { sessionStorage.removeItem('abk_org_session_token'); } catch(ex) {}
      localStorage.removeItem('abk_admin_email'); // legacy key from the old (insecure) mechanism
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

  // 4. Open Organization Settings as a modal overlay iframe.
  //
  // Auth is handled entirely server-side now: org-admin.ts shows its own
  // real login form (email + password, verified with bcrypt against the
  // same hash LibreChat itself checks) whenever there's no valid signed
  // session token yet — see the security hardening pass that replaced the
  // old "trust whatever email/token the client claims" mechanism. This
  // file no longer guesses or intercepts anything to establish identity;
  // it only caches the SIGNED token the server hands back after a real
  // login, so re-opening the panel in the same tab doesn't need a re-login.
  function showOrgAdminModal(extraParams) {
    var existing = document.getElementById('abk-org-overlay');
    if (existing) { existing.remove(); return; }

    var dark = isDarkMode();
    var url = BRIDGE_URL + '/org-admin?embed=1' + (extraParams ? '&' + extraParams : '');

    var overlay = document.createElement('div');
    overlay.id = 'abk-org-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;';

    var modal = document.createElement('div');
    modal.id = 'abk-org-modal';
    modal.style.cssText = 'position:relative;width:92%;max-width:1060px;height:88vh;background:' + (dark ? '#101317' : '#f9f9f9') + ';border-radius:14px;overflow:hidden;box-shadow:0 24px 80px rgba(0,0,0,.35);display:flex;flex-direction:column;';

    var header = document.createElement('div');
    header.id = 'abk-org-header';
    header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:14px 20px;background:' + (dark ? '#15181d' : '#fff') + ';border-bottom:1px solid ' + (dark ? '#282e35' : '#e8e8e8') + ';flex-shrink:0;';
    var title = document.createElement('span');
    title.id = 'abk-org-title';
    title.style.cssText = 'font-weight:600;font-size:14px;color:' + (dark ? '#2e97e0' : '#0071BC') + ';';
    title.textContent = 'Organization Settings';
    var closeBtn = document.createElement('button');
    closeBtn.id = 'abk-org-close';
    closeBtn.innerHTML = '&#215;';
    closeBtn.style.cssText = 'background:none;border:none;font-size:22px;cursor:pointer;color:' + (dark ? '#9aa5b1' : '#888') + ';padding:0 4px;line-height:1;';
    closeBtn.onclick = function() { overlay.remove(); };
    header.appendChild(title);
    header.appendChild(closeBtn);

    var iframe = document.createElement('iframe');
    iframe.id = 'abk-org-iframe';
    iframe.src = url;
    iframe.style.cssText = 'flex:1;width:100%;border:none;';

    modal.appendChild(header);
    modal.appendChild(iframe);
    overlay.appendChild(modal);
    overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });
    document.body.appendChild(overlay);
  }

  var BRIDGE_URL = 'https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io';
  var ORG_SESSION_KEY = 'abk_org_session_token';
  var ORG_EMAIL_HINT_KEY = 'abk_org_email_hint'; // UX-only prefill — never trusted as proof of identity
  localStorage.removeItem('abk_org_role');
  localStorage.removeItem('abk_admin_email'); // legacy key from the old (insecure) mechanism

  // Receive the signed session token back from the org-admin iframe after a
  // real, password-verified login, so re-opening Organization Settings
  // within the same browser tab doesn't require signing in again.
  window.addEventListener('message', function(evt) {
    if (!evt || !evt.data) return;
    if (evt.data.type === 'abk_org_session' && evt.data.token) {
      try { sessionStorage.setItem(ORG_SESSION_KEY, evt.data.token); } catch(ex) {}
      if (evt.data.email) { try { localStorage.setItem(ORG_EMAIL_HINT_KEY, evt.data.email); } catch(ex) {} }
    }
  }, false);

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

  // 5b. Hide the "Or" divider between the (now-hidden) local form and the SSO button
  function hideOrDivider() {
    var path = window.location.pathname;
    var onLoginPage = path === '/' || path.indexOf('login') !== -1;
    if (!onLoginPage) return;
    document.querySelectorAll('main div').forEach(function(el) {
      if (el.dataset.abkOrHidden) return;
      if (el.children.length === 1 && (el.children[0].textContent || '').trim().toLowerCase() === 'or') {
        el.dataset.abkOrHidden = '1';
        el.style.display = 'none';
      }
    });
  }

  // 6. LibreChat renders its OWN native "Admin Settings" button at the
  // bottom of every sidebar sub-panel (Memories, MCP Settings, etc.) for
  // native-role ADMIN users — completely separate from, and not gated the
  // same way as, our own Organization Settings rail icon below. An earlier
  // version of this patch renamed its text to "Organization Settings",
  // which just produced a confusing SECOND (ungated) entry point into the
  // feature, appearing at the bottom of whichever panel happens to be open.
  // Hide it outright everywhere instead — the rail icon (6b) is the only
  // supported entry point now.
  function hideNativeAdminLink() {
    var els = document.querySelectorAll('a, button, li, [role="menuitem"]');
    for (var i = 0; i < els.length; i++) {
      var el = els[i];
      if ((el.innerText || '').trim() === 'Admin Settings' && !el.dataset.abkHiddenNative) {
        el.dataset.abkHiddenNative = '1';
        el.style.display = 'none';
      }
    }
  }

  // 6b. Add an "Organization Settings" icon to the same rail as "Skills".
  var ORG_RAIL_ICON_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="width:20px;height:20px"><path d="M3 21h18"/><path d="M6 21V8l6-4 6 4v13"/><path d="M10 21v-6h4v6"/><path d="M9 11h.01M15 11h.01M9 15h.01M15 15h.01"/></svg>';

  // Gate the rail icon on the user actually being an org ADMIN/OWNER.
  // org-admin.ts keeps LibreChat's own 'role' field ('ADMIN' or 'USER') in
  // sync with ABK's org tier specifically so we can read it here — a plain
  // USER should never even see the entry point, not just be blocked from
  // acting once inside it.
  //
  // IMPORTANT: LibreChat's own REST calls (e.g. /api/user) require an
  // Authorization header carrying a JWT that LibreChat keeps in memory only
  // (never in a cookie or localStorage) — an outside script has no way to
  // read it, so a plain fetch('/api/user') always comes back 401 "No auth
  // token", regardless of the real role. (An earlier version of this patch
  // did exactly that and broke Organization Settings for every role,
  // including admins.) /api/auth/refresh is the one endpoint that instead
  // authenticates purely via the httpOnly refresh-token cookie the browser
  // already carries, and its response includes the full user object
  // (with role) — that's what LibreChat's own app calls to silently renew
  // its session, so hitting it once here piggybacks on a flow LibreChat
  // already expects rather than an unexpected 401 that might trip some
  // global "log the user out on any 401" handling.
  //
  // IMPORTANT #2: /api/auth/refresh rotates a single-use refresh-token
  // cookie on every call. LibreChat's own client already calls it once on
  // initial load to restore the session — if OUR call fires around the same
  // time, whichever of the two loses the race gets a 401 against an
  // already-rotated cookie. When LibreChat's OWN call is the loser, its
  // client treats that 401 as "not logged in" and bounces back to /login
  // even though the sign-in genuinely succeeded (confirmed via server logs:
  // the OpenID exchange always reports "login success" server-side, so this
  // was never an actual auth failure, just us stepping on LibreChat's own
  // token rotation). Delaying our call gives LibreChat's own bootstrap
  // refresh a head start so we rotate the token only after it's done with
  // it, not concurrently.
  // Keep in sync with the PRIMARY_OWNER_EMAIL env var on the agent365-bridge
  // container. LibreChat's own role only distinguishes ADMIN/USER — it has
  // no concept of our org's single PRIMARY OWNER (org-admin.ts syncs both
  // ADMIN and OWNER org tiers down to the same native 'ADMIN' role so the
  // native Admin Settings link stays visible for both) — and org-admin's own
  // API no longer accepts a bare, unauthenticated email to resolve a tier
  // (that "trust whatever email the client claims" mechanism was removed in
  // a security hardening pass; it now requires a real signed-in session).
  // A direct email match against the one fixed owner address is simpler and
  // more reliable than trying to reach an authenticated endpoint from here.
  var ABK_PRIMARY_OWNER_EMAIL = 'snikolaou@abk.gr';

  var abkIsOrgAdmin = null; // null = not checked yet, else boolean
  var abkIsPrimaryOwner = false;
  var abkRoleCheckScheduled = false;
  function checkOrgAdminRole() {
    if (abkIsOrgAdmin !== null || abkRoleCheckScheduled) return;
    abkRoleCheckScheduled = true;
    setTimeout(function() {
      abkIsOrgAdmin = false; // hidden by default until proven otherwise
      fetch('/api/auth/refresh', { method: 'POST', credentials: 'include' })
        .then(function(r) { return r.ok ? r.json() : null; })
        .then(function(d) {
          abkIsOrgAdmin = !!(d && d.user && d.user.role === 'ADMIN');
          if (abkIsOrgAdmin) { moveOrgSettingsToRail(); applyAdminOnlyVisibility(); }
          var lcEmail = d && d.user && d.user.email;
          abkIsPrimaryOwner = !!(lcEmail && lcEmail.toLowerCase() === ABK_PRIMARY_OWNER_EMAIL);
          if (abkIsPrimaryOwner) applyOwnerOnlyVisibility();
        })
        .catch(function() {});
    }, 4000);
  }
  // "Agent Builder" and "Skills" are hidden by default via CSS
  // ([data-abk-admin-ok] gate, see the injected <style>) until we know the
  // signed-in user is an org ADMIN/OWNER — same reasoning as the
  // Organization Settings rail icon above: a plain USER should never see
  // these entry points at all.
  function applyAdminOnlyVisibility() {
    if (abkIsOrgAdmin !== true) return;
    document.querySelectorAll('button[aria-label="Agent Builder"], button[aria-label="Skills"]').forEach(function(b) {
      b.setAttribute('data-abk-admin-ok', '1');
    });
  }

  // Agent sharing ("Share {agent name}" — the label is dynamic, hence the
  // prefix match) is restricted to the primary owner only: they build the
  // canonical ABK Agent and distribute it to everyone else, so other org
  // ADMINs shouldn't be able to spin off their own shared copies. Hidden by
  // default via CSS ([data-abk-owner-ok] gate) until confirmed.
  function applyOwnerOnlyVisibility() {
    if (!abkIsPrimaryOwner) return;
    document.querySelectorAll('button[aria-label^="Share "]').forEach(function(b) {
      b.setAttribute('data-abk-owner-ok', '1');
    });
  }

  // Nobody — not even an org ADMIN/OWNER — should be able to revoke/delete
  // the pre-configured "Agent365 Bridge" MCP connector from the sidebar;
  // it's shared org infrastructure, not a personal connector. Only hide the
  // destructive action on THAT specific row (identified by its own text),
  // so users can still add/remove their own custom MCP servers normally.
  function hideAgent365BridgeRevoke() {
    document.querySelectorAll('[aria-label="Revoke"]').forEach(function(el) {
      if (el.dataset.abkChecked) return;
      var cur = el, isBridgeRow = false;
      for (var d = 0; d < 6 && cur; d++) {
        if (cur.textContent && cur.textContent.indexOf('Agent365 Bridge') !== -1) { isBridgeRow = true; break; }
        cur = cur.parentElement;
      }
      if (isBridgeRow) {
        el.dataset.abkChecked = '1';
        el.style.display = 'none';
      }
    });
  }

  function moveOrgSettingsToRail() {
    checkOrgAdminRole();
    if (!abkIsOrgAdmin) return;
    if (document.querySelector('[data-abk-rail-org]')) return;

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

    var newBtn = document.createElement(template.tagName);
    newBtn.className = template.className;
    newBtn.setAttribute('data-abk-rail-org', '1');
    newBtn.setAttribute('title', 'Organization Settings');
    newBtn.setAttribute('aria-label', 'Organization Settings');
    newBtn.innerHTML = ORG_RAIL_ICON_SVG;
    newBtn.addEventListener('click', function(e) {
      e.preventDefault();
      e.stopPropagation();
      openOrgAdmin();
    });
    rail.appendChild(newBtn);
  }

  function isDarkMode() {
    try { return document.documentElement.classList.contains('dark'); } catch(ex) { return false; }
  }

  function openOrgAdmin() {
    var cachedToken = null;
    try { cachedToken = sessionStorage.getItem(ORG_SESSION_KEY); } catch(ex) {}
    var emailHint = null;
    try { emailHint = localStorage.getItem(ORG_EMAIL_HINT_KEY); } catch(ex) {}

    var params = ['theme=' + (isDarkMode() ? 'dark' : 'light')];
    if (cachedToken) params.push('org_token=' + encodeURIComponent(cachedToken));
    else if (emailHint) params.push('prefill_email=' + encodeURIComponent(emailHint));
    showOrgAdminModal(params.join('&'));
  }

  // Keep the Organization Settings modal (chrome + iframe content) in sync
  // if the user toggles light/dark elsewhere in LibreChat while it's open.
  var abkThemeObserverInstalled = false;
  function installThemeSync() {
    if (abkThemeObserverInstalled) return;
    abkThemeObserverInstalled = true;
    new MutationObserver(function() {
      var iframe = document.getElementById('abk-org-iframe');
      if (!iframe) return;
      var dark = isDarkMode();
      try { iframe.contentWindow.postMessage({ type: 'abk_theme', dark: dark }, BRIDGE_URL); } catch(ex) {}
      var modal = document.getElementById('abk-org-modal');
      var header = document.getElementById('abk-org-header');
      var title = document.getElementById('abk-org-title');
      var closeBtn = document.getElementById('abk-org-close');
      if (modal) modal.style.background = dark ? '#101317' : '#f9f9f9';
      if (header) { header.style.background = dark ? '#15181d' : '#fff'; header.style.borderBottom = '1px solid ' + (dark ? '#282e35' : '#e8e8e8'); }
      if (title) title.style.color = dark ? '#2e97e0' : '#0071BC';
      if (closeBtn) closeBtn.style.color = dark ? '#9aa5b1' : '#888';
    }).observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });
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
    for (var ci = 0; ci < candidates.length; ci++) {
      var el = candidates[ci];
      if (el.dataset.abkGreetDone) continue;
      // Note: the rendered text uses U+00A0 (non-breaking space), not a
      // regular space, between words — \\s matches both (double backslash
      // because this whole script is itself inside patch.js's own outer
      // template literal, which eats a single backslash as its own escape).
      if (!/Good\\s(morning|afternoon|evening)/i.test(el.textContent || '')) continue;
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

      // IMPORTANT: never reparent \`el\` — it's a live React-owned node, and
      // React keeps its own reference to el's ORIGINAL parent. An earlier
      // version of this function wrapped el in a new "stack" div via
      // stack.appendChild(el), which moved it to a different parent than
      // the one React rendered it under. The next time React tried to
      // update or remove that node (e.g. leaving the greeting screen once a
      // chat starts), it called removeChild on the original parent — which
      // no longer contained it — crashing the whole app with "Failed to
      // execute 'removeChild' on 'Node': The node to be removed is not a
      // child of this node." Get the same "logo + bar beside the heading,
      // subtitle below" layout by only ever adding new SIBLING nodes around
      // el, using a flex-wrap + flex-basis trick to push the subtitle onto
      // its own line — el's parent and position in the tree are never
      // touched.
      // LibreChat can have more than one of these greeting nodes alive in
      // the DOM at the same time (e.g. a previous New Chat screen that
      // hasn't been torn down yet) — el.dataset.abkGreetDone only stops us
      // from processing the SAME node twice, not from processing several
      // DISTINCT nodes that all match. Without a single global guard here,
      // each one independently got its own logo/bar/subtitle, which is what
      // produced several duplicate "Πώς μπορώ..." lines stacked on screen.
      // Only the first matching instance gets the branding treatment; the
      // SVG/avatar hiding above still applies to every instance since that
      // part is just a style toggle, not an added/duplicated element.
      if (document.querySelector('[data-abk-greet-sub]')) continue;

      // Force the SAME centered layout regardless of which screen this is
      // (the plain "Good afternoon, {name}" home greeting and a saved
      // Agent's own name/description landing reuse this same heading node,
      // but sit inside different ambient containers/CSS — one already
      // happened to render centered, the other left-aligned. Setting
      // justify-content + width explicitly here, instead of relying on
      // whatever the surrounding LibreChat layout does, makes the row look
      // identical in both places.)
      var insertParent = el.parentElement || container;
      insertParent.style.display = 'flex';
      insertParent.style.flexWrap = 'wrap';
      insertParent.style.alignItems = 'center';
      insertParent.style.justifyContent = 'center';
      insertParent.style.width = '100%';

      // Insert ABK logo + vertical accent bar as siblings right before el
      var img = document.createElement('img');
      img.src = '/assets/abk-logo.png?abk=1';
      img.setAttribute('data-abk-greet-logo', '1');
      img.style.cssText = 'height:26px;width:auto;object-fit:contain;vertical-align:middle;margin-right:14px;display:inline-block;order:-2;';
      insertParent.insertBefore(img, el);

      var bar = document.createElement('span');
      bar.setAttribute('data-abk-greet-bar', '1');
      bar.style.cssText = 'display:inline-block;width:2px;height:42px;background:var(--ring-primary);margin-right:14px;vertical-align:middle;order:-1;';
      insertParent.insertBefore(bar, el);

      // Primary-colored subtitle below the greeting heading, centered under
      // the whole row (not indented under just the heading text) so it
      // lines up the same way whether the row above it is short ("ABK
      // Agent") or long ("Good afternoon, Stavros Nikolaou"). flex-basis:
      // 100% forces it onto its own line within insertParent's wrapping
      // flex row; order:99 keeps it after el regardless of insertion order.
      var sub = document.createElement('p');
      sub.setAttribute('data-abk-greet-sub', '1');
      sub.textContent = 'Πώς μπορώ να βοηθήσω σήμερα;';
      sub.style.cssText = 'margin:4px 0 0;font-size:14px;font-weight:500;flex-basis:100%;order:99;text-align:center;';
      insertParent.appendChild(sub);
      break;
    }
    } catch(e) {}
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
    try { hideOrDivider(); } catch(e) {}
    try { hideSignUp(); } catch(e) {}
    try { hideNativeAdminLink(); } catch(e) {}
    try { moveOrgSettingsToRail(); } catch(e) {}
    try { applyAdminOnlyVisibility(); } catch(e) {}
    try { applyOwnerOnlyVisibility(); } catch(e) {}
    try { hideAgent365BridgeRevoke(); } catch(e) {}
    try { injectConnectorsTab(); } catch(e) {}
    try { hideConnectorsPanelIfClickedElsewhere(); } catch(e) {}
    try { patchGreeting(); } catch(e) {}
  }

  function setup() {
    installThemeSync();
    checkOrgAdminRole();
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

// ── Force Microsoft account picker on the LibreChat OpenID login ────────────
// LibreChat's OpenID strategy has no env var to control the authorization
// `prompt` parameter, so without this patch, Microsoft silently reuses
// whatever AAD session cookie is already in the browser instead of showing
// "Pick an account" — confusing on shared/multi-account machines.
// authorizationRequestParams() in api/strategies/openidStrategy.js is the
// single place both the main login and admin-OIDC strategies build their
// authorize request, so patching it here covers both.
{
  const stratPath = '/app/api/strategies/openidStrategy.js';
  let strat = fs.readFileSync(stratPath, 'utf8');
  const marker = "params.set('prompt', 'select_account'); // ABK: force account picker";
  if (strat.includes(marker)) {
    console.log('  openidStrategy.js already patched for account picker, skipping.');
  } else {
    const anchor = 'return params;';
    const occurrences = strat.split(anchor).length - 1;
    if (occurrences !== 1) {
      throw new Error(
        `openidStrategy.js: expected exactly 1 occurrence of "${anchor}", found ${occurrences}. ` +
        'LibreChat likely changed this file — update the account-picker patch in patch.js.'
      );
    }
    strat = strat.replace(anchor, `${marker}\n    ${anchor}`);
    fs.writeFileSync(stratPath, strat);
    console.log('  openidStrategy.js patched: Microsoft account picker forced on login.');
  }
}

// ── Force fresh OAuth on every Agent365 Bridge reconnect ────────────────────
// The native "Connect" and "Reconnect" MCP card buttons both call the same
// POST /api/mcp/:serverName/reinitialize endpoint (MCPCardActions.tsx —
// both wire to onInitialize()). That endpoint silently reuses an existing,
// still-valid stored OAuth token when one is present, so clicking
// "Reconnect" on an already-connected server does NOT show a fresh
// Microsoft SSO prompt — it only appears the first time (or after a real
// disconnect/expiry). For the shared "Agent365 Bridge" connector we always
// want an explicit, visible SSO confirmation on every reconnect, so a green
// "Connected" dot is always backed by the account the user just picked.
// Clearing the stored token first — using the exact same mechanism the
// native "Revoke" button calls (MCPTokenStorage.deleteUserTokens +
// clearing cached OAuth flow state; see UserController.js
// clearStoredMCPOAuthState) — makes reinitMCPServer() correctly report
// oauthRequired: true, so the normal OAuth flow already handled by this
// route runs every time. Scoped to serverName === 'agent365-bridge' only;
// every other MCP server (including ones users add themselves) is
// untouched.
{
  const routesMcpPath = '/app/api/server/routes/mcp.js';
  let routesMcp = fs.readFileSync(routesMcpPath, 'utf8');
  const marker = '// ABK: force fresh OAuth for agent365-bridge on every reconnect';
  if (routesMcp.includes(marker)) {
    console.log('  routes/mcp.js already patched for forced re-auth, skipping.');
  } else {
    const anchor =
      '      await mcpManager.disconnectUserConnection(user.id, serverName);\n' +
      '      logger.info(\n' +
      '        `[MCP Reinitialize] Disconnected existing user connection for server: ${serverName}`,\n' +
      '      );\n';
    const occurrences = routesMcp.split(anchor).length - 1;
    if (occurrences !== 1) {
      throw new Error(
        `routes/mcp.js: expected exactly 1 occurrence of the reinitialize disconnect anchor, found ${occurrences}. ` +
        'LibreChat likely changed this file — update the forced re-auth patch in patch.js.'
      );
    }
    const injection = anchor +
      '\n' +
      '      ' + marker + '\n' +
      "      if (serverName === 'agent365-bridge') {\n" +
      '        try {\n' +
      '          await MCPTokenStorage.deleteUserTokens({\n' +
      '            userId: user.id,\n' +
      '            serverName,\n' +
      '            deleteToken: async (filter) => {\n' +
      '              await db.deleteTokens(filter);\n' +
      '            },\n' +
      '          });\n' +
      '          const abkFlowsCache = getLogStores(CacheKeys.FLOWS);\n' +
      '          const abkFlowManager = getFlowStateManager(abkFlowsCache);\n' +
      '          const abkFlowId = MCPOAuthHandler.generateFlowId(user.id, serverName);\n' +
      '          await Promise.allSettled([\n' +
      "            abkFlowManager.deleteFlow(abkFlowId, 'mcp_get_tokens'),\n" +
      "            abkFlowManager.deleteFlow(abkFlowId, 'mcp_oauth'),\n" +
      '          ]);\n' +
      '        } catch (abkError) {\n' +
      '          logger.warn(\n' +
      '            `[MCP Reinitialize] ABK: failed to clear stored OAuth state for ${serverName}:`,\n' +
      '            abkError,\n' +
      '          );\n' +
      '        }\n' +
      '      }\n';
    routesMcp = routesMcp.replace(anchor, injection);
    fs.writeFileSync(routesMcpPath, routesMcp);
    console.log('  routes/mcp.js patched: forced fresh OAuth for agent365-bridge on every reconnect.');
  }
}
