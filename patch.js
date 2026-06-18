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
</style>
<script>
(function() {
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
  function openOrgAdmin() {
    var existing = document.getElementById('abk-org-overlay');
    if (existing) { existing.remove(); return; }

    var lct = localStorage.getItem('token') || '';
    var url = 'https://agent365-bridge.lemonsea-0ef310bc.swedencentral.azurecontainerapps.io/org-admin?embed=1' +
              (lct ? '&lc_token=' + encodeURIComponent(lct) : '');

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

  function patchAdminLink() {
    var all = document.querySelectorAll('a, button, [role="menuitem"]');
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      if (!el.dataset.abkAdminPatched && (el.innerText || '').trim() === 'Admin Settings') {
        el.dataset.abkAdminPatched = '1';
        el.addEventListener('click', function(e) {
          e.preventDefault();
          e.stopPropagation();
          openOrgAdmin();
        }, true);
      }
    }
  }

  function tryPatch() {
    patchLoginText();
    patchLogo();
    patchAdminLink();
  }

  // Start observing immediately on <html> — do NOT wait for DOMContentLoaded
  // This ensures we catch React's first render
  new MutationObserver(tryPatch).observe(document.documentElement, { childList: true, subtree: true });
  tryPatch();
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
