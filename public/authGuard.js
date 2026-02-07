/**
 * Auth Guard — Session-based password protection.
 * Blocks the app UI with a lock screen until the correct access code is entered.
 * The code is stored as a reversed base64 string to avoid plain-text visibility.
 */

(function () {
    var TOKEN_KEY = 'fw_access_granted';

    // Obfuscated password: base64-encoded access code
    var _k = 'QGFzaWF0aXMtMjAyNiE=';
    function _d(s) {
        return atob(s);
    }

    // Check if already authenticated
    if (sessionStorage.getItem(TOKEN_KEY) === '1') return;

    // Block scrolling on the body
    document.body.style.overflow = 'hidden';

    // Create lock screen overlay
    var overlay = document.createElement('div');
    overlay.id = 'authOverlay';
    overlay.innerHTML = '<div class="auth-card">' +
        '<div class="auth-logo">' +
        '<svg width="48" height="48" viewBox="0 0 32 32" fill="none">' +
        '<rect x="3" y="4" width="18" height="24" rx="2" stroke="currentColor" stroke-width="2"/>' +
        '<rect x="11" y="4" width="18" height="24" rx="2" fill="rgba(99,102,241,0.15)" stroke="currentColor" stroke-width="2"/>' +
        '<path d="M15 11h10M15 15h10M15 19h7" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>' +
        '</svg>' +
        '</div>' +
        '<h1>Format<span style="background:linear-gradient(135deg,#6366f1,#06b6d4);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;">Word</span></h1>' +
        '<p class="auth-subtitle">Acc\u00e8s restreint \u2014 Veuillez entrer le code d\u2019acc\u00e8s</p>' +
        '<div class="auth-input-group">' +
        '<input type="password" id="authInput" placeholder="Code d\u2019acc\u00e8s" autocomplete="off" spellcheck="false" />' +
        '<button id="authSubmit">\u2192</button>' +
        '</div>' +
        '<p class="auth-error" id="authError"></p>' +
        '</div>';

    // Inject styles
    var style = document.createElement('style');
    style.textContent =
        '#authOverlay{position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:#0a0e1a;background-image:radial-gradient(ellipse 80% 50% at 50% -20%,rgba(99,102,241,.12),transparent),radial-gradient(ellipse 60% 40% at 80% 100%,rgba(6,182,212,.08),transparent);animation:authFadeIn .4s ease}' +
        '@keyframes authFadeIn{from{opacity:0}to{opacity:1}}' +
        '.auth-card{background:rgba(26,31,53,.85);border:1px solid rgba(99,102,241,.2);border-radius:20px;padding:2.5rem 2.5rem 2rem;text-align:center;max-width:400px;width:90%;backdrop-filter:blur(16px);box-shadow:0 8px 40px rgba(0,0,0,.5),0 0 60px rgba(99,102,241,.08);animation:authSlideUp .5s cubic-bezier(.4,0,.2,1)}' +
        '@keyframes authSlideUp{from{transform:translateY(30px);opacity:0}to{transform:translateY(0);opacity:1}}' +
        '.auth-logo{color:#6366f1;margin-bottom:1rem}' +
        '.auth-card h1{font-family:Inter,sans-serif;font-size:1.6rem;font-weight:700;color:#f1f5f9;margin-bottom:.4rem;letter-spacing:-.02em}' +
        '.auth-subtitle{font-family:Inter,sans-serif;font-size:.85rem;color:#64748b;margin-bottom:1.5rem}' +
        '.auth-input-group{display:flex;gap:.5rem}' +
        '#authInput{flex:1;padding:.75rem 1rem;border-radius:10px;border:1px solid rgba(99,102,241,.2);background:#0d1225;color:#f1f5f9;font-family:Inter,sans-serif;font-size:.9rem;outline:none;transition:border-color .25s,box-shadow .25s}' +
        '#authInput:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.15)}' +
        '#authInput::placeholder{color:#475569}' +
        '#authSubmit{padding:.75rem 1.25rem;border-radius:10px;border:none;background:linear-gradient(135deg,#6366f1,#06b6d4);color:#fff;font-size:1.1rem;font-weight:600;cursor:pointer;transition:transform .2s,box-shadow .2s;box-shadow:0 4px 15px rgba(99,102,241,.25)}' +
        '#authSubmit:hover{transform:translateY(-1px);box-shadow:0 6px 25px rgba(99,102,241,.35)}' +
        '.auth-error{font-family:Inter,sans-serif;font-size:.8rem;color:#ef4444;margin-top:.75rem;min-height:1.2em}' +
        '@keyframes authShake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}';

    document.head.appendChild(style);
    document.body.appendChild(overlay);

    var input = document.getElementById('authInput');
    var btn = document.getElementById('authSubmit');
    var errEl = document.getElementById('authError');

    function attempt() {
        if (input.value === _d(_k)) {
            sessionStorage.setItem(TOKEN_KEY, '1');
            overlay.style.transition = 'opacity 0.3s';
            overlay.style.opacity = '0';
            setTimeout(function () {
                overlay.remove();
                style.remove();
                document.body.style.overflow = '';
            }, 300);
        } else {
            errEl.textContent = 'Code incorrect';
            input.value = '';
            input.focus();
            var card = overlay.querySelector('.auth-card');
            card.style.animation = 'authShake 0.4s ease';
            setTimeout(function () { card.style.animation = ''; }, 400);
        }
    }

    btn.addEventListener('click', attempt);
    input.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') attempt();
        errEl.textContent = '';
    });

    setTimeout(function () { input.focus(); }, 100);
})();
