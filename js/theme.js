/**
 * Theme toggle (dark/light mode) and environment detection.
 */

/** Apply dark or light theme and persist preference. */
export function applyTheme(isDark) {
    const btnThemeToggle = document.getElementById('btnThemeToggle');
    document.body.classList.toggle('dark-mode', isDark);
    if (btnThemeToggle) {
        btnThemeToggle.innerHTML = isDark ? '☀️ Light' : '🌙 Dark';
        btnThemeToggle.title = isDark ? 'Chuyển sang chế độ Sáng (Light Mode)' : 'Chuyển sang chế độ Tối (Dark Mode)';
    }
    try { localStorage.setItem('scanVidTheme', isDark ? 'dark' : 'light'); } catch (_) {}
}

/** Load saved theme from localStorage, or fall back to OS preference. */
export function loadSavedTheme() {
    const savedTheme = (() => {
        try { return localStorage.getItem('scanVidTheme'); } catch (_) { return null; }
    })();

    if (savedTheme === 'dark') {
        applyTheme(true);
    } else if (!savedTheme && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        applyTheme(true);
    }
}

/** Detect environment from hostname and show badge (PREVIEW / LOCAL / hidden for production). */
export function detectEnv() {
    const badge = document.getElementById('envBadge');
    if (!badge) return;
    const host = window.location.hostname;

    if (host.includes('-git-preview') || host.includes('preview')) {
        badge.style.display = 'inline';
        badge.style.background = '#f39c12';
        badge.style.color = '#fff';
        badge.textContent = '🧪 PREVIEW';
    } else if (host === 'localhost' || host === '127.0.0.1' || host === '') {
        badge.style.display = 'inline';
        badge.style.background = '#3498db';
        badge.style.color = '#fff';
        badge.textContent = '💻 LOCAL';
    }
    // Production: badge stays hidden
}
