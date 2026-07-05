/**
 * Divya Dental Clinic - Bilingual Core Framework Engine
 * Controller Pattern managing real-time standard dictionary injections without refetches
 */

const i18nData = { en: null, hi: null };
let currentLang = localStorage.getItem('clinic_lang') || 'en';

window.addEventListener('DOMContentLoaded', async () => {
    // 1. Concurrent translation matrix initialization 
    await Promise.all([
        loadLanguagePack('en', 'en.json'),
        loadLanguagePack('hi', 'hi.json')
    ]);
    
    // 2. Initial execution pass to paint localized state
    switchLanguage(currentLang);
    
    // 3. Fallback handler loops for native media missing paths
    initImageFallbackHandlers();

    // 4. Track scroll transformations on navigation header
    const navbar = document.querySelector('.navbar-premium');
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    });
});

async function loadLanguagePack(lang, src) {
    try {
        const res = await fetch(src);
        if (!res.ok) throw new Error(`HTTP Matrix error fetching definitions: ${res.status}`);
        i18nData[lang] = await res.json();
    } catch (err) {
        console.error(`Language asset failure:`, err);
        // Fallback fallback definitions injected safely if missing filesystem privileges
        i18nData[lang] = {}; 
    }
}

function switchLanguage(lang) {
    if (!i18nData[lang]) return;
    currentLang = lang;
    localStorage.setItem('clinic_lang', lang);

    // Hydrate targets containing translation tags
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const keyPath = el.getAttribute('data-i18n');
        const translatedString = getValueByPath(i18nData[lang], keyPath);
        if (translatedString) {
            if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
                el.placeholder = translatedString;
            } else {
                el.innerHTML = translatedString;
            }
        }
    });

    // Update Meta Title and Tags safely
    const localizedTitle = getValueByPath(i18nData[lang], 'meta.title');
    const localizedDesc = getValueByPath(i18nData[lang], 'meta.description');
    if (localizedTitle) document.title = localizedTitle;
    if (localizedDesc) document.getElementById('meta-desc').setAttribute('content', localizedDesc);

    // Update language toggle visually
    document.querySelectorAll('.lang-switcher .btn').forEach(btn => btn.classList.remove('active'));
    const targetButton = document.getElementById(`lang-${lang}`);
    if (targetButton) targetButton.classList.add('active');
    
    // Ensure text read direction is correctly formatted
    document.documentElement.lang = lang;
}

function getValueByPath(obj, path) {
    return path.split('.').reduce((acc, part) => acc && acc[part], obj);
}

function initImageFallbackHandlers() {
    document.querySelectorAll('img').forEach(img => {
        img.addEventListener('error', function() {
            const fallback = this.getAttribute('fallback-src');
            if (fallback && this.src !== fallback) {
                this.src = fallback;
            }
        });
    });
}