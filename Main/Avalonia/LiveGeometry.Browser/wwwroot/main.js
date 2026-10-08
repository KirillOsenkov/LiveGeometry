const is_browser = typeof window != "undefined";
if (!is_browser) throw new Error(`Expected to be running in a browser`);

// "?splash": the splash screen alone, for working on it (index.html, app.css); the app
// is not started.
if (globalThis.location.search.includes("splash")) {
    throw new Error("splash only");
}

// The emoji font, for Program.OpenEmojiFont. The gallery's first rows show emoji, so on the
// gallery's page it is fetched now, while the runtime downloads (at a low priority, after
// the runtime's files). The browser reads a response on this thread a piece at a time, and
// asked for once the gallery was up, the font came in between the tiles' loads, seconds
// after its bytes. Elsewhere it is fetched when the app asks for it.
async function downloadEmojiFont(options) {
    const response = await fetch('fonts/Twemoji.Mozilla.ttf', options);
    if (!response.ok) {
        throw new Error('emoji font: ' + response.status);
    }

    return new Uint8Array(await response.arrayBuffer());
}

let emojiFontDownload = globalThis.location.pathname === '/' ? downloadEmojiFont({ priority: 'low' }) : null;
// a failure is the app's to report, when it asks for the font
emojiFontDownload?.catch(() => { });
let emojiFont = null;

// The splash (index.html) stays up until the app has drawn what the page opened at
// (SplashScreen.cs: the gallery with the tiles in view loaded, a drawing laid out). Its bar
// is the downloads for the first part and what the app reports for the rest. Avalonia adds
// splash-close at its first frame; that hides nothing here (app.css), but starts a deadline
// in case the app never says it is ready: better tiles that pop in than a splash that stays.
const splash = globalThis.document.querySelector('.avalonia-splash');
const progressCover = globalThis.document.getElementById('splash-progress-cover');
const downloadShare = 0.6;
function showProgress(fraction) {
    if (progressCover) {
        progressCover.style.transform = 'translateX(' + Math.round(100 * Math.min(1, Math.max(0, fraction))) + '%)';
    }
}

function hideSplash() {
    splash?.classList.add('app-ready');
}

if (splash) {
    new MutationObserver((mutations, observer) => {
        if (splash.classList.contains('splash-close')) {
            observer.disconnect();
            globalThis.setTimeout(hideSplash, 10000);
        }
    }).observe(splash, { attributes: true, attributeFilter: ['class'] });
}

const { dotnet } = await import('./_framework/dotnet.js');

// The downloads' part of the bar: downloads finished, out of downloads begun. The loader
// asks for every resource right after it has the config, so the denominator settles early.
let downloadsBegun = 0;
let downloadsDone = 0;
function loadResource(type, name, uri) {
    // the runtime's own JavaScript modules must be imported by url, not handed over as
    // a Response (with one, the app never came up); the same for the config it reads
    // first. Only the assemblies, the wasm and the ICU data are fetched and counted.
    if (type !== "assembly" && type !== "pdb" && type !== "dotnetwasm" && type !== "globalization") {
        return undefined;
    }

    downloadsBegun++;
    const response = fetch(uri);
    response.then(() => {
        downloadsDone++;
        showProgress(downloadShare * downloadsDone / downloadsBegun);
    }, () => { });
    return response;
}

const dotnetRuntime = await dotnet
    .withDiagnosticTracing(false)
    .withApplicationArgumentsFromQuery()
    .withResourceLoader(loadResource)
    .create();

// The address bar, for BrowserAddressBar.cs: the app has routes (/gallery/morley); the
// settings, for BrowserSettingsStore.cs: local storage, one entry per key (index.html reads
// the theme from the same entry before the app is up).
const settingPrefix = 'LiveGeometry.';
dotnetRuntime.setModuleImports('main.js', {
    // the emoji font (above): its length once all of it is here, then the bytes in one piece
    fetchEmojiFont: async () => {
        emojiFontDownload ??= downloadEmojiFont();
        emojiFont = await emojiFontDownload;
        return emojiFont.length;
    },
    copyEmojiFont: (target) => {
        target.set(emojiFont);
        emojiFont = null;
        emojiFontDownload = null;
    },
    // a text file of the site, for PlayerEmbed (the player's script, written into a saved page)
    fetchText: async (url) => {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error(url + ': ' + response.status);
        }

        return await response.text();
    },
    // the splash (above): the app's part of the bar, and the app on screen
    reportProgress: (fraction) => showProgress(downloadShare + (1 - downloadShare) * fraction),
    hideSplash: () => hideSplash(),
    getSetting: (key) => {
        try {
            return globalThis.localStorage.getItem(settingPrefix + key);
        } catch {
            return null;
        }
    },
    // a Mac's keyboard (an iPad's too: a browser on one says MacIntel), for the names of keys in hints
    isMac: () => {
        const platform = globalThis.navigator.userAgentData?.platform || globalThis.navigator.platform || '';
        return /Mac|iPhone|iPad|iPod/i.test(platform);
    },
    getSettingKeys: (prefix) => {
        const keys = [];
        try {
            const storage = globalThis.localStorage;
            for (let i = 0; i < storage.length; i++) {
                const key = storage.key(i);
                if (key !== null && key.startsWith(settingPrefix + prefix)) {
                    keys.push(key.substring(settingPrefix.length));
                }
            }
        } catch {
            // storage denied: nothing kept
        }

        return keys;
    },
    setSetting: (key, value) => {
        try {
            if (value === null || value === undefined) {
                globalThis.localStorage.removeItem(settingPrefix + key);
            } else {
                globalThis.localStorage.setItem(settingPrefix + key, value);
            }
        } catch {
            // private mode, storage denied: the choice holds for the session
        }
    },
    getPath: () => globalThis.location.pathname,
    pushState: (path, title) => {
        if (globalThis.location.pathname !== path) {
            globalThis.history.pushState(null, '', path);
        }
        globalThis.document.title = title;
    },
    replaceState: (path, title) => {
        globalThis.history.replaceState(null, '', path);
        globalThis.document.title = title;
    }
});

const config = dotnetRuntime.getConfig();
const exports = await dotnetRuntime.getAssemblyExports(config.mainAssemblyName);
globalThis.addEventListener('popstate', () => {
    exports.LiveGeometry.Browser.BrowserAddressBar.OnPopState(globalThis.location.pathname);
});

// What waits to be stored (the drawing, tweaked theme colors) is stored when the page is
// hidden: a tab switched away from, closed, or a phone's browser sent to the background,
// after which there may be no other chance.
globalThis.document.addEventListener('visibilitychange', () => {
    if (globalThis.document.visibilityState === 'hidden') {
        exports.LiveGeometry.Browser.BrowserSettingsStore.OnPageHidden();
    }
});

await dotnetRuntime.runMain(config.mainAssemblyName, [globalThis.location.href]);

// Main returns once Avalonia is set up, before its first frame. An app that never gets to
// one (a CRASH line in the console) leaves a splash that says "Loading" for good otherwise.
globalThis.setTimeout(hideSplash, 20000);
