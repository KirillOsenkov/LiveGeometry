const is_browser = typeof window != "undefined";
if (!is_browser) throw new Error(`Expected to be running in a browser`);

// "?splash": the splash screen alone, for working on it (index.html, app.css); the app
// is not started.
if (globalThis.location.search.includes("splash")) {
    throw new Error("splash only");
}

const { dotnet } = await import('./_framework/dotnet.js');

// The bar under the splash: downloads finished, out of downloads begun. The loader asks
// for every resource right after it has the config, so the denominator settles early.
const progressBar = globalThis.document.getElementById('splash-progress-bar');
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
        if (progressBar) {
            progressBar.style.width = Math.round(100 * downloadsDone / downloadsBegun) + '%';
        }
    }, () => { });
    return response;
}

const dotnetRuntime = await dotnet
    .withDiagnosticTracing(false)
    .withApplicationArgumentsFromQuery()
    .withResourceLoader(loadResource)
    .create();

// The address bar, for BrowserAddressBar.cs: the app has routes (/gallery/morley).
dotnetRuntime.setModuleImports('main.js', {
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

await dotnetRuntime.runMain(config.mainAssemblyName, [globalThis.location.href]);
