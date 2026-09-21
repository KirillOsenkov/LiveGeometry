import { dotnet } from './_framework/dotnet.js'

const is_browser = typeof window != "undefined";
if (!is_browser) throw new Error(`Expected to be running in a browser`);

const dotnetRuntime = await dotnet
    .withDiagnosticTracing(false)
    .withApplicationArgumentsFromQuery()
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
