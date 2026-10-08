// The player's one global, LiveGeometry: play(element, options) makes a Player, and on page
// load every element with the class "livegeometry" is played - from the <script
// type="text/x-livegeometry"> inside it holding the drawing's XML, or from the .lgf at its
// data-src. data-theme, data-font, data-wheel and data-fit are the options.

const LiveGeometry = {
    version: "1",

    /** The players on the page, by element */
    players: new Map(),

    play(element, options = {}) {
        const existing = LiveGeometry.players.get(element);
        if (existing != null) {
            existing.dispose();
        }

        const player = new Player(element, options);
        LiveGeometry.players.set(element, player);
        return player;
    },

    /** Every .livegeometry element on the page that isn't played yet */
    playAll(root = document) {
        for (const element of root.querySelectorAll(".livegeometry")) {
            if (LiveGeometry.players.has(element)) {
                continue;
            }

            const options = {
                theme: element.dataset.theme,
                font: element.dataset.font,
                wheel: element.dataset.wheel,
                fit: element.dataset.fit
            };
            const inline = element.querySelector("script[type='text/x-livegeometry']");
            if (inline != null) {
                options.lgf = inline.textContent;
            } else if (element.dataset.src != null) {
                options.src = element.dataset.src;
            }

            // a drawing that can't be played says so in its place: the page's other embeds go on
            try {
                LiveGeometry.play(element, options);
            } catch (error) {
                console.error("Live Geometry: " + error.message, error);
                element.textContent = "Live Geometry: this drawing could not be played (" + error.message + ")";
            }
        }
    },

    // the classes, for a page that builds on the player
    Drawing,
    Player,
    FigureTypes,
    Point,
    GeometryMath
};

if (typeof window !== "undefined") {
    window.LiveGeometry = LiveGeometry;
    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", () => LiveGeometry.playAll());
    } else {
        LiveGeometry.playAll();
    }
}
