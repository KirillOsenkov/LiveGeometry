// The Catriona Agg library's script, on the JavaScript player (/player/1/player.js, loaded
// first): the gallery at /CatrionaAgg (index.html) and a puzzle at /CatrionaAgg/<slug>
// (puzzle.html, and the copy of it the build writes for each puzzle). Each page runs its own
// part; the catalog, the theme and the arrangement of the gallery are shared. The arrangement
// (grouping, search, order) is kept for the tab (sessionStorage), so that a puzzle's arrows
// step through the puzzles in the order the gallery showed them.

document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  var root = document.documentElement;
  var Base = '/CatrionaAgg';
  var Mathforge = 'https://notes.mathforge.org/notes/published/';
  var Icons = Base + '/icons.svg';
  var reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var isGallery = document.getElementById('gallery-view') != null;

  // ---------- the theme: the page, every drawing on it, and the site's own setting ----------

  function systemTheme() {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function currentTheme() {
    return root.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  }

  // the theme the site has stored, else the system's (as the head's script reads it)
  function storedTheme() {
    var theme = null;
    try { theme = localStorage.getItem('LiveGeometry.Theme'); } catch (e) { }
    return theme === 'Dark' ? 'dark' : theme === 'Light' ? 'light' : systemTheme();
  }

  function toggleTheme() {
    var theme = currentTheme() === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', theme);
    // as the app does: landing on what the system says stores no choice at all
    try {
      if (theme === systemTheme()) {
        localStorage.removeItem('LiveGeometry.Theme');
      } else {
        localStorage.setItem('LiveGeometry.Theme', theme === 'dark' ? 'Dark' : 'Light');
      }
    } catch (e) { }

    themeChanged();
  }

  function updateThemeButtons() {
    document.querySelectorAll('.icon-button.theme').forEach(function (button) {
      button.title = currentTheme() === 'dark' ? 'Light theme' : 'Dark theme';
    });
  }

  document.querySelectorAll('.icon-button.theme').forEach(function (button) {
    button.addEventListener('click', toggleTheme);
  });
  updateThemeButtons();

  function themeChanged() {
    updateThemeButtons();
    if (puzzlePlayer != null) {
      puzzlePlayer.applyTheme(currentTheme());
    }

    if (featurePlayer != null) {
      featurePlayer.applyTheme(currentTheme());
    }

    if (hover != null) {
      hover.player.applyTheme(currentTheme());
    }

    // the pictures of the tiles are drawn again in the new theme, those in view first
    Thumbs.restart();
  }

  // a page the browser kept and shows again (Back) may have missed a change of theme on another page
  window.addEventListener('pageshow', function (e) {
    if (e.persisted && storedTheme() !== currentTheme()) {
      root.setAttribute('data-theme', storedTheme());
      themeChanged();
    }
  });

  // ---------- the catalog ----------

  var everything = [];   // every puzzle of the catalog
  var puzzles = [];      // the ones with a drawing, newest first
  var bySlug = new Map();
  var texts = new Map(); // by slug: a promise of the drawing's text

  var Topics = ['Basic Angles', 'Angles in Parallel Lines', 'Shapes', 'Measures', 'Transformations and Symmetry', 'Circle Theorems', 'Trigonometry', 'Algebra', 'Other Techniques'];
  var Shapes = [
    ['Triangles', /triangle/],
    ['Squares', /square/],
    ['Rectangles', /rectangle/],
    ['Other quadrilaterals', /rhomb|parallelogram|kite|trapezi|quadrilateral|dart/],
    ['Pentagons', /pentagon/],
    ['Hexagons', /hexagon/],
    ['Other polygons', /octagon|dodecagon|polygon|heptagon|decagon/],
    ['Circles', /circle|circular|arc|lune|annulus|ring/]
  ];
  var Questions = ['Angle', 'Area', 'Length', 'Fraction', 'Other'];

  function shapesOf(puzzle) {
    var title = puzzle.title.toLowerCase();
    var found = Shapes.filter(function (shape) { return shape[1].test(title); }).map(function (shape) { return shape[0]; });
    return found.length > 0 ? found : ['Other shapes'];
  }

  function questionOf(puzzle) {
    var text = (puzzle.statement || '').toLowerCase();
    // the question is the last sentence that asks: "The blue square has area 2. What's the yellow area?"
    var asked = text.split(/(?<=[.?!])\s+/).filter(function (sentence) { return /\?\s*$/.test(sentence); }).pop() || text;
    if (/fraction|proportion|percentage/.test(asked)) { return 'Fraction'; }
    if (/angle/.test(asked)) { return 'Angle'; }
    if (/area/.test(asked)) { return 'Area'; }
    if (/length|diameter|radius|width|height|perimeter|distance|long|side|wide|tall/.test(asked)) { return 'Length'; }
    return 'Other';
  }

  function topicsOf(puzzle) {
    return puzzle.categories.length > 0 ? puzzle.categories : ['Not yet sorted'];
  }

  function parseDate(text) {
    var parts = text.split('-');
    return new Date(+parts[0], +parts[1] - 1, +parts[2]);
  }

  var Months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  function longDate(date) {
    return date.getDate() + ' ' + Months[date.getMonth()] + ' ' + date.getFullYear();
  }

  function shortDate(date) {
    return date.getDate() + ' ' + Months[date.getMonth()].substring(0, 3) + ' ' + date.getFullYear();
  }

  // a wiki page's address on Mathforge Notes: its name, spaces as +
  function mathforge(name) {
    return Mathforge + encodeURIComponent(name).replace(/%20/g, '+').replace(/'/g, '%27');
  }

  // The question as the catalog has it, with the TeX of the wiki: $6$, $x^2$, $\sqrt{2}$, $\frac{1}{2}$
  function plainQuestion(text) {
    return (text || '').replace(/\$([^$]*)\$/g, function (all, tex) {
      return tex
        .replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, '$1/$2')
        .replace(/\\sqrt\{([^}]*)\}/g, '√$1')
        .replace(/\\pi/g, 'π')
        .replace(/\\times/g, '×')
        .replace(/\\circ/g, '°')
        .replace(/\^\{?2\}?/g, '²')
        .replace(/\^\{?3\}?/g, '³')
        .replace(/[{}\\]/g, '');
    });
  }

  function fetchText(puzzle) {
    var promise = texts.get(puzzle.slug);
    if (promise == null) {
      promise = fetch(drawingAddress(puzzle)).then(function (response) {
        if (!response.ok) {
          throw new Error('could not load ' + puzzle.drawing + ' (' + response.status + ')');
        }

        return response.text();
      });
      texts.set(puzzle.slug, promise);
      promise.catch(function () { texts.delete(puzzle.slug); });
    }

    return promise;
  }

  function drawingAddress(puzzle) {
    return Base + '/drawings/' + puzzle.drawing;
  }

  function parseXml(text) {
    var parsed = new DOMParser().parseFromString(text.replace(/^﻿/, '').trimStart(), 'application/xml');
    if (parsed.querySelector('parsererror') != null) {
      throw new Error('not well-formed XML');
    }

    return parsed.documentElement;
  }

  // what a reader drags is left out of a picture: the tile shows her diagram as she drew it
  function withoutHandles(text) {
    var element = parseXml(text);
    element.querySelectorAll('[Style="Handle"]').forEach(function (figure) { figure.setAttribute('Visible', 'false'); });
    return element;
  }

  function hasHandles(text) {
    return /Style="Handle"/.test(text);
  }

  function isPlainClick(e) {
    return e.button === 0 && !e.ctrlKey && !e.metaKey && !e.shiftKey && !e.altKey;
  }

  // ---------- the arrangement of the gallery, kept for the tab ----------

  var DefaultState = { group: 'none', query: '', oldestFirst: false };
  var StateKey = 'CatrionaAgg.Arrangement';
  var ScrollKey = 'CatrionaAgg.GalleryScroll';
  var ReturnKey = 'CatrionaAgg.Returning';

  function readSession(key) {
    try { return sessionStorage.getItem(key); } catch (e) { return null; }
  }

  function writeSession(key, value) {
    try {
      if (value == null) {
        sessionStorage.removeItem(key);
      } else {
        sessionStorage.setItem(key, value);
      }
    } catch (e) { }
  }

  function validState(state) {
    return {
      group: ['none', 'year', 'topic', 'shape', 'question'].indexOf(state.group) >= 0 ? state.group : 'none',
      query: typeof state.query === 'string' ? state.query : '',
      oldestFirst: state.oldestFirst === true
    };
  }

  function storedState() {
    try {
      var text = readSession(StateKey);
      return text != null ? validState(JSON.parse(text)) : null;
    } catch (e) {
      return null;
    }
  }

  function storeState(state) {
    writeSession(StateKey, JSON.stringify(state));
  }

  function matches(puzzle, words) {
    return words.every(function (word) { return puzzle.searchText.indexOf(word) >= 0; });
  }

  // the gallery's sections for a state and what is in each, in the order they show
  function arrangement(state) {
    var words = state.query.toLowerCase().split(/\s+/).filter(function (word) { return word !== ''; });
    var shown = puzzles.filter(function (puzzle) { return matches(puzzle, words); });
    if (state.oldestFirst) {
      shown = shown.slice().reverse();
    }

    var sections = [];
    var add = function (name, list) {
      if (list.length > 0) {
        sections.push({ name: name, puzzles: list });
      }
    };
    if (state.group === 'none') {
      add(null, shown);
    } else if (state.group === 'year') {
      var years = Array.from(new Set(shown.map(function (puzzle) { return puzzle.year; })));
      years.forEach(function (year) { add(String(year), shown.filter(function (puzzle) { return puzzle.year === year; })); });
    } else if (state.group === 'topic') {
      Topics.concat(['Not yet sorted']).forEach(function (topic) { add(topic, shown.filter(function (puzzle) { return topicsOf(puzzle).indexOf(topic) >= 0; })); });
    } else if (state.group === 'shape') {
      Shapes.map(function (shape) { return shape[0]; }).concat(['Other shapes']).forEach(function (shape) {
        add(shape, shown.filter(function (puzzle) { return puzzle.shapes.indexOf(shape) >= 0; }));
      });
    } else if (state.group === 'question') {
      Questions.forEach(function (question) { add(question === 'Other' ? 'Something else' : 'What’s the ' + question.toLowerCase() + '?', shown.filter(function (puzzle) { return puzzle.question === question; })); });
    }

    return { shown: shown, sections: sections };
  }

  // the puzzles as the gallery shows them for a state, each once (a puzzle may be in several sections)
  function orderOf(state) {
    var seen = new Set();
    var order = [];
    arrangement(state).sections.forEach(function (section) {
      section.puzzles.forEach(function (puzzle) {
        if (!seen.has(puzzle.slug)) {
          seen.add(puzzle.slug);
          order.push(puzzle);
        }
      });
    });
    return order;
  }

  function sectionId(name) {
    return 'group-' + name.toLowerCase().replace(/[^a-z0-9]+/g, '-');
  }

  // ---------- the pictures of the tiles ----------

  // Each drawing is played once, off screen, at 300 x 300 drawn at 2.4 times the pixels, and
  // its picture kept as a PNG the tiles show (an image the browser can let go of when it is
  // far out of view, and one picture for every tile of the same puzzle). Only tiles near the
  // view are drawn, a few per frame, the closest first. A change of theme draws them again.
  var ThumbSize = 300;
  var ThumbScale = 2.4;

  function fitScene(drawing, width, height) {
    var scene = drawing.chooseScene(width, height);
    if (scene != null) {
      drawing.showScene(scene);
    } else {
      drawing.coordinateSystem.zoomExtend(LiveGeometry.GalleryDrawing.getPlane(drawing), 1.08);
    }
  }

  // a player whose canvas is always ThumbSize square, drawn through a pixel ratio of its own
  function makeSquarePlayer(Base) {
    return class extends Base {
      constructor(element, options, pixels) {
        super(element, Object.assign({ theme: currentTheme() }, options));
        this.pixels = pixels;
        this.canvasElement.style.pointerEvents = 'none';
      }

      measureElement() {
        var ratio = typeof this.pixels === 'function' ? this.pixels() : this.pixels;
        if (ratio == null) {
          // created by the base constructor before the pixels are known
          return false;
        }

        if (this.canvas.width === ThumbSize && this.canvas.height === ThumbSize && ratio === this.pixelRatio) {
          return false;
        }

        this.canvas.width = ThumbSize;
        this.canvas.height = ThumbSize;
        this.pixelRatio = ratio;
        this.canvasElement.width = Math.max(1, Math.round(ThumbSize * ratio));
        this.canvasElement.height = Math.max(1, Math.round(ThumbSize * ratio));
        return true;
      }

      resize() {
        if (this.measureElement()) {
          this.invalidate();
        }
      }

      show(drawing) {
        drawing.paintsPaper = false;
        super.show(drawing);
      }

      fitInitial() {
        if (this.drawing != null) {
          fitScene(this.drawing, ThumbSize, ThumbSize);
        }
      }
    };
  }

  var Thumbs = {
    pictures: new Map(),   // slug -> object URL of the picture in the theme it was drawn in
    themeOf: new Map(),    // slug -> that theme
    wanted: new Set(),     // tiles whose picture is wanted (near the view)
    busy: false,
    host: null,

    want: function (tile) {
      this.wanted.add(tile);
      this.schedule();
    },

    unwant: function (tile) {
      this.wanted.delete(tile);
    },

    show: function (tile) {
      var url = this.pictures.get(tile.puzzle.slug);
      if (url != null && this.themeOf.get(tile.puzzle.slug) === currentTheme()) {
        if (tile.img.getAttribute('src') !== url) {
          tile.img.onload = function () { tile.img.classList.add('shown'); };
          tile.img.src = url;
        }

        return true;
      }

      return false;
    },

    restart: function () {
      var self = this;
      tilesInView().forEach(function (tile) { self.wanted.add(tile); });
      this.schedule();
    },

    schedule: function () {
      if (this.busy || this.wanted.size === 0 || Player == null) {
        return;
      }

      this.busy = true;
      var self = this;
      requestAnimationFrame(function () {
        setTimeout(function () { self.next(); }, 0);
      });
    },

    next: function () {
      var self = this;
      // the tile nearest the middle of the view first
      var middle = window.innerHeight / 2;
      var best = null;
      var bestDistance = Infinity;
      this.wanted.forEach(function (tile) {
        if (!tile.element.isConnected || self.show(tile)) {
          self.wanted.delete(tile);
          return;
        }

        var rect = tile.element.getBoundingClientRect();
        var distance = Math.abs(rect.top + rect.height / 2 - middle);
        if (distance < bestDistance) {
          best = tile;
          bestDistance = distance;
        }
      });

      if (best == null) {
        this.busy = false;
        return;
      }

      var tile = best;
      var theme = currentTheme();
      fetchText(tile.puzzle).then(function (text) {
        return self.draw(tile.puzzle, text, theme);
      }).then(function () {
        // every tile of this puzzle in view takes the picture
        self.wanted.forEach(function (other) {
          if (other.puzzle === tile.puzzle && self.show(other)) {
            self.wanted.delete(other);
          }
        });
      }).catch(function (error) {
        console.error('CatrionaAgg: ' + tile.puzzle.slug + ': ' + error.message);
        self.wanted.delete(tile);
      }).then(function () {
        self.busy = false;
        self.schedule();
      });
    },

    draw: function (puzzle, text, theme) {
      var self = this;
      if (this.host == null) {
        this.host = document.createElement('div');
        this.host.style.cssText = 'position: fixed; left: -10000px; top: 0; width: ' + ThumbSize + 'px; height: ' + ThumbSize + 'px; font-family: var(--sans);';
        document.body.appendChild(this.host);
      }

      var player = new ThumbPlayer(this.host, { theme: theme }, ThumbScale);
      try {
        player.measureElement();
        player.load(withoutHandles(text));
        player.render();
      } catch (error) {
        player.render = function () { };
        player.dispose();
        return Promise.reject(error);
      }

      return new Promise(function (resolve) {
        player.canvasElement.toBlob(function (blob) {
          player.render = function () { };
          player.dispose();
          var old = self.pictures.get(puzzle.slug);
          if (old != null) {
            URL.revokeObjectURL(old);
          }

          self.pictures.set(puzzle.slug, URL.createObjectURL(blob));
          self.themeOf.set(puzzle.slug, theme);
          resolve();
        }, 'image/png');
      });
    }
  };

  var Player = null;
  var ThumbPlayer = null;
  var LivePlayer = null;

  // ---------- a tile ----------

  var tiles = [];           // the tiles on the page now
  var tilePool = new Map(); // slug -> every tile made for the puzzle, to be used again (a puzzle may be in several sections)

  function Tile(puzzle) {
    this.puzzle = puzzle;
    var element = this.element = document.createElement('a');
    element.className = 'tile';
    element.href = Base + '/' + puzzle.slug;
    var art = this.art = document.createElement('div');
    art.className = 'art';
    var img = this.img = document.createElement('img');
    img.alt = puzzle.title + ': ' + plainQuestion(puzzle.statement);
    img.decoding = 'async';
    art.appendChild(img);
    element.appendChild(art);
    var meta = document.createElement('div');
    meta.className = 'meta';
    var title = document.createElement('span');
    title.className = 'title';
    title.textContent = puzzle.title;
    var sub = document.createElement('span');
    sub.className = 'sub';
    sub.textContent = shortDate(puzzle.date) + ' · ' + puzzle.question;
    meta.appendChild(title);
    meta.appendChild(sub);
    element.appendChild(meta);
    element.title = plainQuestion(puzzle.statement);
    var tile = this;
    element.addEventListener('pointerenter', function (e) {
      if (e.pointerType === 'mouse') {
        startHover(tile);
      }
    });
    element.addEventListener('pointerleave', function () {
      if (hover != null && hover.tile === tile) {
        endHover();
      }
    });
    if (puzzle.live) {
      this.markLive();
    }
  }

  Tile.prototype.markLive = function () {
    if (this.element.querySelector('.live') != null) {
      return;
    }

    var badge = document.createElement('span');
    badge.className = 'live';
    badge.title = 'Something to drag';
    badge.innerHTML = '<svg><use href="' + Icons + '#icon-drag"/></svg>';
    this.element.appendChild(badge);
  };

  // tiles come near the view: their pictures are wanted
  var observer = isGallery ? new IntersectionObserver(function (entries) {
    entries.forEach(function (entry) {
      var tile = entry.target.tile;
      if (tile == null) {
        return;
      }

      tile.near = entry.isIntersecting;
      if (entry.isIntersecting) {
        if (!Thumbs.show(tile)) {
          Thumbs.want(tile);
        }
      } else {
        Thumbs.unwant(tile);
      }
    });
  }, { rootMargin: '600px 0px' }) : null;

  function tilesInView() {
    return tiles.filter(function (tile) { return tile.near; });
  }

  // ---------- the drawing under the mouse comes alive: its points drift ----------

  var hover = null;

  function startHover(tile) {
    endHover();
    if (Player == null || reduceMotion) {
      return;
    }

    var started = { tile: tile, player: null, drifters: null, timer: 0 };
    hover = started;
    // a pass of the mouse across the gallery doesn't start a player per tile
    started.timer = setTimeout(function () {
      fetchText(tile.puzzle).then(function (text) {
        if (hover !== started || !hasHandles(text)) {
          return;
        }

        var art = tile.art;
        started.player = new LivePlayer(art, {}, function () {
          var rect = art.getBoundingClientRect();
          return rect.width > 0 ? rect.width / ThumbSize * (window.devicePixelRatio || 1) : null;
        });
        var canvas = started.player.canvasElement;
        canvas.style.position = 'absolute';
        canvas.style.inset = '0';
        canvas.style.width = '100%';
        canvas.style.height = '100%';
        started.player.measureElement();
        started.player.load(withoutHandles(text));
        started.player.render();
        tile.img.style.visibility = 'hidden';
        started.stop = drift(started.player, 22, 2.8);
      });
    }, 110);
  }

  function endHover() {
    var ended = hover;
    hover = null;
    if (ended == null) {
      return;
    }

    clearTimeout(ended.timer);
    if (ended.stop != null) {
      ended.stop();
    }

    if (ended.player != null) {
      ended.player.render = function () { };
      ended.player.dispose();
    }

    ended.tile.img.style.visibility = '';
  }

  // The points a reader would drag go round in small circles (in pixels of the player's
  // canvas), out of phase, and the construction follows; stop() puts them back.
  function drift(player, radiusInPixels, secondsPerTurn) {
    // the handles, shown or not: in these drawings the only points that can be dragged
    var drawing = player.drawing;
    var drifters = drawing.figures.list.filter(function (figure) {
      return figure.isPoint === true && figure.isMovable === true && figure.allowMove();
    });
    var items = drifters.map(function (point, index) {
      return { point: point, home: point.coordinates, phase: index * 2.4, direction: index % 2 === 0 ? 1 : -1 };
    });
    if (items.length === 0) {
      return function () { };
    }

    var running = true;
    var startedAt = performance.now();
    var frame = function () {
      if (!running) {
        return;
      }

      var seconds = (performance.now() - startedAt) / 1000;
      try {
        var radius = radiusInPixels / drawing.coordinateSystem.unitLength;
        // eases in: no jump when the drift starts
        var amplitude = Math.min(1, seconds / 0.6);
        var turn = 2 * Math.PI * seconds / secondsPerTurn;
        for (const item of items) {
          var angle = item.direction * turn + item.phase;
          item.point.moveTo(new LiveGeometry.Point(
            item.home.x + amplitude * radius * (Math.cos(angle) - Math.cos(item.phase)),
            item.home.y + amplitude * radius * (Math.sin(angle) - Math.sin(item.phase))));
        }

        drawing.recalculate();
        player.invalidate();
      } catch (error) {
        console.error('CatrionaAgg: ' + error.message);
        running = false;
        return;
      }

      requestAnimationFrame(frame);
    };
    requestAnimationFrame(frame);
    return function () {
      running = false;
      try {
        for (const item of items) {
          item.point.moveTo(item.home);
        }

        drawing.recalculate();
        player.invalidate();
      } catch (error) {
      }
    };
  }

  // ---------- the gallery: search, arrangement, order ----------

  var state = Object.assign({}, DefaultState);
  var groupsHost = document.getElementById('groups');
  var countText = document.getElementById('count');
  var emptyText = document.getElementById('empty');
  var jump = document.getElementById('jump');
  var searchBox = document.getElementById('search');
  var arrange = document.getElementById('arrange');
  var arrangeThumb = document.getElementById('arrange-thumb');
  var orderButton = document.getElementById('order');

  // The address says the arrangement when it says anything (a link to the gallery as
  // arranged, a puzzle's tag), else the tab's last one, else the newest first.
  function readGalleryState() {
    var query = new URLSearchParams(location.search);
    if (query.has('by') || query.has('q') || query.has('order')) {
      return validState({ group: query.get('by'), query: query.get('q') || '', oldestFirst: query.get('order') === 'oldest' });
    }

    return storedState() || Object.assign({}, DefaultState);
  }

  function galleryAddress() {
    var query = new URLSearchParams();
    if (state.group !== 'none') { query.set('by', state.group); }
    if (state.query.trim() !== '') { query.set('q', state.query.trim()); }
    if (state.oldestFirst) { query.set('order', 'oldest'); }
    var text = query.toString();
    return Base + (text !== '' ? '?' + text : '') + location.hash;
  }

  // the arrangement goes into the address (to be shared) and into the tab (for the puzzles' arrows)
  function publishState() {
    storeState(state);
    history.replaceState(null, '', galleryAddress());
  }

  // Lays the gallery out for the state, the tiles that stay gliding to their new places
  // (FLIP: where each was, where it is now, and a transform from one to the other that goes to nothing).
  function render(animate) {
    endHover();
    var result = arrangement(state);
    var shownCount = result.shown.length;
    var before = new Map();
    if (animate && !reduceMotion) {
      tiles.forEach(function (tile) {
        if (tile.near) {
          var rect = tile.element.getBoundingClientRect();
          if (!before.has(tile.puzzle.slug)) {
            before.set(tile.puzzle.slug, rect);
          }
        }
      });
    }

    // every tile made so far can be taken again for its puzzle (with its picture), each once a layout
    var used = new Set();
    var take = function (puzzle) {
      var list = tilePool.get(puzzle.slug);
      if (list == null) {
        list = [];
        tilePool.set(puzzle.slug, list);
      }

      var tile = list.find(function (made) { return !used.has(made); });
      if (tile == null) {
        tile = new Tile(puzzle);
        list.push(tile);
      }

      used.add(tile);
      return tile;
    };

    var nextTiles = [];
    var fragment = document.createDocumentFragment();
    result.sections.forEach(function (section) {
      var group = document.createElement('section');
      group.className = 'group';
      if (section.name != null) {
        group.id = sectionId(section.name);
        var heading = document.createElement('h2');
        heading.textContent = section.name;
        var count = document.createElement('span');
        count.className = 'n';
        count.textContent = section.puzzles.length;
        heading.appendChild(count);
        group.appendChild(heading);
      }

      var grid = document.createElement('div');
      grid.className = 'grid';
      section.puzzles.forEach(function (puzzle) {
        var tile = take(puzzle);
        nextTiles.push(tile);
        grid.appendChild(tile.element);
      });
      group.appendChild(grid);
      fragment.appendChild(group);
    });

    // what was not taken stays in the pool, off the page
    tiles.forEach(function (tile) {
      if (!used.has(tile)) {
        observer.unobserve(tile.element);
        tile.near = false;
        Thumbs.unwant(tile);
      }
    });

    groupsHost.replaceChildren(fragment);
    tiles = nextTiles;
    tiles.forEach(function (tile) {
      tile.element.tile = tile;
      observer.observe(tile.element);
      Thumbs.show(tile);
    });

    countText.textContent = shownCount === puzzles.length
      ? puzzles.length + ' puzzles redrawn so far, of ' + everything.length + ' in the catalogue'
      : shownCount + ' of ' + puzzles.length + ' puzzles';
    emptyText.hidden = shownCount > 0;
    if (shownCount === 0) {
      emptyText.textContent = state.query.trim() !== ''
        ? 'Nothing here matches “' + state.query.trim() + '”.'
        : 'Nothing here yet.';
    }

    updateJump(result.sections);

    if (before.size > 0) {
      var viewHeight = window.innerHeight;
      var seenNow = new Set();
      tiles.forEach(function (tile, index) {
        var rect = tile.element.getBoundingClientRect();
        if (rect.bottom < -200 || rect.top > viewHeight + 200) {
          return;
        }

        var old = seenNow.has(tile.puzzle.slug) ? null : before.get(tile.puzzle.slug);
        seenNow.add(tile.puzzle.slug);
        if (old != null) {
          var dx = old.left - rect.left;
          var dy = old.top - rect.top;
          if (Math.abs(dx) > 0.5 || Math.abs(dy) > 0.5) {
            tile.element.animate([
              { transform: 'translate(' + dx + 'px, ' + dy + 'px)' },
              { transform: 'translate(0, 0)' }
            ], { duration: 560, easing: 'cubic-bezier(0.2, 0.8, 0.2, 1)' });
          }
        } else {
          tile.element.animate([
            { opacity: 0, transform: 'scale(0.92)' },
            { opacity: 1, transform: 'scale(1)' }
          ], { duration: 420, delay: Math.min(index, 12) * 18, easing: 'cubic-bezier(0.2, 0.8, 0.2, 1)', fill: 'backwards' });
        }
      });
      groupsHost.querySelectorAll('.group h2').forEach(function (heading) {
        heading.animate([{ opacity: 0, transform: 'translateY(6px)' }, { opacity: 1, transform: 'none' }], { duration: 380, easing: 'ease-out' });
      });
    }
  }

  function scrollToSection(id) {
    var target = document.getElementById(id);
    if (target != null) {
      window.scrollTo(0, target.getBoundingClientRect().top + window.scrollY - 120);
    }
  }

  function updateJump(sections) {
    jump.replaceChildren();
    jump.hidden = state.group === 'none' || sections.length < 2;
    if (jump.hidden) {
      return;
    }

    sections.forEach(function (section) {
      var button = document.createElement('button');
      button.type = 'button';
      button.textContent = section.name;
      var count = document.createElement('span');
      count.textContent = section.puzzles.length;
      button.appendChild(count);
      button.addEventListener('click', function () {
        var target = document.getElementById(sectionId(section.name));
        if (target != null) {
          target.scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
        }
      });
      jump.appendChild(button);
    });
  }

  function updateControls() {
    searchBox.value = state.query;
    arrange.querySelectorAll('button').forEach(function (button) {
      button.setAttribute('aria-pressed', button.getAttribute('data-group') === state.group ? 'true' : 'false');
    });
    orderButton.textContent = state.oldestFirst ? 'Oldest first' : 'Newest first';
    placeArrangeThumb();
  }

  function placeArrangeThumb() {
    var chosen = arrange.querySelector('button[aria-pressed="true"]');
    if (chosen == null) {
      return;
    }

    arrangeThumb.style.width = chosen.offsetWidth + 'px';
    arrangeThumb.style.transform = 'translateX(' + chosen.offsetLeft + 'px)';
  }

  function stateChanged(animate) {
    updateControls();
    publishState();
    render(animate);
  }

  // ---------- the newest puzzle, alive, beside the introduction ----------

  var featurePlayer = null;

  function showFeature() {
    var featured = puzzles.find(function (puzzle) { return puzzle.live; }) || puzzles[0];
    if (featured == null || Player == null) {
      return;
    }

    var feature = document.getElementById('feature');
    var art = document.getElementById('feature-art');
    feature.href = Base + '/' + featured.slug;
    document.getElementById('feature-name').textContent = featured.title;
    feature.hidden = false;
    if (getComputedStyle(feature).display === 'none') {
      return;
    }

    fetchText(featured).then(function (text) {
      featurePlayer = new LivePlayer(art, {}, function () {
        var rect = art.getBoundingClientRect();
        return rect.width > 0 ? rect.width / ThumbSize * (window.devicePixelRatio || 1) : null;
      });
      featurePlayer.measureElement();
      featurePlayer.load(withoutHandles(text));
      if (!reduceMotion) {
        drift(featurePlayer, 30, 6);
      }
    }).catch(function (error) {
      console.error('CatrionaAgg: ' + error.message);
    });
  }

  function startGallery() {
    // the page comes back where it was when the reader returns to it (the scroll is ours to
    // restore: the browser's comes before the tiles are there)
    history.scrollRestoration = 'manual';
    var navigation = performance.getEntriesByType('navigation')[0];
    var returning = readSession(ReturnKey) != null || (navigation != null && navigation.type === 'back_forward');
    writeSession(ReturnKey, null);

    state = readGalleryState();
    updateControls();
    publishState();
    render(false);
    showFeature();

    if (location.hash.length > 1) {
      // a puzzle's tag: the section it names
      scrollToSection(decodeURIComponent(location.hash.substring(1)));
    } else if (returning) {
      var scroll = parseFloat(readSession(ScrollKey));
      if (scroll > 0) {
        window.scrollTo(0, scroll);
      }
    }

    arrange.addEventListener('click', function (e) {
      var button = e.target.closest('button');
      if (button == null || button.getAttribute('data-group') === state.group) {
        return;
      }

      state.group = button.getAttribute('data-group');
      stateChanged(true);
    });

    orderButton.addEventListener('click', function () {
      state.oldestFirst = !state.oldestFirst;
      stateChanged(true);
    });

    var searchTimer = 0;
    searchBox.addEventListener('input', function () {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(function () {
        state.query = searchBox.value;
        publishState();
        render(true);
      }, 140);
    });

    window.addEventListener('resize', placeArrangeThumb);

    var galleryBar = document.getElementById('gallery-bar');
    window.addEventListener('scroll', function () {
      galleryBar.classList.toggle('scrolled', window.scrollY > 4);
    }, { passive: true });

    window.addEventListener('pagehide', function () {
      endHover();
      writeSession(ScrollKey, String(window.scrollY));
    });

    Thumbs.schedule();
  }

  // ---------- a puzzle ----------

  var puzzlePlayer = null;

  function element(tag, className, text) {
    var made = document.createElement(tag);
    if (className) {
      made.className = className;
    }

    if (text != null) {
      made.textContent = text;
    }

    return made;
  }

  // a button to one of Andrew Stacey's pages for the puzzle, on Mathforge Notes
  function linkCard(href, color, icon, title, tooltip) {
    var card = element('a', 'link-card');
    card.href = href;
    card.target = '_blank';
    card.rel = 'noopener';
    card.title = tooltip;
    card.innerHTML = '<svg style="color: ' + color + '"><use href="' + Icons + '#' + icon + '"/></svg>';
    card.appendChild(element('span', null, title));
    return card;
  }

  function fillPanel(puzzle) {
    document.getElementById('p-date').textContent = longDate(puzzle.date);
    document.getElementById('p-title').textContent = puzzle.title;
    document.getElementById('p-question').textContent = '“' + plainQuestion(puzzle.statement) + '”';
    var asked = document.getElementById('p-asked');
    asked.replaceChildren();
    asked.append('Catriona Agg');
    if (puzzle.thread != null) {
      asked.append(', ');
      var thread = element('a', null, 'on BlueSky');
      thread.href = puzzle.thread;
      thread.target = '_blank';
      thread.rel = 'noopener';
      asked.appendChild(thread);
    }

    var links = document.getElementById('p-links');
    links.replaceChildren(
      linkCard(mathforge(puzzle.name), '#14B67F', 'icon-picture', 'Original', 'Catriona’s original drawing, on Mathforge Notes'),
      linkCard(mathforge(puzzle.name + ' techniques'), '#F26313', 'icon-tools', 'Techniques', 'The techniques a solution uses, on Mathforge Notes'),
      linkCard(mathforge(puzzle.name + ' solution'), '#8A12B5', 'icon-key', 'Solution', 'Andrew Stacey’s solution, on Mathforge Notes'));

    // a tag is the gallery arranged by its kind, at its section
    var tags = document.getElementById('p-tags');
    tags.replaceChildren();
    var addTag = function (text, group) {
      var tag = element('a', null, text);
      tag.href = Base + '?by=' + group + '#' + sectionId(text);
      tag.title = 'Show the puzzles by ' + group;
      tags.appendChild(tag);
    };
    addTag(String(puzzle.year), 'year');
    topicsOf(puzzle).forEach(function (topic) { addTag(topic, 'topic'); });
    puzzle.shapes.forEach(function (shape) { addTag(shape, 'shape'); });
  }

  // The arrows step through the puzzles in the order the gallery showed them in this tab;
  // the newest first when it showed none, or a selection without this puzzle.
  function setUpStepper(puzzle) {
    var order = orderOf(storedState() || DefaultState);
    if (order.indexOf(puzzle) < 0) {
      order = puzzles;
    }

    var index = order.indexOf(puzzle);
    var previous = order[(index + order.length - 1) % order.length];
    var next = order[(index + 1) % order.length];
    document.getElementById('previous').href = Base + '/' + previous.slug;
    document.getElementById('next').href = Base + '/' + next.slug;
    document.getElementById('position').textContent = (index + 1) + ' / ' + order.length;

    // the neighbors' drawings, for the next step
    [previous, next].forEach(function (neighbor) {
      if (neighbor !== puzzle) {
        var link = document.createElement('link');
        link.rel = 'prefetch';
        link.href = drawingAddress(neighbor);
        document.head.appendChild(link);
      }
    });
  }

  function slugOfPath(path) {
    var match = /^\/CatrionaAgg\/([^/?#]+)\/?$/i.exec(path);
    return match != null && !/\.html$/i.test(match[1]) ? decodeURIComponent(match[1]).toLowerCase() : '';
  }

  // back to the gallery, which takes the reader back to where they were in it
  function goToGallery() {
    writeSession(ReturnKey, '1');
    location.href = Base;
  }

  function startPuzzle() {
    var slug = slugOfPath(location.pathname);
    var puzzle = bySlug.get(slug);
    if (puzzle == null) {
      var known = everything.find(function (other) { return other.slug === slug; });
      // in the catalogue, not drawn yet: its page is Mathforge's
      showMissing(known != null ? '“' + known.title + '” isn’t redrawn here yet. ' : 'There is no puzzle at this address. ', known);
      return;
    }

    document.title = puzzle.title + ' · Catriona Agg · Live Geometry';
    fillPanel(puzzle);
    setUpStepper(puzzle);

    var playerElement = document.getElementById('player');
    var currentText = null;
    fetchText(puzzle).then(function (text) {
      currentText = text;
      puzzlePlayer = new Player(playerElement, { lgf: text, theme: currentTheme(), wheel: 'zoom' });
      // for a look from the console, as the dev page has LiveGeometry.players
      window.catrionaPlayer = puzzlePlayer;
      if (puzzlePlayer.loadErrors != null) {
        console.warn('CatrionaAgg: ' + puzzle.slug + ': ' + puzzlePlayer.loadErrors);
      }
    }).catch(function (error) {
      console.error('CatrionaAgg: ' + error.message);
    });

    document.getElementById('back').addEventListener('click', function (e) {
      if (isPlainClick(e)) {
        e.preventDefault();
        goToGallery();
      }
    });
    document.getElementById('fit').addEventListener('click', function () {
      if (puzzlePlayer != null) {
        puzzlePlayer.fit();
      }
    });
    document.getElementById('reset').addEventListener('click', function () {
      if (puzzlePlayer != null && currentText != null) {
        puzzlePlayer.load(currentText);
      }
    });

    window.addEventListener('keydown', function (e) {
      if (e.target instanceof HTMLInputElement || e.ctrlKey || e.metaKey || e.altKey) {
        return;
      }

      if (e.key === 'Escape') {
        goToGallery();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        location.href = document.getElementById('previous').href;
      } else if (e.key === 'ArrowRight' || e.key === 'PageDown') {
        e.preventDefault();
        location.href = document.getElementById('next').href;
      } else if (e.key === 'Home') {
        e.preventDefault();
        if (puzzlePlayer != null) {
          puzzlePlayer.fit();
        }
      }
    });
  }

  function showMissing(text, known) {
    document.getElementById('puzzle-body').hidden = true;
    document.querySelector('#puzzle-view .stepper').style.visibility = 'hidden';
    var missing = element('p', 'missing', text);
    document.getElementById('puzzle-view').appendChild(missing);
    if (known != null) {
      var link = element('a', null, 'See it on Mathforge Notes');
      link.href = mathforge(known.name);
      missing.appendChild(link);
      missing.append(', or ');
    }

    var all = element('a', null, 'see all the puzzles');
    all.href = Base;
    missing.appendChild(all);
    missing.append('.');
  }

  // ---------- start ----------

  if (typeof LiveGeometry === 'undefined') {
    (isGallery ? document.getElementById('groups') : document.getElementById('player')).textContent = 'The drawings could not be loaded.';
    return;
  }

  Player = LiveGeometry.Player;
  ThumbPlayer = makeSquarePlayer(Player);
  LivePlayer = makeSquarePlayer(Player);

  var fontsReady = document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve();
  fetch(Base + '/puzzles.json').then(function (response) {
    if (!response.ok) {
      throw new Error('could not load the catalog (' + response.status + ')');
    }

    return response.json();
  }).then(function (catalog) {
    everything = catalog.puzzles.map(function (entry) {
      var puzzle = Object.assign({}, entry);
      puzzle.date = parseDate(entry.date);
      puzzle.year = puzzle.date.getFullYear();
      puzzle.categories = entry.categories || [];
      puzzle.techniques = entry.techniques || [];
      puzzle.shapes = shapesOf(puzzle);
      puzzle.question = questionOf(puzzle);
      // something to drag (the catalog says, from the drawing's handles)
      puzzle.live = entry.live === true;
      puzzle.searchText = [puzzle.title, plainQuestion(puzzle.statement), puzzle.year, puzzle.categories.join(' '), puzzle.techniques.join(' '), puzzle.shapes.join(' '), puzzle.question].join(' ').toLowerCase();
      return puzzle;
    });
    puzzles = everything.filter(function (puzzle) { return puzzle.drawing != null; });
    puzzles.forEach(function (puzzle) { bySlug.set(puzzle.slug, puzzle); });
    // the drawings' numbers are in the page's font
    return fontsReady;
  }).then(function () {
    if (isGallery) {
      startGallery();
    } else {
      startPuzzle();
    }
  }).catch(function (error) {
    console.error('CatrionaAgg: ' + error.message);
    (isGallery ? document.getElementById('groups') : document.getElementById('player')).textContent = 'The puzzles could not be loaded.';
  });
});
