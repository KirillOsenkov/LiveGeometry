// The Geometry Snacks library's script, on the JavaScript player (/player/1/player.js, loaded
// first): the list at /snacks (index.html) and a puzzle at /snacks/<slug> (puzzle.html, and
// the copy of it the build writes for each puzzle). Each page runs its own part; the catalog
// and the theme are shared. It is the Catriona Agg library's script (CatrionaAgg/catriona.js)
// without the gallery's arrangement: the puzzles are listed in the books' order, and a
// puzzle's arrows step through them in that order.

document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  var root = document.documentElement;
  var Base = '/snacks';
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

  var books = new Map();  // by number: the book
  var puzzles = [];       // the ones with a drawing, in the books' order
  var bySlug = new Map();
  var texts = new Map();  // by slug: a promise of the drawing's text

  function bookOf(puzzle) {
    return books.get(puzzle.book);
  }

  // "Geometry Snacks #12"
  function placeOf(puzzle) {
    return bookOf(puzzle).title + ' #' + puzzle.number;
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

  // what a reader drags is left out of a picture: the tile shows the figure as the book draws it
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

  var HandoffKey = 'LiveGeometry.Handoff';

  function writeSession(key, value) {
    try {
      if (value == null) {
        sessionStorage.removeItem(key);
      } else {
        sessionStorage.setItem(key, value);
      }
    } catch (e) { }
  }

  // ---------- the pictures of the tiles ----------

  // Each drawing is played once, off screen, at 300 x 300 drawn at 2.4 times the pixels, and
  // its picture kept as a PNG the tiles show. Only tiles near the view are drawn, one at a
  // time, the closest first; their texts are fetched as soon as they come near, all at once.
  // A change of theme draws them again.
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

      // drawn at once: the canvas's new size has cleared it, and the frame is painted next (Player.resize)
      resize() {
        if (this.measureElement()) {
          this.render();
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
      // the text on its way while the tiles before it are drawn (a failure is reported when it is drawn)
      fetchText(tile.puzzle).catch(function () { });
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
        console.error('Snacks: ' + tile.puzzle.slug + ': ' + error.message);
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

  var tiles = [];

  function Tile(puzzle) {
    this.puzzle = puzzle;
    var element = this.element = document.createElement('a');
    element.className = 'tile';
    element.href = Base + '/' + puzzle.slug;
    var art = this.art = document.createElement('div');
    art.className = 'art';
    var img = this.img = document.createElement('img');
    img.alt = puzzle.title + ': ' + puzzle.statement;
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
    sub.textContent = placeOf(puzzle) + ' · ' + puzzle.chapter;
    meta.appendChild(title);
    meta.appendChild(sub);
    element.appendChild(meta);
    element.title = puzzle.statement;
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
      var badge = document.createElement('span');
      badge.className = 'live';
      badge.title = 'Something to drag';
      badge.innerHTML = '<svg><use href="' + Icons + '#icon-drag"/></svg>';
      element.appendChild(badge);
    }
  }

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

    var started = { tile: tile, player: null, timer: 0 };
    hover = started;
    // a pass of the mouse across the list doesn't start a player per tile
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
        console.error('Snacks: ' + error.message);
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

  // ---------- the list ----------

  function startGallery() {
    var groupsHost = document.getElementById('groups');
    var group = document.createElement('section');
    group.className = 'group';
    var grid = document.createElement('div');
    grid.className = 'grid';
    puzzles.forEach(function (puzzle) {
      var tile = new Tile(puzzle);
      tiles.push(tile);
      grid.appendChild(tile.element);
    });
    group.appendChild(grid);
    groupsHost.replaceChildren(group);
    tiles.forEach(function (tile) {
      tile.element.tile = tile;
      observer.observe(tile.element);
    });

    document.getElementById('count').textContent = puzzles.length + (puzzles.length === 1 ? ' puzzle' : ' puzzles');
    var emptyText = document.getElementById('empty');
    emptyText.hidden = puzzles.length > 0;
    emptyText.textContent = 'Nothing here yet.';

    var galleryBar = document.getElementById('gallery-bar');
    window.addEventListener('scroll', function () {
      galleryBar.classList.toggle('scrolled', window.scrollY > 4);
    }, { passive: true });

    window.addEventListener('pagehide', endHover);
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

  function fillPanel(puzzle) {
    var book = bookOf(puzzle);
    document.getElementById('p-where').textContent = placeOf(puzzle) + ' · ' + puzzle.chapter;
    document.getElementById('p-title').textContent = puzzle.title;
    document.getElementById('p-question').textContent = puzzle.statement;
    var bookLink = document.getElementById('p-book');
    bookLink.textContent = book.title;
    bookLink.href = book.amazon;
  }

  // The arrows step through the puzzles in the books' order
  function setUpStepper(puzzle) {
    var index = puzzles.indexOf(puzzle);
    var previous = puzzles[(index + puzzles.length - 1) % puzzles.length];
    var next = puzzles[(index + 1) % puzzles.length];
    document.getElementById('previous').href = Base + '/' + previous.slug;
    document.getElementById('next').href = Base + '/' + next.slug;
    document.getElementById('position').textContent = (index + 1) + ' / ' + puzzles.length;

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
    var match = /^\/snacks\/([^/?#]+)\/?$/i.exec(path);
    return match != null && !/\.html$/i.test(match[1]) ? decodeURIComponent(match[1]).toLowerCase() : '';
  }

  // The drawing in the editor, the app at /drawing, in a new tab, as the user's own (unsaved):
  // handed over through the tab's storage as the Catriona Agg library and /web do
  // (main.js, MainView.Handoff.cs).
  function editInLiveGeometry(text) {
    var bytes = new TextEncoder().encode(text);
    var binary = '';
    for (var i = 0; i < bytes.length; i += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }

    writeSession(HandoffKey, JSON.stringify({ action: 'edit', data: btoa(binary) }));
    if (window.open('/drawing', '_blank') != null) {
      writeSession(HandoffKey, null);
    } else {
      // a new tab refused (a popup blocker): this tab goes instead
      location.href = '/drawing';
    }
  }

  function startPuzzle() {
    var slug = slugOfPath(location.pathname);
    var puzzle = bySlug.get(slug);
    if (puzzle == null) {
      showMissing();
      return;
    }

    document.title = puzzle.title + ' · Geometry Snacks · Live Geometry';
    fillPanel(puzzle);
    setUpStepper(puzzle);

    var playerElement = document.getElementById('player');
    var currentText = null;
    fetchText(puzzle).then(function (text) {
      currentText = text;
      puzzlePlayer = new Player(playerElement, { lgf: text, theme: currentTheme(), wheel: 'zoom' });
      // for a look from the console, as the dev page has LiveGeometry.players
      window.snacksPlayer = puzzlePlayer;
      if (puzzlePlayer.loadErrors != null) {
        console.warn('Snacks: ' + puzzle.slug + ': ' + puzzlePlayer.loadErrors);
      }
    }).catch(function (error) {
      console.error('Snacks: ' + error.message);
    });

    document.getElementById('back').addEventListener('click', function (e) {
      if (isPlainClick(e)) {
        e.preventDefault();
        location.href = Base;
      }
    });
    document.getElementById('edit').addEventListener('click', function () {
      if (currentText != null) {
        editInLiveGeometry(currentText);
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
        location.href = Base;
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

  function showMissing() {
    document.getElementById('puzzle-body').hidden = true;
    document.querySelector('#puzzle-view .stepper').style.visibility = 'hidden';
    var missing = element('p', 'missing', 'There is no puzzle at this address. ');
    document.getElementById('puzzle-view').appendChild(missing);
    var all = element('a', null, 'See all the puzzles');
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
    catalog.books.forEach(function (book) { books.set(book.book, book); });
    puzzles = catalog.puzzles.filter(function (puzzle) { return puzzle.drawing != null; });
    puzzles.forEach(function (puzzle) {
      puzzle.live = puzzle.live === true;
      bySlug.set(puzzle.slug, puzzle);
    });
    // the drawings' numbers are in the page's font
    return fontsReady;
  }).then(function () {
    if (isGallery) {
      startGallery();
    } else {
      startPuzzle();
    }
  }).catch(function (error) {
    console.error('Snacks: ' + error.message);
    (isGallery ? document.getElementById('groups') : document.getElementById('player')).textContent = 'The puzzles could not be loaded.';
  });
});
