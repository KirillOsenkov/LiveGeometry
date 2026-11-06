/* donut.js: a low-poly donut, as on the cover of Geometry Snacks, drawn on a canvas.

   No dependencies. The donut turns slowly on its own; the pointer drags it, and a fling
   keeps turning and slows down. Plain canvas 2D: a jittered torus grid with a random
   diagonal in every cell (so the facets are irregular, like the cover's), flat shading
   from a fixed light, facets sorted far to near.

     var donut = new Donut(canvas, { spin: 0.12, shadow: true });
     donut.options.spinAxis = 'screen';   // options are read live
     donut.stop(); donut.start(); donut.dispose();

   The canvas is sized to its CSS box (device pixels) and drawn again at once when the box
   changes. The donut is drawn only while the canvas is on screen and the tab is visible. The
   slow turn is on whatever the system says about motion: it is slow and harmless. */
(function (global) {
  'use strict';

  var defaults = {
    seed: 7,
    around: 36,          // facets around the ring
    across: 14,          // facets around the tube
    radius: 1,           // the ring's radius
    tube: 0.56,          // the tube's radius, in ring radii
    jitter: 0.5,         // how far a vertex strays from its grid cell, 0..1
    bumps: 0.05,         // random change of the tube radius per vertex
    tilt: 50,            // degrees the donut is tipped toward the viewer at the start
    turn: -20,           // degrees it is turned around the screen's vertical at the start
    spin: 0.12,          // radians per second of the slow turn
    spinAxis: 'model',   // 'model': around the donut's own axis; 'screen': around the screen's vertical
    friction: 1.4,       // per second: how fast a fling slows down (larger = sooner)
    dragSpeed: 0.0075,   // radians per CSS pixel of drag
    maxFling: 10,        // radians per second, the fastest a fling goes
    fit: 0.84,           // the donut's width as a fraction of the shorter side
    distance: 5,         // the camera's distance, in ring radii (the perspective)
    light: [-0.45, 0.8, 0.55],
    ambient: 0.36,
    diffuse: 0.8,
    gloss: 0.35,         // the frosting's highlights
    shade: 0.18,         // per-facet random brightness, plus or minus
    frosting: [243, 158, 176],
    dough: [214, 128, 48],
    frostingLine: 0.25,  // where the frosting ends: the sine of the tube angle, jagged around it
    shadow: true,        // a soft shadow under the donut
    interactive: true,   // the pointer drags it
    touch: true,         // fingers drag it too (the canvas then takes touch-action: none)
    offset: [0, 0],      // where the center is: fractions of the width and height from the middle
    anchor: null         // an element: the center is the middle of its box instead (offset then adds to it)
  };

  function mulberry32(seed) {
    return function () {
      seed |= 0;
      seed = seed + 0x6D2B79F5 | 0;
      var t = Math.imul(seed ^ seed >>> 15, 1 | seed);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }

  /* ---------- 3x3 matrices, row-major ---------- */
  function multiply(a, b) {
    return [
      a[0] * b[0] + a[1] * b[3] + a[2] * b[6], a[0] * b[1] + a[1] * b[4] + a[2] * b[7], a[0] * b[2] + a[1] * b[5] + a[2] * b[8],
      a[3] * b[0] + a[4] * b[3] + a[5] * b[6], a[3] * b[1] + a[4] * b[4] + a[5] * b[7], a[3] * b[2] + a[4] * b[5] + a[5] * b[8],
      a[6] * b[0] + a[7] * b[3] + a[8] * b[6], a[6] * b[1] + a[7] * b[4] + a[8] * b[7], a[6] * b[2] + a[7] * b[5] + a[8] * b[8]
    ];
  }

  // the rotation by angle around the unit axis (x, y, z)
  function rotation(x, y, z, angle) {
    var c = Math.cos(angle), s = Math.sin(angle), t = 1 - c;
    return [
      t * x * x + c, t * x * y - s * z, t * x * z + s * y,
      t * x * y + s * z, t * y * y + c, t * y * z - s * x,
      t * x * z - s * y, t * y * z + s * x, t * z * z + c
    ];
  }

  // Gram-Schmidt, so that many small rotations don't drift into a skew
  function orthonormalize(m) {
    var r = m.slice();
    var n = Math.hypot(r[0], r[1], r[2]);
    r[0] /= n; r[1] /= n; r[2] /= n;
    var d = r[0] * r[3] + r[1] * r[4] + r[2] * r[5];
    r[3] -= d * r[0]; r[4] -= d * r[1]; r[5] -= d * r[2];
    n = Math.hypot(r[3], r[4], r[5]);
    r[3] /= n; r[4] /= n; r[5] /= n;
    r[6] = r[1] * r[5] - r[2] * r[4];
    r[7] = r[2] * r[3] - r[0] * r[5];
    r[8] = r[0] * r[4] - r[1] * r[3];
    return r;
  }

  /* ---------- the mesh ---------- */
  function buildMesh(o) {
    var rand = mulberry32(o.seed);
    var U = o.around, V = o.across;
    var count = U * V;
    var px = new Float32Array(count), py = new Float32Array(count), pz = new Float32Array(count);
    var sine = new Float32Array(count);
    var i, j, k;
    for (i = 0; i < U; i++) {
      for (j = 0; j < V; j++) {
        var u = (i + (rand() - 0.5) * o.jitter) * 2 * Math.PI / U;
        var v = (j + (rand() - 0.5) * o.jitter) * 2 * Math.PI / V;
        var r = o.tube * (1 + (rand() - 0.5) * o.bumps);
        var w = o.radius + r * Math.cos(v);
        k = i * V + j;
        px[k] = w * Math.cos(u);
        py[k] = r * Math.sin(v);
        pz[k] = w * Math.sin(u);
        sine[k] = Math.sin(v);
      }
    }

    // the frosting's edge wanders around the ring, smoothed so it doesn't zigzag
    var rough = [], edge = [];
    for (i = 0; i < U; i++) {
      rough.push(o.frostingLine + (rand() - 0.5) * 0.6);
    }

    for (i = 0; i < U; i++) {
      edge.push((rough[(i + U - 1) % U] + 2 * rough[i] + rough[(i + 1) % U]) / 4);
    }

    var faces = [];
    function face(a, b, c, i) {
      // wound so that (b - a) x (c - a) points outward, for culling
      var cx = (px[a] + px[b] + px[c]) / 3, cy = (py[a] + py[b] + py[c]) / 3, cz = (pz[a] + pz[b] + pz[c]) / 3;
      var ringLength = Math.hypot(cx, cz) || 1;
      var ox = cx - cx / ringLength * o.radius, oy = cy, oz = cz - cz / ringLength * o.radius;
      var ux = px[b] - px[a], uy = py[b] - py[a], uz = pz[b] - pz[a];
      var vx = px[c] - px[a], vy = py[c] - py[a], vz = pz[c] - pz[a];
      var nx = uy * vz - uz * vy, ny = uz * vx - ux * vz, nz = ux * vy - uy * vx;
      if (nx * ox + ny * oy + nz * oz < 0) {
        var swap = b; b = c; c = swap;
        nx = -nx; ny = -ny; nz = -nz;
      }

      var length = Math.hypot(nx, ny, nz) || 1;
      var frosting = (sine[a] + sine[b] + sine[c]) / 3 > edge[i];
      faces.push({
        a: a, b: b, c: c,
        nx: nx / length, ny: ny / length, nz: nz / length,
        frosting: frosting,
        shade: 1 + (rand() - 0.5) * 2 * o.shade
      });
    }

    for (i = 0; i < U; i++) {
      for (j = 0; j < V; j++) {
        var a = i * V + j;
        var b = ((i + 1) % U) * V + j;
        var c = ((i + 1) % U) * V + (j + 1) % V;
        var d = i * V + (j + 1) % V;
        if (rand() < 0.5) {
          face(a, b, c, i);
          face(a, c, d, i);
        } else {
          face(a, b, d, i);
          face(b, c, d, i);
        }
      }
    }

    return { count: count, px: px, py: py, pz: pz, faces: faces };
  }

  /* ---------- the donut ---------- */
  function Donut(canvas, options) {
    var self = this;
    this.canvas = canvas;
    this.options = Object.assign({}, defaults, options || {});
    var o = this.options;
    this.mesh = buildMesh(o);
    this.rx = new Float32Array(this.mesh.count);
    this.ry = new Float32Array(this.mesh.count);
    this.rz = new Float32Array(this.mesh.count);
    this.sx = new Float32Array(this.mesh.count);
    this.sy = new Float32Array(this.mesh.count);
    this.visible = [];
    this.matrix = multiply(rotation(0, 1, 0, o.turn * Math.PI / 180), rotation(1, 0, 0, o.tilt * Math.PI / 180));
    this.velocityX = 0;     // the fling: radians per second around the screen's horizontal
    this.velocityY = 0;     // and vertical
    this.ramp = 1;          // how much of the slow turn is on, 0..1, rising again after a drag
    this.dragging = false;
    this.dirty = true;
    this.frames = 0;
    this.onScreen = true;
    this.running = false;
    this.frame = null;
    this.lastTime = 0;

    this.resize = function () {
      var dpr = global.devicePixelRatio || 1;
      var w = Math.max(1, Math.round(canvas.clientWidth * dpr));
      var h = Math.max(1, Math.round(canvas.clientHeight * dpr));
      if (canvas.width !== w || canvas.height !== h) {
        canvas.width = w;
        canvas.height = h;
        // drawn at once: the new size has cleared the canvas, and the next animation frame
        // may be a while during a resize of the window
        self.render();
      } else {
        self.dirty = true;
      }
    };

    this.tick = function (time) {
      self.frame = null;
      if (!self.running) {
        return;
      }

      var dt = self.lastTime ? Math.min(0.05, (time - self.lastTime) / 1000) : 0;
      self.lastTime = time;
      self.advance(dt);
      if (self.dirty) {
        self.render();
      }

      self.frame = global.requestAnimationFrame(self.tick);
    };

    if (global.ResizeObserver) {
      this.resizeObserver = new ResizeObserver(this.resize);
      this.resizeObserver.observe(canvas);
    } else {
      global.addEventListener('resize', this.resize);
    }

    if (global.IntersectionObserver) {
      this.intersectionObserver = new IntersectionObserver(function (entries) {
        self.onScreen = entries[0].isIntersecting;
        self.lastTime = 0;
      });
      this.intersectionObserver.observe(canvas);
    }

    this.onVisibility = function () { self.lastTime = 0; };
    document.addEventListener('visibilitychange', this.onVisibility);

    if (o.interactive) {
      this.attachPointer();
    }

    this.resize();
    this.start();
  }

  Donut.prototype.start = function () {
    if (!this.running) {
      this.running = true;
      this.lastTime = 0;
      this.frame = global.requestAnimationFrame(this.tick);
    }
  };

  Donut.prototype.stop = function () {
    this.running = false;
    if (this.frame) {
      global.cancelAnimationFrame(this.frame);
      this.frame = null;
    }
  };

  Donut.prototype.dispose = function () {
    this.stop();
    if (this.resizeObserver) {
      this.resizeObserver.disconnect();
    } else {
      global.removeEventListener('resize', this.resize);
    }

    if (this.intersectionObserver) {
      this.intersectionObserver.disconnect();
    }

    document.removeEventListener('visibilitychange', this.onVisibility);
    if (this.detachPointer) {
      this.detachPointer();
    }
  };

  // the model's own axis (its Y) in screen space
  Donut.prototype.axis = function () {
    var m = this.matrix;
    return [m[1], m[4], m[7]];
  };

  Donut.prototype.rotateScreen = function (aroundX, aroundY) {
    if (aroundY) {
      this.matrix = multiply(rotation(0, 1, 0, aroundY), this.matrix);
    }

    if (aroundX) {
      this.matrix = multiply(rotation(1, 0, 0, aroundX), this.matrix);
    }

    this.dirty = true;
  };

  Donut.prototype.advance = function (dt) {
    if (!this.onScreen || document.hidden || this.dragging || dt <= 0) {
      return;
    }

    var o = this.options;
    // the fling slows down
    var decay = Math.exp(-dt * o.friction);
    this.velocityX *= decay;
    this.velocityY *= decay;
    if (Math.abs(this.velocityX) + Math.abs(this.velocityY) < 0.002) {
      this.velocityX = this.velocityY = 0;
    } else {
      this.rotateScreen(this.velocityX * dt, this.velocityY * dt);
    }

    // the slow turn comes back over a second after a drag
    this.ramp = Math.min(1, this.ramp + dt / 1.2);
    var eased = this.ramp * this.ramp * (3 - 2 * this.ramp);
    var spin = o.spin * eased * dt;
    if (spin) {
      var axis = o.spinAxis === 'screen' ? [0, 1, 0] : this.axis();
      this.matrix = multiply(rotation(axis[0], axis[1], axis[2], spin), this.matrix);
      this.dirty = true;
    }

    if (++this.frames % 90 === 0) {
      this.matrix = orthonormalize(this.matrix);
    }
  };

  // the donut's size and place on the canvas, in device pixels
  Donut.prototype.layout = function () {
    var canvas = this.canvas, o = this.options;
    var w = canvas.width, h = canvas.height;
    var reach = o.radius + o.tube;
    // pixels per ring radius at the middle depth, so that the donut's width is `fit` of the shorter side
    var scale = Math.min(w, h) * o.fit / 2 / reach;
    var cx = w / 2, cy = h / 2;
    if (o.anchor) {
      // the middle of the anchor's box, in the canvas's device pixels
      var own = canvas.getBoundingClientRect();
      var box = o.anchor.getBoundingClientRect();
      var ratio = own.width > 0 ? w / own.width : 1;
      cx = (box.left + box.width / 2 - own.left) * ratio;
      cy = (box.top + box.height / 2 - own.top) * ratio;
    }

    return {
      w: w, h: h,
      cx: cx + o.offset[0] * w,
      cy: cy + o.offset[1] * h,
      scale: scale,
      // the bounding sphere's outline, which the perspective makes a little bigger
      outline: scale * o.distance * reach / Math.sqrt(o.distance * o.distance - reach * reach)
    };
  };

  Donut.prototype.render = function () {
    this.dirty = false;
    var canvas = this.canvas, o = this.options, mesh = this.mesh, m = this.matrix;
    var ctx = canvas.getContext('2d');
    var box = this.layout();
    var w = box.w, h = box.h;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, w, h);

    var d = o.distance;
    var px = mesh.px, py = mesh.py, pz = mesh.pz;
    var rx = this.rx, ry = this.ry, rz = this.rz, sx = this.sx, sy = this.sy;
    var i, n = mesh.count;
    for (i = 0; i < n; i++) {
      var x = px[i], y = py[i], z = pz[i];
      var X = m[0] * x + m[1] * y + m[2] * z;
      var Y = m[3] * x + m[4] * y + m[5] * z;
      var Z = m[6] * x + m[7] * y + m[8] * z;
      rx[i] = X; ry[i] = Y; rz[i] = Z;
      var k = box.scale * d / (d - Z);
      sx[i] = box.cx + X * k;
      sy[i] = box.cy - Y * k;
    }

    if (o.shadow) {
      this.drawShadow(ctx, box);
    }

    // the light and the half vector of its highlights, in screen space
    var lx = o.light[0], ly = o.light[1], lz = o.light[2];
    var ll = Math.hypot(lx, ly, lz) || 1;
    lx /= ll; ly /= ll; lz /= ll;
    var hx = lx, hy = ly, hz = lz + 1;
    var hl = Math.hypot(hx, hy, hz);
    hx /= hl; hy /= hl; hz /= hl;

    var faces = mesh.faces, visible = this.visible;
    visible.length = 0;
    for (i = 0; i < faces.length; i++) {
      var f = faces[i];
      var a = f.a, b = f.b, c = f.c;
      // the outline's winding on screen says which way the facet faces (y points down on screen)
      var area = (sx[b] - sx[a]) * (sy[c] - sy[a]) - (sy[b] - sy[a]) * (sx[c] - sx[a]);
      if (area >= 0) {
        continue;
      }

      f.depth = rz[a] + rz[b] + rz[c];
      var nx = m[0] * f.nx + m[1] * f.ny + m[2] * f.nz;
      var ny = m[3] * f.nx + m[4] * f.ny + m[5] * f.nz;
      var nz = m[6] * f.nx + m[7] * f.ny + m[8] * f.nz;
      var lambert = Math.max(0, nx * lx + ny * ly + nz * lz);
      var value = (o.ambient + o.diffuse * lambert) * f.shade;
      var white = 0;
      if (f.frosting && o.gloss) {
        var spec = Math.max(0, nx * hx + ny * hy + nz * hz);
        white = o.gloss * spec * spec * spec * spec * spec * spec * spec * spec;
      }

      if (value > 1) {
        white += (value - 1) * 1.5;
        value = 1;
      }

      white = Math.min(1, white);
      var base = f.frosting ? o.frosting : o.dough;
      var r = base[0] * value, g = base[1] * value, bl = base[2] * value;
      r += (255 - r) * white; g += (255 - g) * white; bl += (255 - bl) * white;
      f.color = 'rgb(' + (r | 0) + ',' + (g | 0) + ',' + (bl | 0) + ')';
      visible.push(f);
    }

    visible.sort(function (p, q) { return p.depth - q.depth; });

    ctx.lineJoin = 'round';
    ctx.lineWidth = Math.max(1, (global.devicePixelRatio || 1) * 0.75);
    for (i = 0; i < visible.length; i++) {
      var face = visible[i];
      ctx.fillStyle = face.color;
      ctx.strokeStyle = face.color;
      ctx.beginPath();
      ctx.moveTo(sx[face.a], sy[face.a]);
      ctx.lineTo(sx[face.b], sy[face.b]);
      ctx.lineTo(sx[face.c], sy[face.c]);
      ctx.closePath();
      ctx.fill();
      // the same color over the edge, so that no hairline of the background shows between facets
      ctx.stroke();
    }
  };

  Donut.prototype.drawShadow = function (ctx, box) {
    var o = this.options;
    var reach = (o.radius + o.tube) * box.scale;
    var y = box.cy + reach * 1.02;
    var rx = reach * 1.05, ry = reach * 0.16;
    var gradient = ctx.createRadialGradient(box.cx, y, 0, box.cx, y, rx);
    gradient.addColorStop(0, 'rgba(40, 32, 20, 0.22)');
    gradient.addColorStop(0.6, 'rgba(40, 32, 20, 0.1)');
    gradient.addColorStop(1, 'rgba(40, 32, 20, 0)');
    ctx.save();
    ctx.translate(box.cx, y);
    ctx.scale(1, ry / rx);
    ctx.translate(-box.cx, -y);
    ctx.fillStyle = gradient;
    ctx.beginPath();
    ctx.arc(box.cx, y, rx, 0, 2 * Math.PI);
    ctx.fill();
    ctx.restore();
  };

  /* ---------- the pointer ---------- */
  Donut.prototype.hit = function (event) {
    var canvas = this.canvas;
    var rect = canvas.getBoundingClientRect();
    var dpr = canvas.width / (rect.width || 1);
    var box = this.layout();
    var x = (event.clientX - rect.left) * dpr - box.cx;
    var y = (event.clientY - rect.top) * dpr - box.cy;
    return x * x + y * y <= box.outline * box.outline * 1.1;
  };

  Donut.prototype.attachPointer = function () {
    var self = this, canvas = this.canvas, o = this.options;
    var lastX = 0, lastY = 0, lastTime = 0, pointerId = null;
    if (o.touch) {
      canvas.style.touchAction = 'none';
    }

    function down(event) {
      if (pointerId !== null || (event.pointerType === 'touch' && !o.touch) || (event.button && event.button !== 0)) {
        return;
      }

      if (!self.hit(event)) {
        return;
      }

      pointerId = event.pointerId;
      lastX = event.clientX; lastY = event.clientY; lastTime = performance.now();
      self.dragging = true;
      self.velocityX = self.velocityY = 0;
      canvas.style.cursor = 'grabbing';
      try { canvas.setPointerCapture(event.pointerId); } catch (e) { }
      event.preventDefault();
    }

    function move(event) {
      if (pointerId === null) {
        if (event.pointerType !== 'touch') {
          canvas.style.cursor = self.hit(event) ? 'grab' : '';
        }

        return;
      }

      if (event.pointerId !== pointerId) {
        return;
      }

      var now = performance.now();
      var dt = Math.max(1, now - lastTime) / 1000;
      var aroundX = (event.clientY - lastY) * o.dragSpeed;
      var aroundY = (event.clientX - lastX) * o.dragSpeed;
      self.rotateScreen(aroundX, aroundY);
      // the speed of the last moves, for the fling
      var cap = o.maxFling;
      self.velocityX = Math.max(-cap, Math.min(cap, 0.6 * (aroundX / dt) + 0.4 * self.velocityX));
      self.velocityY = Math.max(-cap, Math.min(cap, 0.6 * (aroundY / dt) + 0.4 * self.velocityY));
      lastX = event.clientX; lastY = event.clientY; lastTime = now;
      event.preventDefault();
    }

    function up(event) {
      if (event.pointerId !== pointerId) {
        return;
      }

      pointerId = null;
      self.dragging = false;
      // a pointer that stopped before letting go gives no fling
      if (performance.now() - lastTime > 120) {
        self.velocityX = self.velocityY = 0;
      }

      self.ramp = 0;
      self.lastTime = 0;
      canvas.style.cursor = event.pointerType !== 'touch' && self.hit(event) ? 'grab' : '';
      try { canvas.releasePointerCapture(event.pointerId); } catch (e) { }
    }

    canvas.addEventListener('pointerdown', down);
    canvas.addEventListener('pointermove', move);
    canvas.addEventListener('pointerup', up);
    canvas.addEventListener('pointercancel', up);
    this.detachPointer = function () {
      canvas.removeEventListener('pointerdown', down);
      canvas.removeEventListener('pointermove', move);
      canvas.removeEventListener('pointerup', up);
      canvas.removeEventListener('pointercancel', up);
    };
  };

  global.Donut = Donut;
})(window);
