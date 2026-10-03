'use strict';
// ===============================================================
// load-client.js
// Loads the browser-side scripts in a vm context with minimal DOM /
// storage / jQuery stubs, so the pure helpers can be unit tested.
// ===============================================================

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.join(__dirname, '..', '..');

/**
 * A localStorage/sessionStorage stand-in.
 *
 * Backed by a Proxy so it behaves like the real Storage object: stored keys are
 * visible to `Object.keys(localStorage)` / `for..in`, which core.js relies on
 * when it sweeps cache keys on logout. A plain object with only methods would
 * silently return an empty key list and hide real cleanup bugs.
 */
function makeStorage() {
  const map = new Map();
  const methods = {
    getItem: (k) => (map.has(String(k)) ? map.get(String(k)) : null),
    setItem: (k, v) => { map.set(String(k), String(v)); },
    removeItem: (k) => { map.delete(String(k)); },
    clear: () => map.clear(),
    key: (i) => { const keys = [...map.keys()]; return i < keys.length ? keys[i] : null; },
    _dump: () => Object.fromEntries(map),
  };
  return new Proxy(methods, {
    get(target, prop) {
      if (prop in target) return target[prop];
      if (typeof prop === 'string' && map.has(prop)) return map.get(prop);
      return undefined;
    },
    has(target, prop) {
      return (prop in target) || (typeof prop === 'string' && map.has(prop));
    },
    ownKeys() { return [...map.keys()]; },
    getOwnPropertyDescriptor(target, prop) {
      if (typeof prop !== 'string' || !map.has(prop)) {
        return Reflect.getOwnPropertyDescriptor(target, prop);
      }
      return { value: map.get(prop), writable: true, enumerable: true, configurable: true };
    },
  });
}

/** A minimal element stub that records class/text mutations. */
class FakeEl {
  constructor(tag = 'div', id = '') {
    this.tagName = String(tag).toUpperCase();
    this.id = id;
    this._classes = new Set();
    this._html = '';
    this._text = '';
    this.value = '';
    this.disabled = false;
    this.style = {};
    this.dataset = {};
    this.children = [];
    this._listeners = {};
    this._attrs = {};
    this._qCache = {};
  }
  classList = {
    add: (...c) => c.forEach(x => this._classes.add(x)),
    remove: (...c) => c.forEach(x => this._classes.delete(x)),
    toggle: (c, on) => (on ? this._classes.add(c) : this._classes.delete(c)),
    contains: (c) => this._classes.has(c),
  };
  get innerHTML() { return this._html; }
  set innerHTML(v) { this._html = String(v); this.children = []; this._qCache = {}; }
  insertAdjacentHTML(_pos, html) { this._html = String(html) + this._html; this.children = []; this._qCache = {}; }
  scrollIntoView() {}
  get textContent() { return this._text; }
  set textContent(v) { this._text = String(v); }
  addEventListener(type, fn, opts) {
    // نحاكي حقيقي المتصفح: { signal } يفصل المستمع عند abort.
    if (opts && opts.signal) {
      if (opts.signal.aborted) return;
      opts.signal.addEventListener('abort', () => this.removeEventListener(type, fn), { once: true });
    }
    (this._listeners[type] ||= []).push(fn);
  }
  removeEventListener(type, fn) {
    const l = this._listeners[type]; if (!l) return;
    const i = l.indexOf(fn); if (i >= 0) l.splice(i, 1);
  }
  dispatch(type, ev = {}) { (this._listeners[type] || []).forEach(fn => fn(ev)); }
  querySelector() { return null; }
  /**
   * Class selectors are resolved against `innerHTML`.
   *
   * Rows rendered with insertAdjacentHTML bind their listeners through
   * querySelectorAll, so without this a button that exists in the markup is
   * invisible to a test and its handler can never be exercised. Only `.class`
   * is supported, which is all the row rendering uses. Results are cached per
   * selector and dropped whenever innerHTML changes, so a test that re-queries
   * gets the same instances the page bound its listeners to.
   */
  querySelectorAll(selector) {
    const cls = /^\.([\w-]+)$/.exec(String(selector).trim());
    if (!cls) return [];
    if (this._qCache[selector]) return this._qCache[selector];
    const found = [];
    for (const tag of String(this._html).match(/<[a-zA-Z][^>]*>/g) || []) {
      const classAttr = /class\s*=\s*"([^"]*)"/.exec(tag);
      if (!classAttr || !classAttr[1].split(/\s+/).includes(cls[1])) continue;
      const el = new FakeEl((/^<([a-zA-Z0-9]+)/.exec(tag) || [])[1] || 'div');
      el._attrs = {};
      for (const a of tag.matchAll(/([\w-]+)\s*=\s*"([^"]*)"/g)) el._attrs[a[1]] = a[2];
      for (const a of Object.keys(el._attrs)) {
        if (a.startsWith('data-')) {
          el.dataset[a.slice(5).replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = el._attrs[a];
        }
      }
      found.push(el);
    }
    this._qCache[selector] = found;
    return found;
  }
  appendChild(c) { this.children.push(c); return c; }
  remove() {}
  focus() {}
  click() {}
  reset() {}
  closest() { return null; }
  getAttribute(name) {
    return this._attrs[name] !== undefined ? this._attrs[name] : null;
  }
  setAttribute() {}
  removeAttribute() {}
  contains() { return false; }
}

/**
 * Load a client script with browser globals stubbed.
 * @param {string} file  filename inside the project root
 * @returns {{ctx: object, sandbox: object, evalIn: Function, window: object, doc: object}}
 */
function loadClient(file, opts = {}) {
  const source = fs.readFileSync(path.join(ROOT, file), 'utf8');

  const localStorage = makeStorage();
  const sessionStorage = makeStorage();
  const elements = new Map();

  const document = {
    getElementById: (id) => {
      if (!elements.has(id)) elements.set(id, new FakeEl('div', id));
      return elements.get(id);
    },
    querySelector: () => null,
    querySelectorAll: () => [],
    createElement: (tag) => new FakeEl(tag),
    createElementNS: (ns, tag) => new FakeEl(tag),
    addEventListener: () => {},
    removeEventListener: () => {},
    body: new FakeEl('body'),
    documentElement: new FakeEl('html'),
    cookie: '',
  };

  const noop = () => {};
  const jqStub = () => ({
    on: noop, off: noop, ready: noop, val: () => '', html: () => jqStub(),
    text: () => '', addClass: noop, removeClass: noop, toggleClass: noop,
    attr: () => '', data: () => '', append: noop, prepend: noop, remove: noop,
    each: noop, click: noop, trigger: noop, css: noop, find: () => jqStub(),
    closest: () => jqStub(), parent: () => jqStub(), children: () => jqStub(),
    show: noop, hide: noop, prop: () => '', hasClass: () => false, length: 0,
  });

  // Real window-level event dispatch so page modules that listen for
  // events (e.g. core.js's spaViewRevisited) can be exercised.
  const winListeners = new Map();
  const addWinListener = (type, fn, opts) => {
    if (opts && opts.signal) {
      if (opts.signal.aborted) return;
      opts.signal.addEventListener('abort', () => removeWinListener(type, fn), { once: true });
    }
    if (!winListeners.has(type)) winListeners.set(type, []);
    winListeners.get(type).push(fn);
  };
  const removeWinListener = (type, fn) => {
    const l = winListeners.get(type);
    if (!l) return;
    const i = l.indexOf(fn);
    if (i >= 0) l.splice(i, 1);
  };
  const dispatchWin = (type, detail) => {
    const ev = { type, detail };
    (winListeners.get(type) || []).slice().forEach(fn => fn(ev));
    return true;
  };

  const sandbox = {
    console,
    // Timers are stubbed so long-lived intervals in core.js cannot keep the
    // test process alive. Tests that need timing should use evalIn + promises.
    setTimeout: (fn, ms) => setTimeout(fn, Math.min(Number(ms) || 0, 50)),
    clearTimeout,
    setInterval: () => 0,
    clearInterval: () => {},
    Date, JSON, Math, Number, String, Object, Array, Boolean, RegExp,
    Error, TypeError, RangeError, Promise, Map, Set, WeakMap, Symbol,
    isNaN, isFinite, parseInt, parseFloat,
    AbortController, AbortSignal,
    encodeURIComponent, decodeURIComponent, encodeURI, decodeURI,
    URLSearchParams, URL, fetch: async () => { throw new Error('fetch not stubbed'); },
    localStorage, sessionStorage,
    document,
    navigator: { userAgent: 'node-test', language: 'ar', serviceWorker: null },
    location: { href: 'https://example.test/', origin: 'https://example.test', search: '', reload: noop },
    history: { pushState: noop, replaceState: noop },
    alert: noop, confirm: () => false, prompt: () => null,
    btoa: (s) => Buffer.from(String(s), 'utf8').toString('base64'),
    atob: (s) => Buffer.from(String(s), 'base64').toString('utf8'),
    jQuery: jqStub, $: jqStub,
    bootstrap: { Modal: function () {}, Tooltip: function () {}, Collapse: function () {} },
    Chart: function () { return { destroy: noop, update: noop }; },
    Swal: { fire: async () => ({ isConfirmed: false }), mixin: () => ({ fire: async () => ({}) }) },
    localStorageReady: true,
    addEventListener: addWinListener,
    removeEventListener: removeWinListener,
    dispatchEvent: (ev) => {
      const type = typeof ev === 'string' ? ev : ev && ev.type;
      const detail = typeof ev === 'string' ? undefined : ev && ev.detail;
      return dispatchWin(type, detail);
    },
    matchMedia: () => ({ matches: false, addEventListener: noop, removeEventListener: noop, addListener: noop }),
    innerWidth: 1280, innerHeight: 800,
    scrollTo: noop,
    CustomEvent: function (type, init) { this.type = type; this.detail = init && init.detail; },
    Event: function (type) { this.type = type; },
    getComputedStyle: () => ({ getPropertyValue: () => '' }),
    requestAnimationFrame: (fn) => setTimeout(fn, 0),
    cancelAnimationFrame: (id) => clearTimeout(id),
  };
  sandbox.window = sandbox;
  sandbox.self = sandbox;
  sandbox.globalThis = sandbox;

  const context = vm.createContext(sandbox);
  // Let callers install stubs (registerView, apiGet, ...) BEFORE the page
  // module runs its top-level registerView() call.
  if (typeof opts.setup === 'function') opts.setup(sandbox, { document, elements, dispatchWindowEvent: dispatchWin });
  new vm.Script(source, { filename: file }).runInContext(context);

  return {
    sandbox, context, document, elements,
    localStorage, sessionStorage,
    evalIn: (expr) => vm.runInContext(String(expr), context),
    dispatchWindowEvent: dispatchWin,
  };
}

module.exports = { loadClient, FakeEl, makeStorage, ROOT };
