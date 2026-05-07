import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─────────────────────────────────────────────────────────────────────────────
// TYPES — JSON shape written by FlagJsonWriter plugin
// ─────────────────────────────────────────────────────────────────────────────
interface FlagJsonEntry {
  id:       string;
  name:     string;
  iconUrl:  string;
  isActive: boolean;
}

// ─────────────────────────────────────────────────────────────────────────────
// CSS — from :root in your OffenderFlags web resource
// ─────────────────────────────────────────────────────────────────────────────
const CSS = {
  font:        "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
  iconBg:      "#F5F4F0",
  iconBgHover: "#ECEAE3",
  iconBorder:  "#E8E6DF",
  iconBorderHover: "#C8C6C4",
  iconColor:   "#605E5C",
  textHint:    "#A19F9D",
  tooltipBg:   "#201F1E",   // --text-primary equivalent, works light+dark
  tooltipText: "#FFFFFF",
};

// ─────────────────────────────────────────────────────────────────────────────
// FALLBACK SVG — identical to FALLBACK_SVG in your web resource
// ─────────────────────────────────────────────────────────────────────────────
const FALLBACK_SVG = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor"
  stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"
  style="width:13px;height:13px;display:block">
  <path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/>
  <line x1="4" y1="22" x2="4" y2="15"/>
</svg>`;

// ─────────────────────────────────────────────────────────────────────────────
// MODULE-LEVEL SVG CACHE
// Shared across every row — each unique iconUrl fetched only once per session.
// Mirrors _svgCache in your web resource.
// ─────────────────────────────────────────────────────────────────────────────
const _svgCache    = new Map<string, string>();
const _svgFetching = new Map<string, Promise<string>>();

function resolveUrl(url: string): string {
  if (!url) return "";
  return url.startsWith("http")
    ? url
    : window.location.origin + (url.startsWith("/") ? url : "/" + url);
}

function fetchSvg(url: string): Promise<string> {
  if (!url) return Promise.resolve(FALLBACK_SVG);
  const resolved = resolveUrl(url);
  if (_svgCache.has(resolved))    return Promise.resolve(_svgCache.get(resolved)!);
  if (_svgFetching.has(resolved)) return _svgFetching.get(resolved)!;

  const job = fetch(resolved, { credentials: "same-origin" })
    .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
    .then(text => {
      const match = text.match(/<svg[\s\S]*<\/svg>/i);
      const svg   = match ? match[0] : FALLBACK_SVG;
      const sized = svg.replace(/<svg/, '<svg style="width:13px;height:13px;display:block" ');
      _svgCache.set(resolved, sized);
      return sized;
    })
    .catch(err => {
      console.warn("[ClientFlagViewer] SVG fetch failed:", resolved, (err as Error).message);
      _svgCache.set(resolved, FALLBACK_SVG);
      return FALLBACK_SVG;
    })
    .finally(() => { _svgFetching.delete(resolved); });

  _svgFetching.set(resolved, job);
  return job;
}

function prefetchIcons(flags: FlagJsonEntry[]): void {
  [...new Set(flags.map(f => f.iconUrl).filter(Boolean))].forEach(u => fetchSvg(u));
}

// ─────────────────────────────────────────────────────────────────────────────
// TOOLTIP CLIPPING DETECTION
//
// PCF controls in D365 views run inside a per-cell iframe. A styled tooltip
// using position:absolute can be clipped at the iframe boundary.
//
// Strategy — detect clipping on first show, then permanently downgrade
// that instance to the native browser title attribute:
//   1. Show the styled tooltip.
//   2. After one rAF (paint), check if its bounding rect is fully inside
//      the iframe viewport (window.innerHeight).
//   3. If clipped → hide the styled tooltip, set useNativeTitle = true,
//      fall back to title="" on every button in this instance.
//
// The native title tooltip is rendered by the browser outside the iframe
// so it never clips. It looks less polished but always works.
// ─────────────────────────────────────────────────────────────────────────────
let _activeTooltip: HTMLElement | null = null;

function dismissTooltip(): void {
  if (_activeTooltip) {
    _activeTooltip.style.display = "none";
    _activeTooltip = null;
  }
}

// Global dismiss on click outside — registered once
if (typeof document !== "undefined") {
  document.addEventListener("click", (e: MouseEvent) => {
    if (_activeTooltip && !_activeTooltip.parentElement!.contains(e.target as Node)) {
      dismissTooltip();
    }
  }, { capture: true });
}

/**
 * Returns true if el is fully visible within the current window viewport.
 * Used to detect iframe clipping after the tooltip is shown.
 */
function isFullyVisible(el: HTMLElement): boolean {
  const r = el.getBoundingClientRect();
  return (
    r.top    >= 0 &&
    r.left   >= 0 &&
    r.bottom <= window.innerHeight &&
    r.right  <= window.innerWidth
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PCF CONTROL
// ─────────────────────────────────────────────────────────────────────────────
export class ClientFlagViewer
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _root:           HTMLDivElement;
  private _lastRaw:        string  = "__UNINIT__";
  private _destroyed:      boolean = false;

  // Once set to true for this instance, all icons use title="" instead of
  // the styled tooltip — permanently, for the lifetime of this cell render.
  private _useNativeTitle: boolean = false;

  // All buttons rendered in the current row — needed so we can bulk-apply
  // title="" if clipping is detected on the first tooltip show.
  private _buttons: Array<{ btn: HTMLElement; name: string }> = [];

  // ── init ──────────────────────────────────────────────────────────────────
  public init(
    context:             ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state:              ComponentFramework.Dictionary,
    container:           HTMLDivElement
  ): void {
    this._root = document.createElement("div");
    this._root.style.cssText =
      `display:flex;flex-wrap:wrap;gap:4px;align-items:center;` +
      `padding:2px 0;font-family:${CSS.font};`;
    container.appendChild(this._root);
  }

  // ── updateView ────────────────────────────────────────────────────────────
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    const raw = (context.parameters.flagsJson.raw || "").trim();
    if (raw === this._lastRaw) return;
    this._lastRaw = raw;

    if (!raw) { this._renderEmpty(); return; }

    let allFlags: FlagJsonEntry[];
    try { allFlags = JSON.parse(raw) as FlagJsonEntry[]; }
    catch { this._renderEmpty(); return; }

    if (!Array.isArray(allFlags) || allFlags.length === 0) { this._renderEmpty(); return; }

    prefetchIcons(allFlags);

    const active = allFlags.filter(f => f.isActive === true);
    if (active.length === 0) { this._renderEmpty(); return; }

    this._renderIcons(active);
  }

  // ── _renderEmpty ──────────────────────────────────────────────────────────
  private _renderEmpty(): void {
    this._root.innerHTML = "";
    const span = document.createElement("span");
    span.style.cssText = `font-size:11px;color:${CSS.textHint};font-family:${CSS.font};`;
    span.textContent = "—";
    this._root.appendChild(span);
  }

  // ── _renderIcons ──────────────────────────────────────────────────────────
  // Renders one 26×26 icon button per active flag.
  //
  // Normal path: styled tooltip appears on click (positioned above the icon).
  // Fallback path: if a tooltip is detected as clipped by the iframe boundary
  // on its first show, the entire instance permanently switches to native
  // browser title="" tooltips (shown on hover, rendered outside the iframe).
  private _renderIcons(flags: FlagJsonEntry[]): void {
    this._root.innerHTML = "";
    this._buttons = [];

    flags.forEach(flag => {
      // ── Wrapper ───────────────────────────────────────────────────────────
      const wrap = document.createElement("span");
      wrap.style.cssText = `position:relative;display:inline-flex;flex-shrink:0;`;

      // ── Icon button ───────────────────────────────────────────────────────
      const btn = document.createElement("span");
      btn.setAttribute("role", "button");
      btn.setAttribute("tabindex", "0");
      btn.setAttribute("aria-label", flag.name);

      // If a previous tooltip already triggered the fallback on this instance,
      // apply title immediately — no styled tooltip needed
      if (this._useNativeTitle) {
        btn.setAttribute("title", flag.name);
      }

      btn.style.cssText =
        `width:26px;height:26px;border-radius:6px;` +
        `display:inline-flex;align-items:center;justify-content:center;` +
        `background:${CSS.iconBg};border:.5px solid ${CSS.iconBorder};` +
        `color:${CSS.iconColor};cursor:pointer;flex-shrink:0;` +
        `transition:background .12s,border-color .12s;`;

      btn.addEventListener("mouseenter", () => {
        btn.style.background  = CSS.iconBgHover;
        btn.style.borderColor = CSS.iconBorderHover;
      });
      btn.addEventListener("mouseleave", () => {
        btn.style.background  = CSS.iconBg;
        btn.style.borderColor = CSS.iconBorder;
      });

      // ── Styled tooltip ────────────────────────────────────────────────────
      const tip = document.createElement("span");
      tip.textContent = flag.name;
      tip.style.cssText =
        `display:none;position:absolute;` +
        `bottom:calc(100% + 6px);left:50%;transform:translateX(-50%);` +
        `background:${CSS.tooltipBg};color:${CSS.tooltipText};` +
        `font-size:11px;font-weight:500;font-family:${CSS.font};` +
        `white-space:nowrap;padding:4px 8px;border-radius:6px;` +
        `pointer-events:none;z-index:9999;`;

      const caret = document.createElement("span");
      caret.style.cssText =
        `position:absolute;top:100%;left:50%;transform:translateX(-50%);` +
        `border:4px solid transparent;border-top-color:${CSS.tooltipBg};`;
      tip.appendChild(caret);

      // ── Click handler ─────────────────────────────────────────────────────
      const handleClick = (e: Event) => {
        e.stopPropagation();

        // Native title fallback is active — nothing to do, browser handles it
        if (this._useNativeTitle) return;

        if (_activeTooltip === tip) {
          dismissTooltip();
          return;
        }

        dismissTooltip();
        tip.style.display = "block";
        _activeTooltip    = tip;

        // ── Clipping check ────────────────────────────────────────────────
        // After one animation frame the tooltip is painted — check if it
        // is fully visible within the iframe viewport.
        requestAnimationFrame(() => {
          if (this._destroyed) return;

          if (!isFullyVisible(tip)) {
            // Clipped — dismiss styled tooltip and permanently switch this
            // entire instance to native title="" tooltips.
            dismissTooltip();
            this._useNativeTitle = true;

            // Apply title to every button already rendered in this row
            this._buttons.forEach(({ btn: b, name }) => {
              b.setAttribute("title", name);
            });
          }
        });
      };

      btn.addEventListener("click",   handleClick);
      btn.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleClick(e); }
      });

      // ── SVG icon ──────────────────────────────────────────────────────────
      const resolved  = resolveUrl(flag.iconUrl);
      const cachedSvg = resolved ? _svgCache.get(resolved) : null;
      btn.innerHTML   = cachedSvg || FALLBACK_SVG;

      if (flag.iconUrl && !cachedSvg) {
        fetchSvg(flag.iconUrl).then(svg => {
          if (!this._destroyed) {
            btn.innerHTML = svg;
            wrap.appendChild(tip); // re-attach tip after innerHTML clobber
          }
        });
      }

      this._buttons.push({ btn, name: flag.name });
      wrap.appendChild(btn);
      wrap.appendChild(tip);
      this._root.appendChild(wrap);
    });
  }

  // ── getOutputs ────────────────────────────────────────────────────────────
  public getOutputs(): IOutputs { return {}; }

  // ── destroy ───────────────────────────────────────────────────────────────
  public destroy(): void {
    this._destroyed = true;
    dismissTooltip();
  }
}
