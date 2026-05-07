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
// PAGE-LEVEL TOOLTIP MANAGER
// Only one tooltip visible at a time across all PCF instances on the page.
// Click the same icon again → dismiss. Click elsewhere → dismiss.
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

// ─────────────────────────────────────────────────────────────────────────────
// PCF CONTROL
// ─────────────────────────────────────────────────────────────────────────────
export class ClientFlagViewer
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _root:      HTMLDivElement;
  private _lastRaw:   string  = "__UNINIT__";
  private _destroyed: boolean = false;

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
  // A tooltip showing the flag name appears on click; dismissed on second
  // click, on click elsewhere, or when another icon is clicked.
  private _renderIcons(flags: FlagJsonEntry[]): void {
    this._root.innerHTML = "";

    flags.forEach(flag => {
      // ── Wrapper — positions the tooltip relative to the icon ──────────────
      const wrap = document.createElement("span");
      wrap.style.cssText = `position:relative;display:inline-flex;flex-shrink:0;`;

      // ── Icon button ───────────────────────────────────────────────────────
      const btn = document.createElement("span");
      btn.setAttribute("role", "button");
      btn.setAttribute("tabindex", "0");
      btn.setAttribute("aria-label", flag.name);
      btn.style.cssText =
        `width:26px;height:26px;border-radius:6px;` +
        `display:inline-flex;align-items:center;justify-content:center;` +
        `background:${CSS.iconBg};border:.5px solid ${CSS.iconBorder};` +
        `color:${CSS.iconColor};cursor:pointer;flex-shrink:0;` +
        `transition:background .12s,border-color .12s;`;

      // Hover styles via mouseenter/mouseleave (no stylesheet injection)
      btn.addEventListener("mouseenter", () => {
        btn.style.background   = CSS.iconBgHover;
        btn.style.borderColor  = CSS.iconBorderHover;
      });
      btn.addEventListener("mouseleave", () => {
        btn.style.background   = CSS.iconBg;
        btn.style.borderColor  = CSS.iconBorder;
      });

      // ── Tooltip ───────────────────────────────────────────────────────────
      // Positioned above the icon. Hidden by default, toggled on click.
      const tip = document.createElement("span");
      tip.textContent = flag.name;
      tip.style.cssText =
        `display:none;position:absolute;` +
        `bottom:calc(100% + 6px);left:50%;transform:translateX(-50%);` +
        `background:${CSS.tooltipBg};color:${CSS.tooltipText};` +
        `font-size:11px;font-weight:500;font-family:${CSS.font};` +
        `white-space:nowrap;padding:4px 8px;border-radius:6px;` +
        `pointer-events:none;z-index:9999;`;

      // Caret arrow pointing down toward the icon
      const caret = document.createElement("span");
      caret.style.cssText =
        `position:absolute;top:100%;left:50%;transform:translateX(-50%);` +
        `border:4px solid transparent;border-top-color:${CSS.tooltipBg};`;
      tip.appendChild(caret);

      // ── Click handler ─────────────────────────────────────────────────────
      const handleClick = (e: Event) => {
        e.stopPropagation();

        if (_activeTooltip === tip) {
          // Same icon clicked again — dismiss
          dismissTooltip();
          return;
        }

        // Dismiss any other open tooltip first
        dismissTooltip();

        tip.style.display = "block";
        _activeTooltip    = tip;
      };

      btn.addEventListener("click",   handleClick);
      btn.addEventListener("keydown", (e: KeyboardEvent) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleClick(e); }
      });

      // ── Populate SVG icon ─────────────────────────────────────────────────
      const resolved  = resolveUrl(flag.iconUrl);
      const cachedSvg = resolved ? _svgCache.get(resolved) : null;
      btn.innerHTML   = cachedSvg || FALLBACK_SVG;

      if (flag.iconUrl && !cachedSvg) {
        fetchSvg(flag.iconUrl).then(svg => {
          if (!this._destroyed) btn.innerHTML = svg;
          // Re-append tooltip since innerHTML replaced it
          wrap.appendChild(tip);
        });
      }

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
