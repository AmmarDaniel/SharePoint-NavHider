import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { SPHttpClient } from "@microsoft/sp-http";

const LOG_SOURCE = "NavHiderApplicationCustomizer";

export interface INavHiderProperties {
  // If true or undefined, only hide on the site's Welcome page. If false, hide on all pages.
  hideOnHomeOnly?: boolean;
  // Optional additional server-relative paths to hide on (e.g., ['/SitePages/Landing.aspx'])
  extraHidePaths?: string[];
}

export default class NavHiderApplicationCustomizer extends BaseApplicationCustomizer<INavHiderProperties> {
  private _styleEl: HTMLStyleElement | null = null;
  private _appliedForPath: string | null = null;
  private _observer: MutationObserver | null = null;

  public async onInit(): Promise<void> {
    // Heuristic to decide if we should pre-hide to avoid flash
    const path = window.location.pathname.toLowerCase();
    const webRelRaw = this.context.pageContext.web.serverRelativeUrl || "";
    const webRel = webRelRaw.toLowerCase().replace(/\/$/, "");
    const isRoot = path === webRel || path === `${webRel}/`;

    const extra = (this.properties.extraHidePaths || []).map((p) =>
      p.toLowerCase()
    );
    const cacheKey = `NavHider:Welcome:${
      this.context.pageContext.web.id?.toString() ??
      this.context.pageContext.web.absoluteUrl
    }`;
    const cachedWelcome = (
      sessionStorage.getItem(cacheKey) || ""
    ).toLowerCase();

    const isCachedWelcome =
      cachedWelcome && path === `${webRel}/${cachedWelcome}`;
    const isExtra = extra.includes(path);

    // Prepend style and set rules early ONLY when likely the home page
    if (
      isRoot ||
      isCachedWelcome ||
      isExtra ||
      this.properties.hideOnHomeOnly === false
    ) {
      this._ensureStyle();
      this._setCss(this._getHideCss());
    }

    await this._applyIfNeeded();

    this.context.application.navigatedEvent.add(this, () =>
      this._applyIfNeeded()
    );
    return Promise.resolve();
  }

  private async _applyIfNeeded(): Promise<void> {
    try {
      const currentPath = window.location.pathname.toLowerCase();
      if (this._appliedForPath === currentPath) return;

      // --- EDIT MODE GUARD ---
      const qs = new URLSearchParams(window.location.search);
      const isEditMode =
        qs.get("mode")?.toLowerCase() === "edit" ||
        qs.get("DisplayMode")?.toLowerCase() === "edit";

      if (isEditMode) {
        this._stopObserving();
        this._clearCss(); // show header in edit
        this._appliedForPath = null;
        return;
      }

      const hideOnHomeOnly = this.properties.hideOnHomeOnly !== false; // default true
      const extraPaths = (this.properties.extraHidePaths || []).map((p) =>
        p.toLowerCase()
      );

      // If user added explicit extra paths, honor them
      let shouldHide = extraPaths.includes(currentPath);

      if (!shouldHide && hideOnHomeOnly) {
        const webUrl = this.context.pageContext.web.absoluteUrl;
        const webRelRaw = this.context.pageContext.web.serverRelativeUrl || "";
        const webRel = webRelRaw.toLowerCase().replace(/\/$/, "");

        // --- READ WELCOME PAGE (cached) ---
        const cacheKey = `NavHider:Welcome:${
          this.context.pageContext.web.id?.toString() ?? webUrl
        }`;
        let welcomeRelCached = sessionStorage.getItem(cacheKey) ?? "";

        if (!welcomeRelCached) {
          const resp = await this.context.spHttpClient.get(
            `${webUrl}/_api/web/rootfolder?$select=WelcomePage`,
            SPHttpClient.configurations.v1
          );
          if (resp.ok) {
            const json = (await resp.json()) as {
              WelcomePage?: string;
              d?: { WelcomePage?: string };
            };
            const welcomeRelRaw = (json?.WelcomePage ??
              json?.d?.WelcomePage ??
              "") as string;
            welcomeRelCached = welcomeRelRaw.toLowerCase().replace(/^\//, "");
            if (welcomeRelCached)
              sessionStorage.setItem(cacheKey, welcomeRelCached);
          } else {
            Log.warn(
              LOG_SOURCE,
              `Failed to read WelcomePage. Status: ${resp.status}`
            );
          }
        }

        if (welcomeRelCached) {
          const homePath = `${webRel}/${welcomeRelCached}`;
          const isSiteRoot =
            currentPath === webRel || currentPath === `${webRel}/`;
          const isWelcomePage = currentPath === homePath;
          shouldHide = isSiteRoot || isWelcomePage;
        }
      }

      if (!hideOnHomeOnly || shouldHide) {
        // Keep hide rules, start observing for sticky header insertions
        this._setCss(this._getHideCss());
        this._startObserving();
        this._appliedForPath = currentPath;
      } else {
        // Not a page we should hide on -> remove CSS & observer
        this._stopObserving();
        this._clearCss();
        this._unhideAll();
        this._appliedForPath = null;
      }
    } catch (e) {
      Log.error(LOG_SOURCE, e as Error);
      this._stopObserving();
      this._clearCss();
      this._appliedForPath = null;
    }
  }

  // ---------------- CSS / Style handling ----------------

  private _ensureStyle(): void {
    if (!this._styleEl) {
      this._styleEl = document.createElement("style");
      this._styleEl.setAttribute("data-navhider", "true");
      // Prepend so it applies before most site CSS (reduces flash)
      document.head.prepend(this._styleEl);
    }
  }

  private _setCss(css: string): void {
    this._ensureStyle();
    this._styleEl!.textContent = css;
  }

  private _clearCss(): void {
    if (this._styleEl) {
      this._styleEl.textContent = ""; // keep node but no rules
    }
  }

  private _getHideCss(): string {
    return `
/* Hide entire site header host (covers full + sticky) */
div#spSiteHeader,
div[data-sp-feature-tag="Site header host"],
[aria-label="SharePoint Site Header"] {
  display: none !important;
  height: 0 !important;
  min-height: 0 !important;
  overflow: hidden !important;
  padding: 0 !important;
  margin: 0 !important;
  border: 0 !important;
}

/* Fallback inner surfaces */
div[data-automationid="SiteHeader"],
div[data-automationid="HorizontalNav"],
div[data-automationid="MegaMenu"],
div[data-automationid="StickyHeader"],
div[data-automationid="StickyTopHeader"],
div[data-automation-id="SiteHeader"],
div[data-automation-id="HorizontalNav"],
div[data-automation-id="MegaMenu"],
div[data-automation-id="StickyHeader"],
div[data-automation-id="StickyTopHeader"] {
  display: none !important;
}

/* Lift content to the top */
main[role="main"],
div[data-automationid="ContentScrollRegion"],
div[data-automationid="CanvasZone"],
div[data-automationid="CanvasSection"],
div[data-automation-id="ContentScrollRegion"],
div[data-automation-id="CanvasZone"],
div[data-automation-id="CanvasSection"] {
  margin-top: 0 !important;
  padding-top: 0 !important;
}

/* (Optional) hub nav — comment out if you need it */
div[data-automationid="HubNav"],
div[data-automation-id="HubNav"],
div[class*="hubNav"] {
  display: none !important;
}
`;
  }

  // ---------------- MutationObserver for sticky header ----------------

  private _startObserving(): void {
    if (this._observer) return;

    this._observer = new MutationObserver((mutations) => {
      // We don't apply inline styles anymore; CSS rules in the <style> handle all matches.
      // The observer simply ensures our CSS keeps applying on dynamically injected nodes.
      // (No body needed; keeping the observer is optional. You can remove it entirely if you prefer.)
    });

    this._observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
    });
  }

  private readonly _headerSelectors: string[] = [
    "div#spSiteHeader",
    '[data-sp-feature-tag="Site header host"]',
    '[aria-label="SharePoint Site Header"]',
    // Fallback inner surfaces
    'div[data-automationid="SiteHeader"]',
    'div[data-automationid="HorizontalNav"]',
    'div[data-automationid="MegaMenu"]',
    'div[data-automationid="StickyHeader"]',
    'div[data-automationid="StickyTopHeader"]',
    'div[data-automation-id="SiteHeader"]',
    'div[data-automation-id="HorizontalNav"]',
    'div[data-automation-id="MegaMenu"]',
    'div[data-automation-id="StickyHeader"]',
    'div[data-automation-id="StickyTopHeader"]',
  ];

  // Remove the inline styles we may have set in earlier builds
  private _unhideAll(): void {
    const props = [
      "display",
      "height",
      "min-height",
      "overflow",
      "padding",
      "margin",
      "border",
    ];
    for (const sel of this._headerSelectors) {
      document.querySelectorAll<HTMLElement>(sel).forEach((el) => {
        for (const p of props) el.style.removeProperty(p);
      });
    }
  }

  private _stopObserving(): void {
    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }
  }
}
