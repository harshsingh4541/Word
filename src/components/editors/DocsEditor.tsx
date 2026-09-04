"use client";

import { useEffect, useImperativeHandle, useRef, useState } from "react";
import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import { UniverDocsCorePreset } from "@univerjs/preset-docs-core";
import UniverPresetDocsCoreEnUS from "@univerjs/preset-docs-core/locales/en-US";
import { UniverDocsDrawingPreset } from "@univerjs/preset-docs-drawing";
import UniverPresetDocsDrawingEnUS from "@univerjs/preset-docs-drawing/locales/en-US";
import { UniverDocsHyperLinkPreset } from "@univerjs/preset-docs-hyper-link";
import UniverPresetDocsHyperLinkEnUS from "@univerjs/preset-docs-hyper-link/locales/en-US";
import { UniverDocsThreadCommentPreset } from "@univerjs/preset-docs-thread-comment";
import UniverPresetDocsThreadCommentEnUS from "@univerjs/preset-docs-thread-comment/locales/en-US";
import { UniverDocsFindReplacePlugin } from "@univerjs/docs-find-replace";
import { DocumentFlavor, ICommandService, IContextService, UniverInstanceType, validateDocumentStructure } from "@univerjs/core";
import type { DocumentDataModel, IDocumentData, Injector, Nullable } from "@univerjs/core";
import { IUniverInstanceService } from "@univerjs/core";
import { DocSelectionManagerService, DocSkeletonManagerService, SetTextSelectionsOperation } from "@univerjs/docs";
import { IRenderManagerService } from "@univerjs/engine-render";
import {
  ALL_TABLE_STYLE_COMMANDS,
  clearRememberedTableRange,
  resolveLiveTableRange,
} from "@/lib/univer/table-style-commands";
import { SetBorderPenCommand } from "@/lib/univer/border-pen";
import { loadSnapshot, saveSnapshot, clearSnapshot } from "@/lib/univer/persistence";
import {
  createWordCommands,
  SetIndentCommandId,
  SetPageMarginsCommandId,
  SetZoomCommandId,
} from "@/lib/univer/word-commands";
import WordRuler, { type RulerGeometry } from "./WordRuler";
import WordVerticalRuler from "./WordVerticalRuler";
import { BuiltInUIPart, IUIPartsService } from "@univerjs/ui";
import { installWordRibbon, RELOCATED_UNIVER_MENU_ITEMS, WORD_CURSOR_IN_TABLE_CTX, WORD_UI_LOCALE } from "@/lib/univer/word-ribbon";
import { createTableResizeInteraction } from "@/lib/univer/table-resize";
import { createTableMoveInteraction } from "@/lib/univer/table-move";
import { hidePageMarginMarks } from "@/lib/univer/page-chrome";
import { disableSlashMenu } from "@/lib/univer/slash-key";
import { restoreFocusAfterDialogs } from "@/lib/univer/editor-focus";
import { createWordFeatureCommands } from "@/lib/univer/word-features";
import { createSpellCheckCommand, createSpellChecker } from "@/lib/univer/spell-check";
import { createTrackChanges, createTrackChangesCommands } from "@/lib/univer/track-changes";
import { createWatermarkCommand } from "@/lib/univer/watermark";
import { buildWordLocale, WORD_THEME } from "@/lib/univer/word-theme";

const STORAGE_KEY = "docs-default";

// ─── Word paste utilities (module-level so PasteDialog can call them) ─────────

const WORD_HTML_RE = /urn:schemas-microsoft-com|mso-|class=?["']?Mso|ProgId=?["']?(?:Word|Excel)|Generator[^>]*(?:Microsoft Word|Microsoft Excel)|xmlns:[owv]=/i;

function pinWordComputedStyles(html: string, mode: "keep" | "clean"): string {
  const styleBlocks = (html.match(/<style[\s\S]*?<\/style>/gi) ?? []).join("\n");
  const body = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i)?.[1] ?? html;
  const holder = document.createElement("div");
  holder.style.cssText = "position:fixed;left:-99999px;top:0;visibility:hidden;width:660px;font-family:Arial,sans-serif;font-size:11pt;color:#000";
  holder.innerHTML = `${styleBlocks}${body}`;
  document.body.appendChild(holder);

  try {
    const blockTags = /^(P|DIV|LI|H[1-6]|BLOCKQUOTE|PRE)$/;
    const paragraphTags = /^(P|H[1-6]|LI)$/;
    holder.querySelectorAll<HTMLTableRowElement>("tr").forEach((row) => {
      if (/mso-height-rule\s*:\s*exactly/i.test(row.getAttribute("style") ?? "")) row.dataset.wordExactHeight = "true";
    });
    holder.querySelectorAll<HTMLElement>("*").forEach((el) => {
      if (/^(STYLE|META|LINK|TITLE|XML)$/i.test(el.tagName)) return;
      const computed = window.getComputedStyle(el);
      if (!el.style.fontWeight && computed.fontWeight !== "400" && computed.fontWeight !== "normal") el.style.fontWeight = computed.fontWeight;
      if (!el.style.fontStyle && computed.fontStyle !== "normal") el.style.fontStyle = computed.fontStyle;
      const decoration = computed.textDecorationLine || computed.textDecoration;
      if (!el.style.textDecoration && decoration && decoration !== "none") el.style.textDecoration = decoration;
      if (!el.style.color && computed.color) el.style.color = computed.color;
      if (!el.style.verticalAlign && (computed.verticalAlign === "super" || computed.verticalAlign === "sub")) el.style.verticalAlign = computed.verticalAlign;
      if (/^(TD|TH)$/i.test(el.tagName)) {
        const rowBackground = el.parentElement ? window.getComputedStyle(el.parentElement).backgroundColor : "rgba(0, 0, 0, 0)";
        const background = computed.backgroundColor !== "rgba(0, 0, 0, 0)" ? computed.backgroundColor : rowBackground;
        if (background !== "rgba(0, 0, 0, 0)") el.style.backgroundColor = background;
        (["Top", "Right", "Bottom", "Left"] as const).forEach((side) => {
          const style = computed[`border${side}Style`];
          const width = computed[`border${side}Width`];
          const color = computed[`border${side}Color`];
          el.dataset[`wordBorder${side}`] = style === "none" || width === "0px" ? "0px solid #000000" : `${width} ${style} ${color}`;
        });
        const cellAlign = el.getAttribute("align") || computed.textAlign;
        if (/^(left|center|right|justify|start|end)$/i.test(cellAlign)) {
          el.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6").forEach((paragraph) => {
            if (!paragraph.style.textAlign) paragraph.style.textAlign = cellAlign;
          });
        }
      }
      if (mode === "keep") {
        if (!el.style.fontFamily && computed.fontFamily) el.style.fontFamily = computed.fontFamily;
        if (!el.style.fontSize && computed.fontSize) el.style.fontSize = computed.fontSize;
      }
      if (!blockTags.test(el.tagName)) return;
      if (!el.style.lineHeight && computed.lineHeight !== "normal") el.style.lineHeight = computed.lineHeight;
      if (!el.style.marginTop && computed.marginTop) el.style.marginTop = computed.marginTop;
      if (!el.style.marginBottom && computed.marginBottom) el.style.marginBottom = computed.marginBottom;
      if (!el.style.textAlign) {
        const explicitAlign = el.getAttribute("align");
        const marginLeft = parseFloat(computed.marginLeft) || 0;
        const marginRight = parseFloat(computed.marginRight) || 0;
        if (explicitAlign) el.style.textAlign = explicitAlign;
        else if (paragraphTags.test(el.tagName) && marginLeft > 330 && marginRight < 40) el.style.textAlign = "right";
        else el.style.textAlign = computed.textAlign === "start" ? "left" : computed.textAlign;
      }
      if (!el.style.textIndent && computed.textIndent !== "0px") el.style.textIndent = computed.textIndent;
      if (paragraphTags.test(el.tagName) && !el.style.marginLeft && computed.marginLeft !== "0px") el.style.marginLeft = computed.marginLeft;
    });

    holder.querySelectorAll("style, meta, link, title, xml, o\\:p").forEach((el) => el.remove());
    holder.querySelectorAll<HTMLElement>("td, th").forEach((cell) => {
      const hasContent = /[^\s\u00A0\u200B]/.test(cell.textContent ?? "") || !!cell.querySelector("img, table");
      if (!hasContent) return;
      cell.querySelectorAll("p").forEach((p) => {
        if (p.querySelector("img, table") || /[^\s\u00A0\u200B]/.test(p.textContent ?? "") || p.innerHTML.includes("{{")) return;
        p.remove();
      });
    });
    return holder.innerHTML;
  } finally {
    holder.remove();
  }
}

function cleanWordHtml(html: string, mode: "keep" | "clean" = "keep"): string {
  const normalizedHtml = html
    .replace(/mso-margin-top-alt\s*:/gi, "margin-top:")
    .replace(/mso-margin-bottom-alt\s*:/gi, "margin-bottom:")
    .replace(/mso-line-height-alt\s*:/gi, "line-height:")
    .replace(/mso-para-margin-top\s*:/gi, "margin-top:")
    .replace(/mso-para-margin-bottom\s*:/gi, "margin-bottom:");

  // Phase 0: DOM conversions before mso-* stripping
  let working = pinWordComputedStyles(normalizedHtml, mode);
  try {
    const p0 = new DOMParser().parseFromString(working, "text/html");

    // ① MsoHeading → <h1>–<h5>
    p0.querySelectorAll("p").forEach((p) => {
      const m = (p as HTMLElement).className.match(/\bMsoHeading(\d)\b/i);
      if (!m) return;
      const level = Math.min(5, Number(m[1]));
      const h = p0.createElement(`h${level}`);
      [...(p as HTMLElement).attributes].forEach((a) => h.setAttribute(a.name, a.value));
      h.innerHTML = p.innerHTML;
      p.replaceWith(h);
    });

    // ② Word list paragraphs → <ul>/<ol>/<li> (supports multi-level via mso-list:lX levelN)
    (function convertWordLists(doc: Document) {
      const parseMsoList = (el: HTMLElement): { listId: string; level: number } | null => {
        const style = el.getAttribute("style") ?? "";
        const m = style.match(/mso-list:\s*l(\d+)\s+level(\d+)/i);
        if (m) return { listId: m[1], level: Number(m[2]) };
        if (/MsoListParagraph/i.test(el.className)) return { listId: "0", level: 1 };
        return null;
      };
      const isWordListPara = (el: Element): el is HTMLElement =>
        el.tagName === "P" && parseMsoList(el as HTMLElement) !== null;
      const containers = new Set<Element>();
      doc.querySelectorAll("p").forEach((p) => {
        if (isWordListPara(p)) containers.add(p.parentElement ?? doc.body);
      });
      for (const container of containers) {
        const kids = Array.from(container.children);
        let i = 0;
        while (i < kids.length) {
          if (!isWordListPara(kids[i])) { i++; continue; }
          const group: HTMLElement[] = [];
          while (i < kids.length && isWordListPara(kids[i])) { group.push(kids[i] as HTMLElement); i++; }

          // Determine top-level list type from first item's marker
          const firstIgnore = group[0].querySelector<HTMLElement>('[style*="mso-list:Ignore"]');
          const firstMarker = (firstIgnore?.textContent ?? "").replace(/\s/g, "");
          const firstOrdered = /^[0-9]+[.)]|^[a-zA-Z]{1,3}[.)]/.test(firstMarker);

          // Build nested structure: stack[level-1] = current list at that level
          const rootList = doc.createElement(firstOrdered ? "ol" : "ul");
          const listStack: HTMLUListElement[] = [rootList as unknown as HTMLUListElement];

          group.forEach((p) => {
            const info = parseMsoList(p) ?? { listId: "0", level: 1 };
            const level = Math.max(1, info.level);
            p.querySelectorAll('[style*="mso-list:Ignore"]').forEach((s) => s.remove());
            const li = doc.createElement("li");
            li.innerHTML = p.innerHTML;
            const cls = p.getAttribute("class");
            if (cls) li.setAttribute("class", cls);
            const rawStyle = (p.getAttribute("style") ?? "").split(";")
              .filter((s) => s.trim() && !/^mso-|text-indent/i.test(s.trim())).join("; ").trim();
            if (rawStyle) li.setAttribute("style", rawStyle);

            while (listStack.length < level) {
              // Need a deeper list: append to last li of current deepest list
              const parent = listStack[listStack.length - 1];
              const lastLi = parent.lastElementChild ?? parent.appendChild(doc.createElement("li"));
              const nested = doc.createElement("ul") as unknown as HTMLUListElement;
              lastLi.appendChild(nested);
              listStack.push(nested);
            }
            while (listStack.length > level) listStack.pop();
            listStack[listStack.length - 1].appendChild(li);
          });

          group[0].replaceWith(rootList);
          group.slice(1).forEach((el) => el.remove());
        }
      }
    })(p0);

    // ③ <br> inside paragraphs → paragraph splits
    // When a <br> is inside a <span>, siblings after it must be re-wrapped in
    // a clone of that span so formatting (font, bold, color) is preserved.
    p0.querySelectorAll("p").forEach((p) => {
      const brs = [...p.querySelectorAll("br")];
      if (brs.length === 0) return;
      brs.forEach((br) => {
        const newP = p0.createElement("p");
        const cls = p.getAttribute("class"); const sty = p.getAttribute("style");
        if (cls) newP.setAttribute("class", cls);
        if (sty) newP.setAttribute("style", sty);
        // If the <br> is inside a span, wrap trailing siblings in a clone of that span
        const parentSpan = br.parentElement !== p && br.parentElement?.tagName === "SPAN"
          ? br.parentElement : null;
        if (parentSpan) {
          // Move nodes after the br that are inside the span into a new span clone
          const spanClone = parentSpan.cloneNode(false) as HTMLElement;
          let next: ChildNode | null = br.nextSibling;
          while (next) { const tmp: ChildNode | null = next.nextSibling; spanClone.appendChild(next); next = tmp; }
          br.remove();
          parentSpan.insertAdjacentElement("afterend", spanClone);
          // Now move the span clone and everything after it into newP
          let after: ChildNode | null = spanClone;
          while (after) { const tmp: ChildNode | null = after.nextSibling; newP.appendChild(after); after = tmp; }
        } else {
          let next: ChildNode | null = br.nextSibling;
          while (next) { const tmp: ChildNode | null = next.nextSibling; newP.appendChild(next); next = tmp; }
          br.remove();
        }
        p.insertAdjacentElement("afterend", newP);
      });
    });

    // ④ CSS vertical-align super/sub → <sup>/<sub>
    p0.querySelectorAll<HTMLElement>('span[style*="vertical-align"]').forEach((span) => {
      const va = span.style.verticalAlign;
      const tag = va === "super" ? "sup" : va === "sub" ? "sub" : null;
      if (!tag) return;
      const el = p0.createElement(tag);
      [...span.attributes].forEach((a) => el.setAttribute(a.name, a.value));
      el.style.removeProperty("vertical-align");
      el.innerHTML = span.innerHTML;
      span.replaceWith(el);
    });

    // ⑤ Track changes: strip deleted text, unwrap inserted text
    p0.querySelectorAll("del").forEach((el) => el.remove());
    p0.querySelectorAll<HTMLElement>('[class*="MsoDelText"], [class*="msoDel"]').forEach((el) => el.remove());
    p0.querySelectorAll("ins").forEach((ins) => {
      const frag = p0.createDocumentFragment();
      while (ins.firstChild) frag.appendChild(ins.firstChild);
      ins.replaceWith(frag);
    });

    // ⑥ Footnotes/endnotes: extract footnote text and append as a section at bottom.
    // Word HTML places footnote markers as <a href="#_ftn1"> in body text and
    // the actual footnote content in a div[style*="mso-element:footnote-list"].
    const footnoteContainer = p0.querySelector<HTMLElement>('[style*="mso-element:footnote-list"], [style*="mso-element:endnote-list"]');
    if (footnoteContainer) {
      const entries: { num: string; text: string }[] = [];
      footnoteContainer.querySelectorAll<HTMLElement>('[style*="mso-element:footnote"], [style*="mso-element:endnote"]').forEach((fn) => {
        const numEl = fn.querySelector("a[href]") ?? fn.querySelector("sup");
        const num = (numEl?.textContent ?? "").trim() || String(entries.length + 1);
        fn.querySelectorAll("a").forEach((a) => a.remove());
        const text = fn.textContent?.trim() ?? "";
        if (text) entries.push({ num, text });
      });
      footnoteContainer.remove();
      if (entries.length > 0) {
        const hr = p0.createElement("hr");
        p0.body.appendChild(hr);
        const sect = p0.createElement("div");
        sect.style.cssText = "font-size:10pt; margin-top:8pt;";
        entries.forEach(({ num, text }) => {
          const p = p0.createElement("p");
          const escaped = text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          p.innerHTML = `<sup>${num}</sup>&nbsp;${escaped}`;
          sect.appendChild(p);
        });
        p0.body.appendChild(sect);
      }
    }

    // ⑦ Floating / absolutely-positioned images → inline
    // Word's clipboard HTML wraps floating images in <v:shape> with an
    // <!--[if !vml]--> fallback <img position:absolute>. We keep the img (via
    // the $1 substitution in Phase 2) but must strip the absolute positioning.
    p0.querySelectorAll<HTMLImageElement>("img").forEach((img) => {
      const pos = img.style.position;
      if (pos === "absolute" || pos === "fixed") {
        img.style.removeProperty("position");
        img.style.removeProperty("left");
        img.style.removeProperty("top");
        img.style.removeProperty("z-index");
        img.style.removeProperty("margin-left");
        img.style.removeProperty("margin-top");
      }
      const fl = img.style.float;
      if (fl === "left" || fl === "right") {
        img.style.removeProperty("float");
        img.style.display = "block";
        img.style.margin = "4pt 0";
      }
      if (!img.style.maxWidth) img.style.maxWidth = "100%";
    });
    // Also handle wrapping <span>/<p> that are position:absolute (image anchors)
    p0.querySelectorAll<HTMLElement>('span[style*="position:absolute"], p[style*="position:absolute"]').forEach((el) => {
      el.style.removeProperty("position");
      el.style.removeProperty("left");
      el.style.removeProperty("top");
    });

    // ⑧ Preserve inter-run spaces. Word's clipboard HTML frequently kerns
    // individual runs by wrapping a lone space (or, in wrapped source HTML,
    // a bare newline — both collapse to a single space under normal HTML
    // whitespace rules) in its own <span style="letter-spacing:...">, e.g.
    // `Folio<span> </span>No.:` or `amounts<span>\n</span>in`. Univer's
    // clipboard importer (normalizeTextNode + hasTextSibling in
    // html-to-udm/converter.ts) decides whether to keep a whitespace-only
    // text node by walking that node's *immediate DOM siblings* — but since
    // the whitespace is the sole child of its own dedicated span, it has no
    // siblings at all, so Univer drops it outright regardless of its
    // content (and even a bare "\n", unlike tab/&nbsp;, never counts as
    // "explicit spacing" for it either) — silently gluing words together
    // ("Folio No." -> "FolioNo.", "All amounts in INR" -> "AllamountsinINR").
    // Re-encoding it as &nbsp; alone doesn't help since the sibling check
    // still fails first. Instead, splice a normalized single space directly
    // into the neighboring word's own text node (in document order,
    // regardless of nesting) so it's never isolated, and drop the now-empty
    // wrapper span.
    (function preserveInterRunSpaces(doc: Document) {
      const BLOCK_SELECTOR = "p,li,h1,h2,h3,h4,h5,h6,td,th,div,body";
      const nearestBlock = (node: Node): Element | null =>
        (node.nodeType === Node.ELEMENT_NODE ? (node as Element) : node.parentElement)?.closest(BLOCK_SELECTOR) ?? null;

      const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);
      const textNodes: Text[] = [];
      let node: Node | null;
      while ((node = walker.nextNode())) textNodes.push(node as Text);

      textNodes.forEach((t, i) => {
        if (!/^\s+$/.test(t.nodeValue ?? "") || t.nodeValue === "") return;
        const prev = textNodes[i - 1];
        const next = textNodes[i + 1];
        if (!prev?.nodeValue?.trim() || !next?.nodeValue?.trim()) return;
        const block = nearestBlock(t);
        if (!block || nearestBlock(prev) !== block || nearestBlock(next) !== block) return;
        // Fold a single normalized space into the end of the previous word's
        // text node (matching HTML's own whitespace-collapsing rules) so it
        // always has real text siblings, then remove the now-empty wrapper.
        prev.nodeValue = (prev.nodeValue ?? "") + " ";
        const parent = t.parentNode;
        parent?.removeChild(t);
        if (parent && parent.nodeType === Node.ELEMENT_NODE && !(parent as Element).hasChildNodes()) {
          (parent as Element).remove();
        }
      });
    })(p0);

    working = p0.body.innerHTML;
  } catch { /* DOMParser unavailable */ }

  // Phase 2: Regex cleanup
  let clean = working
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<!--\[if\s+vml\b[\s\S]*?<!\[endif\]-->/gi, "")
    .replace(/<!--\[if\s*![^\]]*\]>([\s\S]*?)<!\[endif\]-->/gi, "$1")
    .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "")
    .replace(/<!--\[if[^\]]*\]-->/gi, "")
    .replace(/<!--\[endif\]-->/gi, "")
    .replace(/<o:p[^>]*>[\s\S]*?<\/o:p>/gi, "")
    .replace(/<o:p\s*\/>/gi, "")
    .replace(/<\/?w:[^>]*>/gi, "").replace(/<\/?v:[^>]*>/gi, "").replace(/<\/?m:[^>]*>/gi, "")
    .replace(/(style="[^"]*?)(?:\s*mso-[^:]+:[^;";]+;?)+/gi, "$1")
    .replace(/(style='[^']*?)(?:\s*mso-[^:]+:[^;';]+;?)+/gi, "$1")
    .replace(/\s+xmlns[^=]*="[^"]*"/gi, "")
    .replace(/\s+(?:v|o|w):\w+="[^"]*"/gi, "");

  // Phase 3: DOMParser normalisation
  try {
    const tmpDoc = new DOMParser().parseFromString(clean, "text/html");

    // Strip remaining mso-* from all inline styles
    tmpDoc.querySelectorAll("*").forEach((el) => {
      const s = (el as HTMLElement).style;
      if (!s?.cssText) return;
      s.cssText = s.cssText.replace(/\s*mso-[^:]+:[^;]+;?\s*/gi, "").trim();
    });

    // Drop letter-spacing entirely. Word bakes tiny per-run kerning values
    // (e.g. -0.1px to -0.3px) into justified/optically-aligned text for
    // pixel-perfect fidelity in its own renderer. Univer's beta canvas text
    // layout appears to mishandle these micro negative-tracking values on
    // short runs, visually collapsing adjacent characters/spaces to zero
    // width (e.g. "Folio No.:" rendering as "FolioNo.:") even though the
    // underlying text/space is present and correct in the document model.
    // The values are cosmetically inconsequential, so removing them avoids
    // the rendering bug entirely without any visible fidelity loss.
    tmpDoc.querySelectorAll<HTMLElement>("*").forEach((el) => {
      if (el.style?.letterSpacing) el.style.removeProperty("letter-spacing");
    });

    // line-height: convert to a unitless ratio so Univer parses it correctly.
    //
    // Univer's own paragraph-style reader (getParagraphStyle in
    // html-to-udm/utils.ts) does `Number.parseFloat(cssValue)` on
    // `line-height` and stores that raw number as `lineSpacing` — a
    // *multiplier* of the paragraph's normal line height. It only special-
    // cases the "%" suffix (dividing by 100); any other unit's numeric part
    // is used verbatim. Word commonly emits absolute line-heights like
    // `line-height:11.5pt` (its "single spacing" annotation for an 11pt
    // font, mso-line-height-rule:exactly) — a *previous* version of this
    // step converted that to `15px`, which Univer then reads as
    // `lineSpacing: 15`, i.e. 1500% line height, ballooning that paragraph's
    // (and its table row's) height by ~15x. The fix is to never hand Univer
    // an absolute unit at all: resolve the paragraph's effective font size
    // and express line-height as the ratio Word actually intended.
    const getEffectiveFontSizePt = (el: HTMLElement): number => {
      const parsePt = (v: string | null | undefined): number | null => {
        if (!v) return null;
        const m = v.match(/^([\d.]+)(pt|px)?$/i);
        if (!m) return null;
        const num = parseFloat(m[1]);
        return (m[2] ?? "pt").toLowerCase() === "px" ? num / 1.3333 : num;
      };
      let found = parsePt(el.style?.fontSize);
      if (found) return found;
      const withSize = el.querySelector<HTMLElement>('[style*="font-size"]');
      found = withSize ? parsePt(withSize.style.fontSize) : null;
      if (found) return found;
      let ancestor = el.parentElement;
      while (ancestor) {
        found = parsePt(ancestor.style?.fontSize);
        if (found) return found;
        ancestor = ancestor.parentElement;
      }
      return 11; // Word's common "Normal" style default
    };
    tmpDoc.querySelectorAll<HTMLElement>("*").forEach((el) => {
      const lh = el.style?.lineHeight;
      if (!lh) return;
      if (lh.endsWith("%")) {
        el.style.lineHeight = (parseFloat(lh) / 100).toFixed(2);
        return;
      }
      const m = lh.match(/^([\d.]+)(pt|px|cm|in)$/i);
      if (!m) return; // already unitless (or unrecognized) — leave as-is
      const num = parseFloat(m[1]);
      const unit = m[2].toLowerCase();
      const lhPt = unit === "pt" ? num : unit === "px" ? num / 1.3333 : unit === "cm" ? num / 0.03528 : num * 72;
      const fontPt = getEffectiveFontSizePt(el);
      const ratio = fontPt > 0 ? lhPt / fontPt : 1.15;
      el.style.lineHeight = ratio.toFixed(2);
    });

    // paragraph-level spacing: margin-top/bottom pt → px
    tmpDoc.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, li").forEach((el) => {
      ["marginTop", "marginBottom"].forEach((prop) => {
        const val: string = (el.style as unknown as Record<string, string>)[prop] ?? "";
        if (val.endsWith("pt")) {
          (el.style as unknown as Record<string, string>)[prop] =
            `${Math.round(parseFloat(val) * 1.3333)}px`;
        }
      });
    });

    tmpDoc.querySelectorAll<HTMLTableCellElement>("td, th").forEach((cell) => {
      const parents = new Set<Element>();
      cell.querySelectorAll("p, h1, h2, h3, h4, h5, h6, li").forEach((block) => {
        if (block.closest("td, th") === cell && block.parentElement) parents.add(block.parentElement);
      });
      const spacingToPx = (value: string): number => {
        const match = value.trim().match(/^(-?[\d.]+)\s*(px|pt|cm|mm|in)?$/i);
        if (!match) return 0;
        const number = parseFloat(match[1]);
        switch ((match[2] ?? "px").toLowerCase()) {
          case "pt": return number * 96 / 72;
          case "cm": return number * 96 / 2.54;
          case "mm": return number * 96 / 25.4;
          case "in": return number * 96;
          default: return number;
        }
      };
      parents.forEach((parent) => {
        const blocks = [...parent.children].filter((child): child is HTMLElement => /^(P|H[1-6]|LI)$/i.test(child.tagName));
        for (let index = 1; index < blocks.length; index++) {
          const previous = blocks[index - 1];
          const current = blocks[index];
          const collapsedGap = Math.max(spacingToPx(previous.style.marginBottom), spacingToPx(current.style.marginTop));
          previous.style.marginBottom = `${Math.round(collapsedGap * 100) / 100}px`;
          current.style.marginTop = "0px";
        }
      });
    });

    // text-transform:uppercase → bake uppercase text before stripping styles
    tmpDoc.querySelectorAll<HTMLElement>("*").forEach((el) => {
      if (el.style?.textTransform === "uppercase") {
        el.childNodes.forEach((node) => {
          if (node.nodeType === Node.TEXT_NODE && node.textContent) {
            node.textContent = node.textContent.toUpperCase();
          }
        });
        el.style.removeProperty("text-transform");
      }
    });

    // Table width normalisation: 100% table, proportional px cells
    //
    // Word's raw clipboard HTML mixes units freely between the table and its
    // cells: the <table> often only has a unitless `width=700` HTML
    // attribute (px-equivalent), while each <td> carries a CSS
    // `style="width:330.35pt"` (points) alongside its own `width=440` HTML
    // attribute (px-equivalent, matching the table's unit). A naive
    // `parseFloat(cell.style.width)` strips the "pt" suffix and treats 330.35
    // as if it were in the same unit as the table's 700, badly under-scaling
    // the cell (e.g. 311px instead of the correct ~414px) — which starves
    // its text of width and forces excess line-wrapping, visually blowing up
    // that row's height. Always resolve both table and cell widths to px
    // first, preferring the (unit-consistent) HTML attribute over a
    // mismatched-unit CSS style when the two disagree in kind.
    const widthToPx = (value: string | null | undefined): number => {
      const match = value?.trim().match(/^(-?[\d.]+)\s*(px|pt|cm|mm|in|pc)?$/i);
      if (!match) return 0;
      const number = parseFloat(match[1]);
      switch ((match[2] ?? "px").toLowerCase()) {
        case "pt": return number * 96 / 72;
        case "cm": return number * 96 / 2.54;
        case "mm": return number * 96 / 25.4;
        case "in": return number * 96;
        case "pc": return number * 16;
        default: return number;
      }
    };
    const elementWidthToPx = (el: HTMLElement, relativeTo = 0): number => {
      const value = el.getAttribute("width") || el.style.width;
      const percent = value?.trim().match(/^([\d.]+)%$/);
      return percent && relativeTo > 0 ? relativeTo * parseFloat(percent[1]) / 100 : widthToPx(value);
    };

    type TablePlacement = { cell: HTMLTableCellElement; start: number; span: number };
    const buildTableGrid = (table: HTMLTableElement) => {
      const rows = [...table.rows].filter((row) => row.closest("table") === table);
      const slots: Array<Array<HTMLTableCellElement | undefined>> = rows.map(() => []);
      const placements = new Map<HTMLTableCellElement, TablePlacement>();
      rows.forEach((row, rowIndex) => {
        let column = 0;
        [...row.cells].forEach((cell) => {
          while (slots[rowIndex][column]) column++;
          const span = Math.max(1, cell.colSpan || 1);
          const rowSpan = Math.max(1, cell.rowSpan || 1);
          placements.set(cell, { cell, start: column, span });
          for (let rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            if (!slots[rowIndex + rowOffset]) slots[rowIndex + rowOffset] = [];
            for (let columnOffset = 0; columnOffset < span; columnOffset++) slots[rowIndex + rowOffset][column + columnOffset] = cell;
          }
          column += span;
        });
      });
      return { rows, slots, placements, columnCount: Math.max(0, ...slots.map((row) => row.length)) };
    };
    const isEmptyTableCell = (cell: HTMLTableCellElement) =>
      !cell.querySelector("img, table, hr") && !cell.innerHTML.includes("{{") && !/[^\s\u00A0\u200B]/.test(cell.textContent ?? "");
    const trimTrailingEmptyColumns = (table: HTMLTableElement): boolean => {
      let changed = false;
      for (let pass = 0; pass < 40; pass++) {
        const grid = buildTableGrid(table);
        if (grid.columnCount <= 1) return changed;
        const lastColumn = grid.columnCount - 1;
        const owners = new Set(grid.slots.map((row) => row[lastColumn]).filter((cell): cell is HTMLTableCellElement => !!cell));
        let foundDedicatedEmpty = false;
        let removable = owners.size > 0;
        owners.forEach((cell) => {
          const placement = grid.placements.get(cell);
          if (!placement) return;
          if (placement.start < lastColumn && placement.start + placement.span - 1 >= lastColumn) return;
          if (placement.span === 1 && isEmptyTableCell(cell)) foundDedicatedEmpty = true;
          else removable = false;
        });
        if (!removable || !foundDedicatedEmpty) return changed;
        owners.forEach((cell) => {
          const placement = grid.placements.get(cell);
          if (!placement) return;
          if (placement.start < lastColumn && placement.span > 1) cell.colSpan = placement.span - 1;
          else if (placement.start === lastColumn && placement.span === 1 && isEmptyTableCell(cell)) cell.remove();
        });
        changed = true;
      }
      return changed;
    };

    [...tmpDoc.querySelectorAll<HTMLTableElement>("table")]
      .filter((table) => !!table.parentElement?.closest("td, th"))
      .reverse()
      .forEach((table) => {
        const replacement = tmpDoc.createElement("div");
        [...table.rows].filter((row) => row.closest("table") === table).forEach((row) => {
          const paragraph = tmpDoc.createElement("p");
          paragraph.className = "UniverNormal";
          [...row.cells].forEach((cell, index) => {
            if (index) paragraph.append(" | ");
            [...cell.childNodes].forEach((node) => {
              if (node.nodeType === Node.ELEMENT_NODE && /^(P|DIV)$/i.test((node as Element).tagName)) {
                [...node.childNodes].forEach((child) => paragraph.appendChild(child.cloneNode(true)));
              } else {
                paragraph.appendChild(node.cloneNode(true));
              }
            });
          });
          replacement.appendChild(paragraph);
        });
        table.replaceWith(replacement);
      });

    tmpDoc.querySelectorAll<HTMLTableElement>("table").forEach((table) => {
      [...table.querySelectorAll<HTMLTableCellElement>("td, th")]
        .filter((cell) => cell.closest("table") === table && cell.style.display === "none")
        .forEach((cell) => cell.remove());
      const trimmedTrailingColumns = trimTrailingEmptyColumns(table);

      const explicitTableWidth = elementWidthToPx(table);
      const parentCell = table.parentElement?.closest<HTMLTableCellElement>("td, th");
      const availableWidth = parentCell ? Math.max(80, elementWidthToPx(parentCell) - 20) : 660;
      const grid = buildTableGrid(table);
      const ownCols = [...table.querySelectorAll<HTMLTableColElement>("colgroup > col")]
        .filter((col) => col.closest("table") === table);
      const columnWidths = Array<number>(grid.columnCount).fill(0);
      let columnIndex = 0;
      ownCols.forEach((col) => {
        const width = elementWidthToPx(col, explicitTableWidth || availableWidth);
        const span = Math.max(1, Number(col.span) || 1);
        for (let offset = 0; offset < span && columnIndex + offset < columnWidths.length; offset++) columnWidths[columnIndex + offset] = width;
        columnIndex += span;
      });
      grid.placements.forEach(({ cell, start, span }) => {
        if (span !== 1 || columnWidths[start] > 0) return;
        columnWidths[start] = elementWidthToPx(cell, explicitTableWidth || availableWidth);
      });
      grid.placements.forEach(({ cell, start, span }) => {
        if (span <= 1) return;
        const cellWidth = elementWidthToPx(cell, explicitTableWidth || availableWidth);
        if (!cellWidth) return;
        const indexes = Array.from({ length: span }, (_, offset) => start + offset).filter((index) => index < columnWidths.length);
        const missing = indexes.filter((index) => !columnWidths[index]);
        if (!missing.length) return;
        const knownWidth = indexes.reduce((sum, index) => sum + columnWidths[index], 0);
        const share = Math.max(1, (cellWidth - knownWidth) / missing.length);
        missing.forEach((index) => { columnWidths[index] = share; });
      });
      const knownWidth = columnWidths.reduce((sum, width) => sum + width, 0);
      const missingColumns = columnWidths.filter((width) => !width).length;
      if (missingColumns) {
        const fallbackTotal = explicitTableWidth || knownWidth || availableWidth;
        const share = Math.max(8, (fallbackTotal - knownWidth) / missingColumns || fallbackTotal / grid.columnCount);
        columnWidths.forEach((width, index) => { if (!width) columnWidths[index] = share; });
      }
      const resolvedColumnWidth = columnWidths.reduce((sum, width) => sum + width, 0);
      const sourceTotal = (trimmedTrailingColumns ? resolvedColumnWidth : explicitTableWidth) || resolvedColumnWidth || availableWidth;
      const targetWidth = Math.min(sourceTotal, availableWidth);
      const scale = sourceTotal > 0 ? targetWidth / sourceTotal : 1;
      const normalizedColumns = columnWidths.map((width) => Math.max(8, Math.round(width * scale)));
      if (normalizedColumns.length) normalizedColumns[normalizedColumns.length - 1] += Math.round(targetWidth) - normalizedColumns.reduce((sum, width) => sum + width, 0);

      [...table.querySelectorAll("colgroup")]
        .filter((group) => group.closest("table") === table)
        .forEach((group) => group.remove());
      if (normalizedColumns.length) {
        const group = tmpDoc.createElement("colgroup");
        normalizedColumns.forEach((width) => {
          const col = tmpDoc.createElement("col");
          col.style.width = `${width}px`;
          col.setAttribute("width", String(width));
          group.appendChild(col);
        });
        table.insertBefore(group, table.firstChild);
      }

      table.removeAttribute("width");
      table.removeAttribute("height");
      table.removeAttribute("cellpadding");
      table.removeAttribute("contenteditable");
      table.removeAttribute("unselectable");
      table.style.removeProperty("height");
      table.style.removeProperty("min-height");
      table.style.removeProperty("pointer-events");
      table.style.width = `${Math.max(80, normalizedColumns.reduce((sum, width) => sum + width, 0) || targetWidth)}px`;
      table.style.borderCollapse = "collapse";
      const isBorderless = table.getAttribute("border") === "0";

      grid.rows.forEach((row) => {
        const rowHeight = widthToPx(row.getAttribute("height") || row.style.height);
        const hasExactHeight = row.dataset.wordExactHeight === "true";
        delete row.dataset.wordExactHeight;
        row.style.removeProperty("min-height");
        if (hasExactHeight && rowHeight >= 16 && rowHeight <= 1000) {
          const normalizedHeight = Math.round(rowHeight * 100) / 100;
          row.setAttribute("height", String(normalizedHeight));
          row.style.height = `${normalizedHeight}px`;
        } else {
          row.removeAttribute("height");
          row.style.removeProperty("height");
        }
        [...row.cells].forEach((cell) => {
          const placement = grid.placements.get(cell);
          const normalizedWidth = placement
            ? normalizedColumns.slice(placement.start, placement.start + placement.span).reduce((sum, width) => sum + width, 0)
            : 0;
          cell.removeAttribute("width");
          cell.removeAttribute("height");
          cell.removeAttribute("nowrap");
          cell.removeAttribute("contenteditable");
          cell.removeAttribute("unselectable");
          cell.style.removeProperty("height");
          cell.style.removeProperty("min-height");
          cell.style.removeProperty("pointer-events");
          cell.style.removeProperty("user-select");
          if (normalizedWidth > 0) {
            cell.style.width = `${normalizedWidth}px`;
            cell.setAttribute("width", String(normalizedWidth));
          }
          cell.querySelectorAll<HTMLElement>('[contenteditable="false"], [unselectable="on"], [style*="pointer-events"], [style*="user-select"]').forEach((el) => {
            el.removeAttribute("contenteditable");
            el.removeAttribute("unselectable");
            el.style.removeProperty("pointer-events");
            el.style.removeProperty("user-select");
          });
          const borders = (["Top", "Right", "Bottom", "Left"] as const).map((side) => ({
            side,
            value: cell.dataset[`wordBorder${side}`] || (isBorderless ? "0px solid #000000" : ""),
          }));
          ["border", "border-width", "border-style", "border-color", "border-top", "border-right", "border-bottom", "border-left"].forEach((property) => cell.style.removeProperty(property));
          const borderCss = borders.filter(({ value }) => value).map(({ side, value }) => `border-${side.toLowerCase()}:${value}`).join(";");
          (["Top", "Right", "Bottom", "Left"] as const).forEach((side) => { delete cell.dataset[`wordBorder${side}`]; });
          if (borderCss) cell.setAttribute("style", `${cell.getAttribute("style") ?? ""};${borderCss}`);
          const hasBlock = cell.querySelector("p, div, h1, h2, h3, h4, h5, h6, ul, ol, pre, table");
          if (!hasBlock) {
            const paragraph = tmpDoc.createElement("p");
            paragraph.className = "UniverNormal";
            paragraph.innerHTML = cell.innerHTML.trim() || "&nbsp;";
            cell.replaceChildren(paragraph);
          }
        });
      });
    });

    // Headings: preserve heading semantics via data-heading attribute so
    // Univer's getHeadingNamedStyleType fires, while also adding UniverNormal
    // for paragraph style resolution (text-align, line-height, spacing).
    tmpDoc.querySelectorAll("h1, h2, h3, h4, h5").forEach((h) => {
      const level = h.tagName.toLowerCase();
      (h as HTMLElement).setAttribute("data-heading", level);
      (h as HTMLElement).className = ((h as HTMLElement).className + " UniverNormal").trim();
    });

    // UniverNormal triggers getParagraphStyle() for text-align, line-height, margins
    tmpDoc.querySelectorAll("p").forEach((p) => {
      (p as HTMLElement).className = ((p as HTMLElement).className + " UniverNormal").trim();
    });

    // 'clean' mode: strip font-family/size/color so document styles apply
    if (mode === "clean") {
      tmpDoc.querySelectorAll("*").forEach((el) => {
        const s = (el as HTMLElement).style;
        if (!s?.cssText) return;
        ["font-family", "font-size", "color", "background-color", "background"].forEach((p) => s.removeProperty(p));
      });
    }

    clean = tmpDoc.body.innerHTML;
  } catch { /* DOMParser unavailable */ }
  return clean;
}

// ─── Paste-from-Word dialog ────────────────────────────────────────────────────

function PasteDialog({
  rawHtml,
  plainText,
  editorEl,
  pendingHtmlRef,
  pendingPlainRef,
  onClose,
}: {
  rawHtml: string;
  plainText: string;
  editorEl: Element | null;
  pendingHtmlRef: React.RefObject<string | null>;
  pendingPlainRef: React.RefObject<string | null>;
  onClose: () => void;
}) {
  const insert = (mode: "keep" | "clean" | "text") => {
    onClose();
    let html: string;
    if (mode === "text") {
      html = "<p class=\"UniverNormal\">" +
        plainText
          .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
          .split("\n").filter(Boolean).join("</p><p class=\"UniverNormal\">") +
        "</p>";
    } else {
      html = cleanWordHtml(rawHtml, mode);
    }
    pendingHtmlRef.current = html;
    pendingPlainRef.current = plainText;
    const target = editorEl as HTMLElement | null;
    try { target?.focus(); } catch { /* ignore */ }
    const dt = new DataTransfer();
    dt.setData("text/html", "<p>x</p>");
    dt.setData("text/plain", plainText);
    (target ?? document.body).dispatchEvent(
      new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true })
    );
  };

  const btnBase: React.CSSProperties = {
    display: "block", width: "100%", textAlign: "left",
    padding: "11px 16px", marginBottom: 8, borderRadius: 7,
    cursor: "pointer", border: "1.5px solid #e2e8f0", background: "#fff",
  };

  return (
    <div
      style={{
        position: "fixed", inset: 0, zIndex: 9999,
        background: "rgba(0,0,0,0.45)",
        display: "flex", alignItems: "center", justifyContent: "center",
      }}
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div style={{
        background: "#fff", borderRadius: 10, padding: "28px 32px",
        maxWidth: 400, width: "90%", boxShadow: "0 8px 40px rgba(0,0,0,0.18)",
        fontFamily: "inherit",
      }}>
        <div style={{ fontSize: 17, fontWeight: 600, marginBottom: 6, color: "#0f172a" }}>
          Paste from Microsoft Office
        </div>
        <div style={{ fontSize: 13, color: "#64748b", marginBottom: 20 }}>
          How would you like to paste this content?
        </div>
        <button
          onClick={() => insert("keep")}
          style={{ ...btnBase, borderColor: "#2563eb", background: "#eff6ff" }}
        >
          <div style={{ fontWeight: 600, fontSize: 14, color: "#1d4ed8" }}>Keep Formatting</div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Preserve fonts, colors, and table layout</div>
        </button>
        <button
          onClick={() => insert("clean")}
          style={btnBase}
        >
          <div style={{ fontWeight: 600, fontSize: 14, color: "#1e293b" }}>Match Document Style</div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Keep structure, use document fonts</div>
        </button>
        <button
          onClick={() => insert("text")}
          style={btnBase}
        >
          <div style={{ fontWeight: 600, fontSize: 14, color: "#1e293b" }}>Text Only</div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Remove all formatting</div>
        </button>
        <button
          onClick={onClose}
          style={{
            display: "block", width: "100%", textAlign: "center",
            padding: "9px", marginTop: 4, border: "none",
            borderRadius: 7, cursor: "pointer", background: "transparent",
            color: "#64748b", fontSize: 13,
          }}
        >
          Cancel
        </button>
      </div>
    </div>
  );
}
const AUTOSAVE_DELAY_MS = 600;
const DEFAULT_DOCUMENT_NAME = "Untitled document";
const STATUS_REFRESH_DELAY_MS = 400;
// A4 at 96 DPI. Traditional flavor is what unlocks Word-compatible real
// pagination (page breaks, ruler-visible page bounds) and header/footer
// editing — both crash on creation-time documentStyle in Univer 0.25.x but
// work cleanly as of 1.0.0-beta.2.
const DEFAULT_DOCUMENT_STYLE = {
  pageSize: { width: 794, height: 1123 },
  documentFlavor: DocumentFlavor.TRADITIONAL,
};

import "@univerjs/preset-docs-core/lib/index.css";
import "@univerjs/preset-docs-drawing/lib/index.css";
import "@univerjs/preset-docs-hyper-link/lib/index.css";
import "@univerjs/preset-docs-thread-comment/lib/index.css";

/** What the Word-style title bar and status bar display. */
export type WordDocumentStatus = {
  name: string;
  wordCount: number;
  pageCount: number;
  currentPage: number;
  zoom: number;
};

/**
 * What the surrounding Word chrome can do to the document. Everything else
 * — formatting, layout, export — is a ribbon command inside Univer.
 */
export type DocsEditorHandle = {
  setName: (name: string) => void;
  setZoom: (zoom: number) => void;
  /** Live page geometry for the ruler, or null before the doc renders. */
  getRulerGeometry: () => RulerGeometry | null;
  setIndents: (indents: { indentStart?: number; indentEnd?: number; indentFirstLine?: number }) => void;
  setMargins: (margins: { marginLeft?: number; marginRight?: number }) => void;
};

export default function DocsEditor({
  apiRef,
  onStatusChange,
}: {
  apiRef?: React.RefObject<DocsEditorHandle | null>;
  onStatusChange?: (status: WordDocumentStatus) => void;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const disposedRef = useRef(false);
  const commandServiceRef = useRef<ICommandService | null>(null);
  const rulerGeometryRef = useRef<() => RulerGeometry | null>(() => null);
  const documentNameRef = useRef<(name: string) => void>(() => {});
  const statusListenerRef = useRef(onStatusChange);
  const [ready, setReady] = useState(false);
  const pendingHtmlRef = useRef<string | null>(null);
  const pendingPlainRef = useRef<string | null>(null);
  const [pasteDialog, setPasteDialog] = useState<{
    rawHtml: string;
    plainText: string;
    editorEl: Element | null;
  } | null>(null);

  // The editor is created once; the callback identity may change on every
  // parent render, so it is read through a ref rather than re-running setup.
  useEffect(() => {
    statusListenerRef.current = onStatusChange;
  }, [onStatusChange]);

  useEffect(() => {
    if (!containerRef.current || disposedRef.current) return;
    disposedRef.current = true;

    // Word types "/" as a character; Univer's block menu steals the key.
    disableSlashMenu();

    const { univer, univerAPI } = createUniver({
      theme: WORD_THEME,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: buildWordLocale(
          mergeLocales(
            UniverPresetDocsCoreEnUS,
            UniverPresetDocsDrawingEnUS,
            UniverPresetDocsHyperLinkEnUS,
            UniverPresetDocsThreadCommentEnUS,
          ),
          WORD_UI_LOCALE,
        ),
      },
      presets: [
        UniverDocsCorePreset({
          container: containerRef.current,
          // Word's ribbon: a tab strip over grouped, two-row controls.
          ribbonType: "grid",
          // Univer's own footer is replaced by a Word status bar that also
          // reports the page count.
          footer: false,
          menu: RELOCATED_UNIVER_MENU_ITEMS,
        }),
        UniverDocsDrawingPreset(),
        UniverDocsHyperLinkPreset(),
        UniverDocsThreadCommentPreset(),
      ],
      plugins: [UniverDocsFindReplacePlugin],
    });

    // Docs saved before the 1.0.0-beta.2 upgrade won't have a documentStyle
    // (it used to crash at creation time in 0.25.x — see git history), so
    // they'd silently lose pagination/header-footer on load. Backfill it
    // for any saved doc that predates this, without touching its content.
    let saved = loadSnapshot<Partial<IDocumentData>>(STORAGE_KEY);

    // 1.0.0-beta.2 added a strict structural-integrity check that now runs
    // on every edit (table start/end tokens, section IDs, etc.) and throws
    // if violated — Univer 0.25.x never validated this, so a doc edited
    // under the old version (in particular through our own dataStream-
    // editing MergeTableCellsCommand) can carry corruption that only
    // surfaces now, crashing on the very first edit after load. Check
    // before handing anything to createDocument(): a corrupt snapshot is
    // backed up under its own key (nothing is silently destroyed) and the
    // editor falls back to a fresh document instead of hard-crashing.
    if (saved?.body) {
      const issues = validateDocumentStructure(saved as Pick<IDocumentData, "body" | "headers" | "footers">);
      if (issues.length > 0) {
        console.warn("[DocKaro] Saved document failed structure validation, starting fresh:", issues);
        saveSnapshot(`${STORAGE_KEY}.corrupted.${Date.now()}`, saved);
        clearSnapshot(STORAGE_KEY);
        saved = null;
      }
    }

    const initialData: Partial<IDocumentData> = saved
      ? { ...saved, documentStyle: { ...DEFAULT_DOCUMENT_STYLE, ...saved.documentStyle } }
      : { documentStyle: DEFAULT_DOCUMENT_STYLE };
    // Word names a new document rather than leaving it blank, and this name
    // is what the title bar shows and what the export is filed under.
    if (!initialData.title) initialData.title = DEFAULT_DOCUMENT_NAME;
    const fDoc = univerAPI.createDocument(initialData);

    const injector = univer.__getInjector() as Injector;
    const commandService = injector.get(ICommandService);
    const spellChecker = createSpellChecker(injector, fDoc, () => containerRef.current);
    const trackChanges = createTrackChanges(injector, fDoc);
    const registrations = [
      SetBorderPenCommand,
      ...ALL_TABLE_STYLE_COMMANDS,
      ...createWordCommands({ doc: fDoc, getContainer: () => containerRef.current }),
      ...createWordFeatureCommands(fDoc),
      createSpellCheckCommand(spellChecker),
      ...createTrackChangesCommands(trackChanges),
      createWatermarkCommand(fDoc),
    ].map((command) => commandService.registerCommand(command));
    commandServiceRef.current = commandService;
    documentNameRef.current = (name: string) => {
      fDoc.setName(name);
      saveSnapshot(STORAGE_KEY, fDoc.save());
      void refreshStatus();
    };

    const wordRibbon = installWordRibbon(injector);
    const contextService = injector.get(IContextService);

    // Word puts its ruler between the ribbon and the page. Univer renders a
    // header slot in exactly that spot, so the ruler goes in as a UI part
    // rather than a sibling element that would sit above the ribbon.
    function DocumentRuler() {
      return (
        <WordRuler
          getGeometry={() => rulerGeometryRef.current()}
          handlers={{
            onIndentChange: (indents) => void commandService.executeCommand(SetIndentCommandId, indents),
            onMarginChange: (margins) => void commandService.executeCommand(SetPageMarginsCommandId, margins),
          }}
        />
      );
    }
    const rulerPart = injector.get(IUIPartsService).registerComponent(BuiltInUIPart.HEADER, () => DocumentRuler);

    // The ruler needs the page's on-screen position, which is the document
    // component's own offset inside the scene, shifted by the horizontal
    // scroll and multiplied by the zoom.
    rulerGeometryRef.current = () => {
      const container = containerRef.current;
      const renderUnit = renderManagerService.getRenderUnitById(fDoc.getId());
      const canvas = container?.querySelector("canvas");
      if (!container || !renderUnit || !canvas) return null;

      const documents = renderUnit.mainComponent as unknown as { left: number; top: number } | undefined;
      const scene = renderUnit.scene;
      const scale = scene.getAncestorScale().scaleX || 1;
      const scrollX = scene.getViewport("viewMain")?.viewportScrollX ?? 0;
      const canvasOffset = canvas.getBoundingClientRect().left - container.getBoundingClientRect().left;

      const docModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
      const style = docModel?.getDocumentStyle();
      if (!documents || !style?.pageSize?.width) return null;

      const paragraphStyle = currentParagraphStyle(docModel);
      const canvasRect = canvas.getBoundingClientRect();
      const containerRect = container.getBoundingClientRect();
      const scrollY = scene.getViewport("viewMain")?.viewportScrollY ?? 0;
      return {
        pageLeft: canvasOffset + (documents.left - scrollX) * scale,
        pageTop: canvasRect.top - containerRect.top + (documents.top - scrollY) * scale,
        pageWidth: style.pageSize.width * scale,
        pageHeight: (style.pageSize.height ?? 1123) * scale,
        marginLeft: style.marginLeft ?? 72,
        marginRight: style.marginRight ?? 72,
        marginTop: style.marginTop ?? 72,
        marginBottom: style.marginBottom ?? 72,
        indentStart: paragraphStyle?.indentStart?.v ?? 0,
        indentEnd: paragraphStyle?.indentEnd?.v ?? 0,
        indentFirstLine: paragraphStyle?.indentFirstLine?.v ?? 0,
        scale,
      };
    };
    // Word's table borders are draggable; Univer's have no such interaction.
    const tableResize = createTableResizeInteraction(injector, fDoc.getId(), () => containerRef.current);
    // Word shows a move handle at the top-left corner of a hovered table.
    const tableMove = createTableMoveInteraction(injector, fDoc.getId(), () => containerRef.current);

    // Word paste interception:
    // 1. Capture-phase listener shows the paste dialog for Word HTML.
    // 2. DataTransfer patch returns pendingHtmlRef content for synthetic paste
    //    dispatched by the dialog, and falls back to silent clean otherwise.
    const originalGetData = DataTransfer.prototype.getData;
    DataTransfer.prototype.getData = function (type: string): string {
      if (type === "text/html" && pendingHtmlRef.current !== null) {
        const h = pendingHtmlRef.current; pendingHtmlRef.current = null;
        return h;
      }
      if (type === "text/plain" && pendingPlainRef.current !== null) {
        const t = pendingPlainRef.current; pendingPlainRef.current = null; return t;
      }
      const data = originalGetData.call(this, type) as string;
      // Fallback: clean silently if Word HTML bypasses the capture listener
      if (type === "text/html" && WORD_HTML_RE.test(data)) return cleanWordHtml(data);
      return data;
    };

    const handleWordPasteCapture = (e: ClipboardEvent) => {
      const html = originalGetData.call(e.clipboardData, "text/html") as string;
      if (!html || !WORD_HTML_RE.test(html)) return;
      e.preventDefault();
      e.stopPropagation();
      const plain = originalGetData.call(e.clipboardData, "text/plain") as string;
      setPasteDialog({ rawHtml: html, plainText: plain, editorEl: document.activeElement });
    };
    document.addEventListener("paste", handleWordPasteCapture, true);

    // Secondary interception: programmatic clipboard reads
    const originalClipboardRead = navigator.clipboard.read.bind(navigator.clipboard);
    navigator.clipboard.read = async (...args) => {
      const items = await originalClipboardRead(...args);
      const cleaned: ClipboardItem[] = [];
      for (const item of items) {
        if (item.types.includes("text/html")) {
          const blob = await item.getType("text/html");
          const html = await blob.text();
          if (WORD_HTML_RE.test(html)) {
            const parts: Record<string, Blob | Promise<Blob>> = {
              "text/html": new Blob([cleanWordHtml(html)], { type: "text/html" }),
            };
            if (item.types.includes("text/plain")) parts["text/plain"] = item.getType("text/plain");
            cleaned.push(new ClipboardItem(parts));
            continue;
          }
        }
        cleaned.push(item);
      }
      return cleaned;
    };
    const pageChrome = hidePageMarginMarks(injector, fDoc.getId());
    const dialogFocus = restoreFocusAfterDialogs(injector, fDoc.getId());

    const renderManagerService = injector.get(IRenderManagerService);
    const docSelectionManagerService = injector.get(DocSelectionManagerService);
    const univerInstanceService = injector.get(IUniverInstanceService);

    /** The paragraph the cursor is in, whose indents the ruler shows. */
    const currentParagraphStyle = (docModel: Nullable<DocumentDataModel>) => {
      const offset = docSelectionManagerService.getActiveTextRange()?.startOffset;
      if (offset == null) return undefined;
      const paragraphs = docModel?.getBody()?.paragraphs ?? [];
      return paragraphs.find((paragraph) => paragraph.startIndex >= offset)?.paragraphStyle;
    };

    // Word's status bar: which page the cursor is on, how many pages there
    // are, the word count and the zoom level.
    let statusTimeout: ReturnType<typeof setTimeout> | undefined;
    const refreshStatus = async () => {
      const listener = statusListenerRef.current;
      if (!listener) return;
      const docModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(UniverInstanceType.UNIVER_DOC);
      if (!docModel) return;

      const skeleton = renderManagerService.getRenderUnitById(fDoc.getId())?.with(DocSkeletonManagerService)?.getSkeleton();
      const pages = skeleton?.getSkeletonData()?.pages ?? [];
      const cursor = docSelectionManagerService.getActiveTextRange()?.startOffset ?? 0;
      const currentIndex = pages.findIndex((page) => cursor >= page.st && cursor <= page.ed);

      let wordCount = 0;
      try {
        wordCount = (await docModel.getStatistics()).words;
      } catch {
        // Statistics are best-effort: an aborted run (fast typing) must not
        // blank out the rest of the status bar.
      }

      listener({
        name: fDoc.getName(),
        wordCount,
        pageCount: Math.max(pages.length, 1),
        currentPage: currentIndex >= 0 ? currentIndex + 1 : 1,
        zoom: Math.round((docModel.zoomRatio || 1) * 100),
      });
    };
    const scheduleStatusRefresh = () => {
      clearTimeout(statusTimeout);
      statusTimeout = setTimeout(() => void refreshStatus(), STATUS_REFRESH_DELAY_MS);
    };

    // Autosave: debounce so a fast typist doesn't hit localStorage on every
    // keystroke, and flush immediately on refresh/close so the last edit
    // isn't lost (React's unmount cleanup never runs on a hard refresh).
    let saveTimeout: ReturnType<typeof setTimeout> | undefined;
    const flushSave = () => saveSnapshot(STORAGE_KEY, fDoc.save());
    // Word shows its Table Design tab whenever the cursor is inside a
    // table. The caret's offset against the document's own table ranges is
    // the reliable test: the selection's node path is empty right after a
    // table mutation (a merge, say), and `textSelection$` alone misses
    // pointer-driven moves, so the selection operation Univer's own toolbar
    // items listen to drives this too.
    const isCursorInsideTable = (): boolean | null => {
      const docDataModel = univerInstanceService.getCurrentUnitOfType<DocumentDataModel>(
        UniverInstanceType.UNIVER_DOC,
      );
      if (resolveLiveTableRange(docSelectionManagerService, docDataModel)) return true;
      const offset = docSelectionManagerService.getActiveTextRange()?.startOffset;
      // No selection at all says nothing about where the user is (a table
      // mutation clears it), so the tab keeps whatever state it had.
      if (offset == null) return null;
      const tables = docDataModel?.getBody()?.tables;
      return Boolean(tables?.some((table) => offset > table.startIndex && offset < table.endIndex));
    };
    const refreshTableContext = () => {
      const inside = isCursorInsideTable();
      if (inside !== null) {
        wordRibbon.setTableContextActive(inside);
        contextService.setContextValue(WORD_CURSOR_IN_TABLE_CTX, inside);
      }
    };

    const commandSubscription = commandService.onCommandExecuted((command) => {
      // Using a table tool keeps the tab up even though the mutation clears
      // the cell selection it was applied to; the next selection change
      // decides again, exactly as in Word.
      if (command.id.startsWith("dockaro.command.table-")) wordRibbon.setTableContextActive(true);
      else if (command.id === SetTextSelectionsOperation.id) refreshTableContext();
      clearTimeout(saveTimeout);
      saveTimeout = setTimeout(flushSave, AUTOSAVE_DELAY_MS);
      scheduleStatusRefresh();
    });
    window.addEventListener("beforeunload", flushSave);

    const subscription = docSelectionManagerService.textSelection$.subscribe(() => {
      // Reflect the CURRENT selection exactly, like Word's Table Design tab:
      // show it only while the selection is actually inside a table, and
      // drop it the instant it isn't.
      refreshTableContext();
      scheduleStatusRefresh();
    });

    setReady(true);
    void refreshStatus();

    return () => {
      subscription.unsubscribe();
      commandSubscription.dispose();
      registrations.forEach((registration) => registration.dispose());
      wordRibbon.dispose();
      rulerPart.dispose();
      tableResize.dispose();
      tableMove.dispose();
      DataTransfer.prototype.getData = originalGetData;
      navigator.clipboard.read = originalClipboardRead;
      document.removeEventListener("paste", handleWordPasteCapture, true);
      pageChrome.dispose();
      dialogFocus.dispose();
      spellChecker.dispose();
      trackChanges.dispose();
      window.removeEventListener("beforeunload", flushSave);
      clearTimeout(saveTimeout);
      clearTimeout(statusTimeout);
      flushSave();
      clearRememberedTableRange();

      // univer.dispose() torn down while Univer's async preset init hasn't
      // yet reached its "steady" lifecycle stage (unmounting/navigating away
      // very quickly after mount) leaves an internal
      // firstValueFrom(lifecycle$...) with nothing left to emit once
      // disposal completes the source stream — RxJS rejects that with
      // EmptyError ("no elements in sequence"), surfaced by V8's async
      // stack traces as if thrown right here. Harmless: the instance is
      // being torn down either way. Swallow only this specific error so a
      // fast unmount doesn't crash the dev overlay / bubble as an uncaught
      // rejection, while any other dispose failure still surfaces.
      const swallowEmptyError = (event: PromiseRejectionEvent) => {
        if (event.reason?.name === "EmptyError") event.preventDefault();
      };
      window.addEventListener("unhandledrejection", swallowEmptyError);

      // Same race, different symptom: dispose() can synchronously unmount
      // an internal React root Univer owns (its own toolbar/canvas overlay)
      // while THIS component's own unmount is still mid-render for the same
      // commit. React reports that via console.error, not a thrown
      // exception, so the try/catch below can't see it — only a scoped
      // console.error filter can. Restored synchronously right after
      // dispose() returns, so no unrelated error in this window gets lost.
      const originalConsoleError = console.error;
      console.error = (...args: unknown[]) => {
        if (typeof args[0] === "string" && args[0].includes("synchronously unmount a root")) return;
        originalConsoleError(...args);
      };
      try {
        univer.dispose();
      } catch (err) {
        if ((err as Error)?.name !== "EmptyError") throw err;
      } finally {
        console.error = originalConsoleError;
        setTimeout(() => window.removeEventListener("unhandledrejection", swallowEmptyError), 0);
      }

      disposedRef.current = false;
      commandServiceRef.current = null;
      documentNameRef.current = () => {};
      rulerGeometryRef.current = () => null;
      setReady(false);
    };
  }, []);

  useImperativeHandle(apiRef, () => ({
    setName: (name: string) => documentNameRef.current(name),
    setZoom: (zoom: number) => {
      void commandServiceRef.current?.executeCommand(SetZoomCommandId, { value: zoom });
    },
    getRulerGeometry: () => rulerGeometryRef.current(),
    setIndents: (indents) => {
      void commandServiceRef.current?.executeCommand(SetIndentCommandId, indents);
    },
    setMargins: (margins) => {
      void commandServiceRef.current?.executeCommand(SetPageMarginsCommandId, margins);
    },
  }));

  return (
    <div ref={containerRef} className="relative h-full min-h-0 w-full flex-1">
      {ready && (
        <WordVerticalRuler
          getGeometry={() => rulerGeometryRef.current()}
          onMarginChange={(margins) => {
            void commandServiceRef.current?.executeCommand(SetPageMarginsCommandId, margins);
          }}
        />
      )}
      {pasteDialog && (
        <PasteDialog
          rawHtml={pasteDialog.rawHtml}
          plainText={pasteDialog.plainText}
          editorEl={pasteDialog.editorEl}
          pendingHtmlRef={pendingHtmlRef}
          pendingPlainRef={pendingPlainRef}
          onClose={() => setPasteDialog(null)}
        />
      )}
    </div>
  );
}
