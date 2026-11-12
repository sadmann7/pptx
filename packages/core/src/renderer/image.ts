import { renderCustomGeometry } from "../geometry/custom";
import { getPresetShapePath } from "../geometry/preset";
import { parseEmfContent } from "../media/emf";
import { renderPdfToImage } from "../media/pdf";
import {
  findMediaByTarget,
  findMediaByTargetAsync,
  getOrCreateBlobUrl,
  resolveMediaPath,
} from "../media/resolve";
import { PictureNodeData } from "../model/nodes/picture";
import { isExternalTargetMode, RelEntry } from "../ooxml/rel";
import { emuToPx } from "../ooxml/unit";
import { SafeXmlNode } from "../ooxml/xml";
import { hexToRgb } from "../utils/color";
import { isAllowedExternalMediaUrl, isAllowedExternalUrl } from "../utils/url";
import { RenderContext } from "./context";
import { resolveSlideNavigationIndex, slideJumpTitle } from "./navigation";
import { resolveColor, resolveFill, resolveLineStyle } from "./style";

function isUnsupportedFormat(path: string): boolean {
  const ext = path.split(".").pop()?.toLowerCase() || "";
  return ext === "wmf";
}

function isEmfFormat(path: string): boolean {
  const ext = path.split(".").pop()?.toLowerCase() || "";
  return ext === "emf";
}

let pictureClipPathIdCounter = 0;

function resolveImageRelUrl(rel: RelEntry, ctx: RenderContext): string | undefined {
  if (isExternalTargetMode(rel.targetMode)) {
    return isAllowedExternalMediaUrl(rel.target) ? rel.target : undefined;
  }

  const resolved = findMediaByTarget(rel.target, ctx.presentation.media);
  if (!resolved) return undefined;
  const { mediaPath, data } = resolved;
  if (isUnsupportedFormat(mediaPath)) return undefined;

  return getOrCreateBlobUrl(mediaPath, data, ctx.mediaUrlCache);
}

export function renderImage(node: PictureNodeData, ctx: RenderContext): HTMLElement {
  const wrapper = document.createElement("div");
  wrapper.style.position = "absolute";
  wrapper.style.left = `${node.position.x}px`;
  wrapper.style.top = `${node.position.y}px`;
  wrapper.style.width = `${node.size.w}px`;
  wrapper.style.height = `${node.size.h}px`;
  wrapper.style.overflow = "hidden";
  const geometryClipPath = getPictureGeometryClipPath(node);

  const transforms: string[] = [];
  if (node.rotation !== 0) {
    transforms.push(`rotate(${node.rotation}deg)`);
  }
  if (node.flipH && !geometryClipPath) {
    transforms.push("scaleX(-1)");
  }
  if (node.flipV && !geometryClipPath) {
    transforms.push("scaleY(-1)");
  }
  if (transforms.length > 0) {
    wrapper.style.transform = transforms.join(" ");
  }

  applyPictureShapeProperties(wrapper, node, ctx);

  if (node.isVideo) {
    renderVideo(node, ctx, wrapper);
    return wrapper;
  }

  if (node.isAudio) {
    renderAudio(node, ctx, wrapper);
    return wrapper;
  }

  const embedId = node.blipEmbed;
  let url: string | undefined;

  if (embedId) {
    const rel = ctx.slide.rels.get(embedId);
    if (!rel) {
      renderPlaceholder(wrapper, "Missing image reference");
      return wrapper;
    }

    if (isExternalTargetMode(rel.targetMode)) {
      url = isAllowedExternalMediaUrl(rel.target) ? rel.target : undefined;
      if (!url) {
        renderPlaceholder(wrapper, "Image not found");
        return wrapper;
      }
    } else {
      const mediaPathForType = resolveMediaPath(rel.target);
      if (isUnsupportedFormat(mediaPathForType)) {
        renderUnsupportedPlaceholder(wrapper, mediaPathForType);
        return wrapper;
      }

      const resolved = findMediaByTarget(rel.target, ctx.presentation.media);
      if (!resolved) {
        if (ctx.presentation.mediaResolver) {
          const task = findMediaByTargetAsync(
            rel.target,
            ctx.presentation.media,
            ctx.presentation.mediaResolver,
          )
            .then((lazyResolved) => {
              if (!lazyResolved) {
                renderPlaceholder(wrapper, "Image not found");
                return;
              }

              renderResolvedImage(node, ctx, wrapper, lazyResolved.mediaPath, lazyResolved.data);
            })
            .catch(() => {
              renderPlaceholder(wrapper, "Image not found");
            });
          ctx.asyncTasks?.push(task);
          if (!ctx.asyncTasks) void task;
          return wrapper;
        }

        renderPlaceholder(wrapper, "Image not found");
        return wrapper;
      }
      renderResolvedImage(node, ctx, wrapper, resolved.mediaPath, resolved.data);
      return wrapper;
    }
  } else if (node.blipLink) {
    url = resolveMediaUrl(node.blipLink, ctx);
    if (!url) {
      renderPlaceholder(wrapper, "Image not found");
      return wrapper;
    }
  } else {
    renderPlaceholder(wrapper, "No image data");
    return wrapper;
  }

  renderImageUrl(node, ctx, wrapper, url);
  return wrapper;
}

function renderResolvedImage(
  node: PictureNodeData,
  ctx: RenderContext,
  wrapper: HTMLElement,
  mediaPath: string,
  data: Uint8Array,
): void {
  if (isEmfFormat(mediaPath)) {
    if (!ctx.thumbnail) {
      const emfData = data instanceof Uint8Array ? data : new Uint8Array(data);
      renderEmf(emfData, node, ctx, wrapper, mediaPath);
    }
    return;
  }

  const url = getOrCreateBlobUrl(mediaPath, data, ctx.mediaUrlCache);
  renderImageUrl(node, ctx, wrapper, url);
}

function renderImageUrl(
  node: PictureNodeData,
  ctx: RenderContext,
  wrapper: HTMLElement,
  url: string,
): void {
  const blipFill = node.source.child("blipFill");
  const blip = blipFill.child("blip");
  const blipOpacity = resolveBlipOpacity(blip);

  const tile = blipFill.child("tile");
  if (tile.exists()) {
    wrapper.style.backgroundImage = `url("${url}")`;
    wrapper.style.backgroundRepeat = "repeat";
    wrapper.style.backgroundSize = "auto";
    if (blipOpacity < 1) {
      wrapper.style.opacity = `${Number(blipOpacity.toFixed(4))}`;
    }
    return;
  }

  const img = document.createElement("img");
  img.src = url;
  img.style.width = "100%";
  img.style.height = "100%";
  img.style.objectFit = "fill";
  img.style.display = "block";
  img.draggable = false;

  const fillRect = node.source.child("blipFill").child("stretch").child("fillRect");
  const fillRectBox = fillRect.exists() ? getFillRectBox(fillRect) : undefined;
  const geometryClipPath = getPictureGeometryClipPath(node);
  if (geometryClipPath) {
    renderClippedSvgImage(node, wrapper, url, geometryClipPath, fillRectBox);
    if (blipOpacity < 1) {
      wrapper.style.opacity = `${Number(blipOpacity.toFixed(4))}`;
    }
    return;
  }

  if (fillRectBox) {
    applyImageFillRect(img, fillRectBox);
  }

  // The srcRect remainder stretches to fill the shape's bounding box.
  if (node.crop) {
    const { top, right, bottom, left } = node.crop;
    const visibleW = 1 - left - right;
    const visibleH = 1 - top - bottom;
    if (visibleW > 0.001 && visibleH > 0.001) {
      const scaleX = 1 / visibleW;
      const scaleY = 1 / visibleH;
      // Pixels, because CSS margin percentages are relative to the containing block's width, even vertically.
      const wrapperW = node.size.w * ((fillRectBox?.width ?? 100) / 100);
      const wrapperH = node.size.h * ((fillRectBox?.height ?? 100) / 100);
      img.style.width = `${(scaleX * wrapperW).toFixed(4)}px`;
      img.style.height = `${(scaleY * wrapperH).toFixed(4)}px`;
      img.style.marginLeft = `${(-left * scaleX * wrapperW).toFixed(4)}px`;
      img.style.marginTop = `${(-top * scaleY * wrapperH).toFixed(4)}px`;
    }
  }

  if (blipOpacity < 1) {
    wrapper.style.opacity = `${Number(blipOpacity.toFixed(4))}`;
  }

  const grayscl = blip.child("grayscl");
  if (grayscl.exists()) {
    appendCssFilter(img, "grayscale(1)");
  }

  const duotone = blip.child("duotone");
  if (duotone.exists()) {
    applyDuotoneFilter(duotone, ctx, img, wrapper);
  }

  const lum = blip.child("lum");
  if (lum.exists()) {
    applyLumEffect(lum, img);
  }

  const biLevel = blip.child("biLevel");
  if (biLevel.exists()) {
    applyBiLevelEffect(biLevel, img);
  }

  wrapper.appendChild(img);
}

function pctAttr(node: SafeXmlNode, name: string): number {
  return (node.numAttr(name) ?? 0) / 1000;
}

interface FillRectBox {
  left: number;
  top: number;
  width: number;
  height: number;
}

function getFillRectBox(fillRect: SafeXmlNode): FillRectBox {
  const left = pctAttr(fillRect, "l");
  const top = pctAttr(fillRect, "t");
  const right = pctAttr(fillRect, "r");
  const bottom = pctAttr(fillRect, "b");

  return {
    left,
    top,
    width: 100 - left - right,
    height: 100 - top - bottom,
  };
}

function applyImageFillRect(img: HTMLImageElement, fillRect: FillRectBox): void {
  img.style.position = "absolute";
  img.style.left = `${fillRect.left}%`;
  img.style.top = `${fillRect.top}%`;
  img.style.width = `${fillRect.width}%`;
  img.style.height = `${fillRect.height}%`;
}

function getPictureGeometryClipPath(node: PictureNodeData): string | undefined {
  const sourceCustomGeometry = node.source.child("spPr").child("custGeom");
  const customGeometry =
    node.customGeometry ?? (sourceCustomGeometry.exists() ? sourceCustomGeometry : undefined);
  if (customGeometry?.exists()) {
    const extNode = node.source.child("spPr").child("xfrm").child("ext");
    const sourceExtentEmu = {
      w: extNode.numAttr("cx") ?? 0,
      h: extNode.numAttr("cy") ?? 0,
    };
    const d = renderCustomGeometry(customGeometry, node.size.w, node.size.h, sourceExtentEmu);
    return d || undefined;
  }

  const sourcePreset = node.source.child("spPr").child("prstGeom").attr("prst");
  const preset = node.presetGeometry ?? sourcePreset;
  if (!preset || preset === "rect") return undefined;
  const d = getPresetShapePath(preset, node.size.w, node.size.h);
  return d || undefined;
}

function getClipTransform(node: PictureNodeData): string | undefined {
  if (node.flipH && node.flipV) return `translate(${node.size.w} ${node.size.h}) scale(-1 -1)`;
  if (node.flipH) return `translate(${node.size.w} 0) scale(-1 1)`;
  if (node.flipV) return `translate(0 ${node.size.h}) scale(1 -1)`;
  return undefined;
}

function renderClippedSvgImage(
  node: PictureNodeData,
  wrapper: HTMLElement,
  url: string,
  clipPathD: string,
  fillRectBox?: FillRectBox,
): void {
  const svgNs = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(svgNs, "svg");
  svg.setAttribute("viewBox", `0 0 ${node.size.w} ${node.size.h}`);
  svg.setAttribute("width", "100%");
  svg.setAttribute("height", "100%");
  svg.style.display = "block";
  svg.style.overflow = "hidden";

  const clipId = `picture-clip-${pictureClipPathIdCounter++}`;
  const defs = document.createElementNS(svgNs, "defs");
  const clipPath = document.createElementNS(svgNs, "clipPath");
  clipPath.setAttribute("id", clipId);
  clipPath.setAttribute("clipPathUnits", "userSpaceOnUse");
  const path = document.createElementNS(svgNs, "path");
  path.setAttribute("d", clipPathD);
  const clipTransform = getClipTransform(node);
  if (clipTransform) {
    path.setAttribute("transform", clipTransform);
  }
  clipPath.appendChild(path);
  defs.appendChild(clipPath);
  svg.appendChild(defs);

  const image = document.createElementNS(svgNs, "image");
  image.setAttribute("href", url);
  image.setAttribute("preserveAspectRatio", "none");
  if (clipTransform) {
    image.setAttribute("transform", clipTransform);
  }
  const clippedImageGroup = document.createElementNS(svgNs, "g");
  clippedImageGroup.setAttribute("clip-path", `url(#${clipId})`);

  let x = ((fillRectBox?.left ?? 0) / 100) * node.size.w;
  let y = ((fillRectBox?.top ?? 0) / 100) * node.size.h;
  let width = ((fillRectBox?.width ?? 100) / 100) * node.size.w;
  let height = ((fillRectBox?.height ?? 100) / 100) * node.size.h;

  if (node.crop) {
    const { top, right, bottom, left } = node.crop;
    const visibleW = 1 - left - right;
    const visibleH = 1 - top - bottom;
    if (visibleW > 0.001 && visibleH > 0.001) {
      const scaleX = 1 / visibleW;
      const scaleY = 1 / visibleH;
      width *= scaleX;
      height *= scaleY;
      x += -left * scaleX * ((fillRectBox?.width ?? 100) / 100) * node.size.w;
      y += -top * scaleY * ((fillRectBox?.height ?? 100) / 100) * node.size.h;
    }
  }

  image.setAttribute("x", String(x));
  image.setAttribute("y", String(y));
  image.setAttribute("width", String(width));
  image.setAttribute("height", String(height));
  wrapper.appendChild(svg);
  svg.appendChild(clippedImageGroup);
  clippedImageGroup.appendChild(image);
}

function applyPictureShapeProperties(
  wrapper: HTMLElement,
  node: PictureNodeData,
  ctx: RenderContext,
): void {
  applyPictureFill(wrapper, node, ctx);
  applyPictureOutline(wrapper, node, ctx);
  applyPictureEffects(wrapper, node, ctx);
  applyPictureHyperlink(wrapper, node, ctx);
}

function applyPictureFill(wrapper: HTMLElement, node: PictureNodeData, ctx: RenderContext): void {
  const spPr = node.source.child("spPr");
  const fillCss = resolveFill(spPr, ctx);
  if (!fillCss) return;
  applyCssFillBackground(wrapper, fillCss);
}

function applyCssFillBackground(el: HTMLElement, fillCss: string): void {
  clearCssFillBackground(el);

  if (fillCss.includes("gradient") && fillCss.includes(" 0 0 / ")) {
    const bgMatch = fillCss.match(/,\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)\s*$/);
    if (bgMatch && bgMatch.index !== undefined) {
      const imageLayers = fillCss.slice(0, bgMatch.index).replace(/\s+0 0\s*\/\s*8px 8px/g, "");
      el.style.backgroundImage = imageLayers;
      el.style.backgroundSize = "8px 8px";
      el.style.backgroundRepeat = "repeat";
      el.style.backgroundColor = bgMatch[1];
      return;
    }
  }

  if (
    fillCss.includes("gradient") ||
    fillCss.startsWith("url(") ||
    fillCss.includes("repeating-")
  ) {
    el.style.background = fillCss;
  } else {
    el.style.backgroundColor = fillCss;
  }
}

function clearCssFillBackground(el: HTMLElement): void {
  el.style.background = "";
  el.style.backgroundColor = "";
  el.style.backgroundImage = "";
  el.style.backgroundRepeat = "";
  el.style.backgroundSize = "";
}

function applyPictureOutline(
  wrapper: HTMLElement,
  node: PictureNodeData,
  ctx: RenderContext,
): void {
  const lnRef = node.source.child("style").child("lnRef");
  const lineIsNoFill = node.line?.child("noFill").exists() ?? false;
  if (lineIsNoFill) return;

  const hasExplicitLine = node.line?.exists() ?? false;
  const themeLineFromLnRef =
    !hasExplicitLine &&
    lnRef.exists() &&
    (lnRef.numAttr("idx") ?? 0) > 0 &&
    (ctx.theme.lineStyles?.length ?? 0) >= (lnRef.numAttr("idx") ?? 0)
      ? ctx.theme.lineStyles![(lnRef.numAttr("idx") ?? 1) - 1]
      : undefined;
  const line = hasExplicitLine ? node.line! : themeLineFromLnRef;
  if (!line?.exists()) return;

  const style = resolveLineStyle(line, ctx, lnRef);
  if (style.width <= 0 || style.color === "transparent") return;

  wrapper.style.boxSizing = "border-box";
  wrapper.style.border = `${style.width}px ${style.dash} ${style.color}`;
}

function applyPictureEffects(
  wrapper: HTMLElement,
  node: PictureNodeData,
  ctx: RenderContext,
): void {
  const effectLst = node.source.child("spPr").child("effectLst");
  if (!effectLst.exists()) return;

  const outerShdw = effectLst.child("outerShdw");
  if (outerShdw.exists()) {
    applyPictureOuterShadow(wrapper, node, outerShdw, ctx);
  }

  const glow = effectLst.child("glow");
  if (glow.exists()) {
    applyPictureGlow(wrapper, glow, ctx);
  }

  const softEdge = effectLst.child("softEdge");
  if (softEdge.exists()) {
    applyPictureSoftEdge(wrapper, node, softEdge);
  }

  const reflection = effectLst.child("reflection");
  if (reflection.exists()) {
    applyPictureReflection(wrapper, reflection);
  }
}

function applyPictureOuterShadow(
  wrapper: HTMLElement,
  node: PictureNodeData,
  outerShdw: SafeXmlNode,
  ctx: RenderContext,
): void {
  const dir = outerShdw.numAttr("dir") ?? 0;
  const distPx = emuToPx(outerShdw.numAttr("dist") ?? 0);
  const blurPx = emuToPx(outerShdw.numAttr("blurRad") ?? 0);
  const dirDeg = dir / 60000;
  const offsetX = distPx * Math.cos((dirDeg * Math.PI) / 180);
  const offsetY = distPx * Math.sin((dirDeg * Math.PI) / 180);
  const shadowColor = resolveEffectColor(outerShdw, ctx, "rgba(0,0,0,0.4)");

  const sx = outerShdw.numAttr("sx");
  const sy = outerShdw.numAttr("sy");
  if (sx != null && sy != null && sx > 0 && sy > 0) {
    const scaleX = sx / 100000;
    const scaleY = sy / 100000;
    const spreadX = (node.size.w * (scaleX - 1)) / 2;
    const spreadY = (node.size.h * (scaleY - 1)) / 2;
    const spread = Math.max(0, (spreadX + spreadY) / 2);
    wrapper.style.boxShadow = `${offsetX.toFixed(1)}px ${offsetY.toFixed(1)}px ${blurPx.toFixed(1)}px ${spread.toFixed(1)}px ${shadowColor}`;
    return;
  }

  appendCssFilter(
    wrapper,
    `drop-shadow(${offsetX.toFixed(1)}px ${offsetY.toFixed(1)}px ${blurPx.toFixed(1)}px ${shadowColor})`,
  );
}

function resolveEffectColor(node: SafeXmlNode, ctx: RenderContext, fallback: string): string {
  const { color, alpha } = resolveColor(node, ctx);
  if (!color) return fallback;
  const hex = color.startsWith("#") ? color : `#${color}`;
  const { r, g, b } = hexToRgb(hex);
  return `rgba(${r},${g},${b},${alpha.toFixed(3)})`;
}

function appendCssFilter(el: HTMLElement, filter: string): void {
  const current = el.style.filter.trim();
  el.style.filter = current ? `${current} ${filter}` : filter;
}

function applyPictureGlow(wrapper: HTMLElement, glow: SafeXmlNode, ctx: RenderContext): void {
  const radiusPx = emuToPx(glow.numAttr("rad") ?? 0);
  if (!(radiusPx > 0)) return;

  const { color, alpha } = resolveColor(glow, ctx);
  if (!color || alpha <= 0) return;

  const hex = color.startsWith("#") ? color : `#${color}`;
  const { r, g, b } = hexToRgb(hex);
  appendCssFilter(
    wrapper,
    `drop-shadow(0px 0px ${radiusPx.toFixed(1)}px rgba(${r},${g},${b},${alpha.toFixed(3)}))`,
  );
}

function applyPictureSoftEdge(
  wrapper: HTMLElement,
  node: PictureNodeData,
  softEdge: SafeXmlNode,
): void {
  const radiusPx = emuToPx(softEdge.numAttr("rad") ?? 0);
  if (!(radiusPx > 0)) return;

  const maxRadius = Math.max(0, Math.min(node.size.w, node.size.h) / 2);
  const clampedRadiusPx = maxRadius > 0 ? Math.min(radiusPx, maxRadius) : radiusPx;
  const radiusCss = `${Number(clampedRadiusPx.toFixed(4))}px`;
  const radiusVar = "var(--pptx-soft-edge-radius)";
  const maskImage = [
    `linear-gradient(to right, transparent 0, black ${radiusVar}, black calc(100% - ${radiusVar}), transparent 100%)`,
    `linear-gradient(to bottom, transparent 0, black ${radiusVar}, black calc(100% - ${radiusVar}), transparent 100%)`,
  ].join(", ");

  wrapper.style.setProperty("--pptx-soft-edge-radius", radiusCss);
  wrapper.style.setProperty("-webkit-mask-image", maskImage);
  wrapper.style.setProperty("mask-image", maskImage);
  wrapper.style.setProperty("-webkit-mask-size", "100% 100%");
  wrapper.style.setProperty("mask-size", "100% 100%");
  wrapper.style.setProperty("-webkit-mask-repeat", "no-repeat");
  wrapper.style.setProperty("mask-repeat", "no-repeat");
  wrapper.style.setProperty("-webkit-mask-composite", "source-in");
  wrapper.style.setProperty("mask-composite", "intersect");
  const style = wrapper.style as CSSStyleDeclaration & {
    webkitMaskImage?: string;
    webkitMaskComposite?: string;
    maskImage?: string;
    maskComposite?: string;
  };
  style.webkitMaskImage = maskImage;
  style.maskImage = maskImage;
  style.webkitMaskComposite = "source-in";
  style.maskComposite = "intersect";
}

function applyPictureReflection(wrapper: HTMLElement, reflection: SafeXmlNode): void {
  const dist = emuToPx(reflection.numAttr("dist") ?? 0);
  const stA = (reflection.numAttr("stA") ?? 50000) / 100000;
  const endA = (reflection.numAttr("endA") ?? 0) / 100000;
  const stPos = Math.max(0, Math.min(100, (reflection.numAttr("stPos") ?? 0) / 1000));
  const endPos = Math.max(0, Math.min(100, (reflection.numAttr("endPos") ?? 100000) / 1000));
  const mask = `linear-gradient(to bottom, rgba(255,255,255,${stA.toFixed(3)}) ${stPos.toFixed(1)}%, rgba(255,255,255,${endA.toFixed(3)}) ${endPos.toFixed(1)}%)`;
  const reflectValue = `below ${dist.toFixed(1)}px ${mask}`;
  wrapper.style.setProperty("-webkit-box-reflect", reflectValue);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (wrapper.style as any).webkitBoxReflect = reflectValue;
}

function applyPictureHyperlink(
  wrapper: HTMLElement,
  node: PictureNodeData,
  ctx: RenderContext,
): void {
  if (!node.hlinkClick || !ctx.onNavigate) return;

  const { action, rId } = node.hlinkClick;
  const rel = rId ? ctx.slide.rels.get(rId) : undefined;
  const slideIndex = resolveSlideNavigationIndex(ctx, action, rel);
  if (slideIndex !== undefined) {
    wrapper.style.cursor = "pointer";
    wrapper.title = node.hlinkClick.tooltip || slideJumpTitle(slideIndex);
    wrapper.addEventListener("click", (e) => {
      e.stopPropagation();
      ctx.onNavigate!({ slideIndex });
    });
    return;
  }

  if (!rId) return;
  if (!rel || !isExternalTargetMode(rel.targetMode) || !isAllowedExternalUrl(rel.target)) return;

  wrapper.style.cursor = "pointer";
  wrapper.title = node.hlinkClick.tooltip || rel.target;
  wrapper.addEventListener("click", (e) => {
    e.stopPropagation();
    ctx.onNavigate!({ url: rel.target });
  });
}

function resolveBlipOpacity(blip: SafeXmlNode): number {
  let alpha = 1;

  const alphaModFix = blip.child("alphaModFix");
  if (alphaModFix.exists()) {
    alpha *= (alphaModFix.numAttr("amt") ?? 100000) / 100000;
  }

  const alphaMod = blip.child("alphaMod");
  if (alphaMod.exists()) {
    alpha *= (alphaMod.numAttr("val") ?? 100000) / 100000;
  }

  const alphaOff = blip.child("alphaOff");
  if (alphaOff.exists()) {
    alpha += (alphaOff.numAttr("val") ?? 0) / 100000;
  }

  return Math.max(0, Math.min(1, alpha));
}

function renderVideo(node: PictureNodeData, ctx: RenderContext, wrapper: HTMLElement): void {
  const videoUrl = resolveMediaUrl(node.mediaRId, ctx);

  let posterUrl: string | undefined;
  if (node.blipEmbed) {
    const rel = ctx.slide.rels.get(node.blipEmbed);
    if (rel) {
      posterUrl = resolveImageRelUrl(rel, ctx);
    }
  }

  if (videoUrl) {
    const video = document.createElement("video");
    video.src = videoUrl;
    video.preload = "none";
    video.controls = true;
    video.style.width = "100%";
    video.style.height = "100%";
    video.style.objectFit = "contain";
    video.style.backgroundColor = "#000";
    if (posterUrl) {
      video.poster = posterUrl;
    }
    wrapper.appendChild(video);
  } else if (posterUrl) {
    const img = document.createElement("img");
    img.src = posterUrl;
    img.style.width = "100%";
    img.style.height = "100%";
    img.style.objectFit = "fill";
    wrapper.appendChild(img);

    const overlay = document.createElement("div");
    overlay.style.position = "absolute";
    overlay.style.inset = "0";
    overlay.style.display = "flex";
    overlay.style.alignItems = "center";
    overlay.style.justifyContent = "center";
    overlay.style.backgroundColor = "rgba(0,0,0,0.3)";
    overlay.style.color = "#fff";
    overlay.style.fontSize = "24px";
    overlay.textContent = "\u25B6";
    wrapper.appendChild(overlay);
  } else {
    renderPlaceholder(wrapper, "Video");
  }
}

function renderAudio(node: PictureNodeData, ctx: RenderContext, wrapper: HTMLElement): void {
  const audioUrl = resolveMediaUrl(node.mediaRId, ctx);

  if (audioUrl) {
    if (node.blipEmbed) {
      const rel = ctx.slide.rels.get(node.blipEmbed);
      if (rel) {
        const cached = resolveImageRelUrl(rel, ctx);
        if (cached) {
          const img = document.createElement("img");
          img.src = cached;
          img.style.width = "100%";
          img.style.height = "calc(100% - 32px)";
          img.style.objectFit = "contain";
          wrapper.appendChild(img);
        }
      }
    }

    const audio = document.createElement("audio");
    audio.src = audioUrl;
    audio.preload = "none";
    audio.controls = true;
    audio.style.width = "100%";
    audio.style.position = "absolute";
    audio.style.bottom = "0";
    audio.style.left = "0";
    wrapper.appendChild(audio);
  } else {
    renderPlaceholder(wrapper, "Audio");
  }
}

function resolveMediaUrl(rId: string | undefined, ctx: RenderContext): string | undefined {
  if (!rId) return undefined;

  const rel = ctx.slide.rels.get(rId);
  if (!rel) return undefined;

  if (isExternalTargetMode(rel.targetMode)) {
    return isAllowedExternalMediaUrl(rel.target) ? rel.target : undefined;
  }

  const resolved = findMediaByTarget(rel.target, ctx.presentation.media);
  if (!resolved) return undefined;
  const { mediaPath, data } = resolved;

  return getOrCreateBlobUrl(mediaPath, data, ctx.mediaUrlCache);
}

function renderPlaceholder(wrapper: HTMLElement, message: string): void {
  const placeholder = document.createElement("div");
  placeholder.style.width = "100%";
  placeholder.style.height = "100%";
  placeholder.style.display = "flex";
  placeholder.style.alignItems = "center";
  placeholder.style.justifyContent = "center";
  placeholder.style.backgroundColor = "#f0f0f0";
  placeholder.style.color = "#888";
  placeholder.style.fontSize = "12px";
  placeholder.style.border = "1px dashed #ccc";
  placeholder.textContent = message;
  wrapper.appendChild(placeholder);
}

function renderUnsupportedPlaceholder(wrapper: HTMLElement, path: string): void {
  const ext = path.split(".").pop()?.toUpperCase() || "Unknown";
  const placeholder = document.createElement("div");
  placeholder.style.width = "100%";
  placeholder.style.height = "100%";
  placeholder.style.display = "flex";
  placeholder.style.flexDirection = "column";
  placeholder.style.alignItems = "center";
  placeholder.style.justifyContent = "center";
  placeholder.style.backgroundColor = "#f5f5f5";
  placeholder.style.color = "#999";
  placeholder.style.fontSize = "11px";
  placeholder.style.border = "1px dashed #ddd";

  const icon = document.createElement("div");
  icon.style.fontSize = "24px";
  icon.style.marginBottom = "4px";
  icon.textContent = "\uD83D\uDDBC";

  const label = document.createElement("div");
  label.textContent = `Unsupported format: ${ext}`;

  placeholder.appendChild(icon);
  placeholder.appendChild(label);
  wrapper.appendChild(placeholder);
}

function renderEmf(
  data: Uint8Array,
  node: PictureNodeData,
  ctx: RenderContext,
  wrapper: HTMLElement,
  mediaPath: string,
): void {
  const content = parseEmfContent(data);

  switch (content.type) {
    case "pdf":
      renderEmfPdf(content.data, wrapper, node, ctx, mediaPath);
      break;
    case "bitmap":
      renderEmfBitmap(content.imageData, wrapper, ctx, mediaPath);
      break;
    case "empty":
      break;
    case "unsupported":
      // A placeholder would add artifacts that PowerPoint and PDF exports don't show.
      break;
  }
}

function renderEmfPdf(
  pdfData: Uint8Array,
  wrapper: HTMLElement,
  node: PictureNodeData,
  ctx: RenderContext,
  mediaPath: string,
): void {
  const cacheKey = `${mediaPath}:emf-pdf`;
  const cached = ctx.mediaUrlCache.get(cacheKey);
  if (cached) {
    wrapper.appendChild(createFillImage(cached));
    return;
  }

  const task = renderPdfToImage(pdfData, node.size.w, node.size.h, ctx.pdfjs)
    .then((url) => {
      if (url) {
        ctx.mediaUrlCache.set(cacheKey, url);
        wrapper.appendChild(createFillImage(url));
      }
    })
    .catch(() => {
      // A failed decode leaves the wrapper transparent.
    });
  ctx.asyncTasks?.push(task);
}

function renderEmfBitmap(
  imageData: ImageData,
  wrapper: HTMLElement,
  ctx: RenderContext,
  mediaPath: string,
): void {
  const cacheKey = `${mediaPath}:emf-bitmap`;
  const cached = ctx.mediaUrlCache.get(cacheKey);
  if (cached) {
    wrapper.appendChild(createFillImage(cached));
    return;
  }

  const canvas = document.createElement("canvas");
  canvas.width = imageData.width;
  canvas.height = imageData.height;
  const canvasCtx = canvas.getContext("2d");
  if (!canvasCtx) return;

  canvasCtx.putImageData(imageData, 0, 0);
  const task = new Promise<void>((resolve) => {
    canvas.toBlob((blob) => {
      if (blob) {
        const url = URL.createObjectURL(blob);
        ctx.mediaUrlCache.set(cacheKey, url);
        wrapper.appendChild(createFillImage(url));
      }
      resolve();
    }, "image/png");
  });
  ctx.asyncTasks?.push(task);
}

function createFillImage(url: string): HTMLImageElement {
  const img = document.createElement("img");
  img.src = url;
  img.style.width = "100%";
  img.style.height = "100%";
  img.style.objectFit = "fill";
  img.style.display = "block";
  img.draggable = false;
  return img;
}

function applyDuotoneFilter(
  duotone: SafeXmlNode,
  ctx: RenderContext,
  img: HTMLImageElement,
  _wrapper: HTMLElement,
): void {
  // `a:duotone` lists the dark color first, then the light one.
  const colorChildren = duotone.allChildren();
  if (colorChildren.length < 2) return;

  const { color: c1 } = resolveColor(colorChildren[0], ctx);
  const { color: c2 } = resolveColor(colorChildren[1], ctx);
  if (!c1 || !c2) return;

  const hex1 = c1.startsWith("#") ? c1 : `#${c1}`;
  const hex2 = c2.startsWith("#") ? c2 : `#${c2}`;
  const rgb1 = hexToRgb(hex1);
  const rgb2 = hexToRgb(hex2);

  const apply = () => {
    const w = img.naturalWidth;
    const h = img.naturalHeight;
    if (!w || !h) return;

    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const c = canvas.getContext("2d");
    if (!c) return;

    c.drawImage(img, 0, 0);
    const imageData = c.getImageData(0, 0, w, h);
    const data = imageData.data;

    for (let i = 0; i < data.length; i += 4) {
      const gray = (0.2126 * data[i] + 0.7152 * data[i + 1] + 0.0722 * data[i + 2]) / 255;
      data[i] = Math.round(rgb1.r + (rgb2.r - rgb1.r) * gray);
      data[i + 1] = Math.round(rgb1.g + (rgb2.g - rgb1.g) * gray);
      data[i + 2] = Math.round(rgb1.b + (rgb2.b - rgb1.b) * gray);
    }

    c.putImageData(imageData, 0, 0);
    img.src = canvas.toDataURL();
  };

  if (img.complete && img.naturalWidth) {
    apply();
  } else {
    img.addEventListener("load", apply, { once: true });
  }
}

// `bright` is an additive offset and `contrast` a multiplier, both in 1/100000 units.
function applyLumEffect(lum: SafeXmlNode, img: HTMLImageElement): void {
  const bright = (lum.numAttr("bright") ?? 0) / 100000;
  const contrast = (lum.numAttr("contrast") ?? 0) / 100000;

  if (bright === 0 && contrast === 0) return;

  const apply = () => {
    const w = img.naturalWidth;
    const h = img.naturalHeight;
    if (!w || !h) return;

    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const c = canvas.getContext("2d");
    if (!c) return;

    c.drawImage(img, 0, 0);
    const imageData = c.getImageData(0, 0, w, h);
    const data = imageData.data;

    for (let i = 0; i < data.length; i += 4) {
      for (let ch = 0; ch < 3; ch++) {
        let v = data[i + ch] / 255;
        if (contrast !== 0) {
          v = 0.5 + (v - 0.5) * (1 + contrast);
        }
        v += bright;
        data[i + ch] = Math.round(Math.max(0, Math.min(255, v * 255)));
      }
    }

    c.putImageData(imageData, 0, 0);
    img.src = canvas.toDataURL();
  };

  if (img.complete && img.naturalWidth) {
    apply();
  } else {
    img.addEventListener("load", apply, { once: true });
  }
}

function applyBiLevelEffect(biLevel: SafeXmlNode, img: HTMLImageElement): void {
  const thresh = (biLevel.numAttr("thresh") ?? 50000) / 100000;

  const apply = () => {
    const w = img.naturalWidth;
    const h = img.naturalHeight;
    if (!w || !h) return;

    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const c = canvas.getContext("2d");
    if (!c) return;

    c.drawImage(img, 0, 0);
    const imageData = c.getImageData(0, 0, w, h);
    const data = imageData.data;

    for (let i = 0; i < data.length; i += 4) {
      const gray = (0.2126 * data[i] + 0.7152 * data[i + 1] + 0.0722 * data[i + 2]) / 255;
      const val = gray >= thresh ? 255 : 0;
      data[i] = val;
      data[i + 1] = val;
      data[i + 2] = val;
    }

    c.putImageData(imageData, 0, 0);
    img.src = canvas.toDataURL();
  };

  if (img.complete && img.naturalWidth) {
    apply();
  } else {
    img.addEventListener("load", apply, { once: true });
  }
}
