/**
 * DOM data-attribute contract written by the slide renderer.
 *
 * Consumers (e.g. the React selection overlay) use these to locate shapes,
 * paragraphs, and runs in rendered slide DOM, and to map edited content back
 * to model indices. Treat the values as a public API: renaming one is a
 * breaking change for anything selecting against rendered output.
 */
export const PPTX_ATTRS = {
  /** Shape/node wrapper element; the value is the node id. */
  nodeId: "data-pptx-node-id",
  /** Paragraph div inside a text body; the value is the source paragraph index. */
  paragraph: "data-pptx-p",
  /** Styled run span inside a paragraph; the value is the source run index. */
  run: "data-pptx-r",
  /** Bullet marker span inside a paragraph (excluded from text read-back). */
  bullet: "data-pptx-bullet",
  /** Placeholder prompt overlay ("Click to add text"). */
  placeholderPrompt: "data-pptx-placeholder-prompt",
  /** Wrapper of an empty placeholder shape. */
  placeholderEmpty: "data-pptx-placeholder-empty",
} as const;

/**
 * `element.dataset` keys corresponding to {@link PPTX_ATTRS} (camelCase per
 * the HTML dataset API), for reads like `el.dataset[PPTX_DATASET.run]`.
 */
export const PPTX_DATASET = {
  nodeId: "pptxNodeId",
  paragraph: "pptxP",
  run: "pptxR",
  bullet: "pptxBullet",
} as const;
