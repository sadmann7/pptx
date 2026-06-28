/**
 * SVG shape renderers for OOXML preset geometry types.
 *
 * Each function receives the element's rendered width and height (in pt)
 * and returns JSX-ready SVG props. The SVG itself is rendered by ShapeElement.
 *
 * Coordinate convention: all paths are defined in a 100×100 space and
 * scaled via the SVG viewBox. This lets us define shapes once and reuse them
 * regardless of the element dimensions.
 */

export interface ShapePathProps {
  /** SVG element type to render */
  element: "rect" | "ellipse" | "polygon" | "path" | "line";
  /** Attributes forwarded to the SVG child element */
  attrs: Record<string, string | number>;
  /** Optional rx/ry for rounded corners on rect */
  rx?: number;
  ry?: number;
}

/**
 * Returns the SVG path/primitive descriptor for a given preset shape type.
 * Falls back to `rect` for unknown shape types.
 */
export function getShapePath(shapeType: string): ShapePathProps {
  switch (shapeType) {
    // ── Basic shapes ──────────────────────────────────────────────────────
    case "rect":
    case "squareTabs":
      return rect();

    case "ellipse":
    case "oval":
      return ellipse();

    case "roundRect":
      return { element: "rect", attrs: { x: 0, y: 0, width: 100, height: 100 }, rx: 8, ry: 8 };

    case "round1Rect":
      return { element: "rect", attrs: { x: 0, y: 0, width: 100, height: 100 }, rx: 8, ry: 0 };

    case "snip1Rect":
      return { element: "path", attrs: { d: "M 15 0 L 100 0 L 100 100 L 0 100 L 0 0 Z" } };

    case "triangle":
      return { element: "polygon", attrs: { points: "50,0 100,100 0,100" } };

    case "rtTriangle":
      return { element: "polygon", attrs: { points: "0,0 100,100 0,100" } };

    case "diamond":
      return { element: "polygon", attrs: { points: "50,0 100,50 50,100 0,50" } };

    case "parallelogram":
      return { element: "polygon", attrs: { points: "20,0 100,0 80,100 0,100" } };

    case "trapezoid":
      return { element: "polygon", attrs: { points: "15,0 85,0 100,100 0,100" } };

    case "pentagon":
      return { element: "polygon", attrs: { points: "50,0 100,38 81,100 19,100 0,38" } };

    case "hexagon":
      return { element: "polygon", attrs: { points: "25,0 75,0 100,50 75,100 25,100 0,50" } };

    case "octagon":
      return {
        element: "polygon",
        attrs: { points: "29,0 71,0 100,29 100,71 71,100 29,100 0,71 0,29" },
      };

    case "heptagon":
      return {
        element: "polygon",
        attrs: { points: "50,0 93,21 100,68 73,100 27,100 0,68 7,21" },
      };

    case "decagon":
      return {
        element: "polygon",
        attrs: {
          points: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
            .map((i) => {
              const angle = (Math.PI * 2 * i) / 10 - Math.PI / 2;
              return `${50 + 50 * Math.cos(angle)},${50 + 50 * Math.sin(angle)}`;
            })
            .join(" "),
        },
      };

    // ── Stars ─────────────────────────────────────────────────────────────
    case "star4":
      return { element: "polygon", attrs: { points: star(4, 50, 25) } };

    case "star5":
      return { element: "polygon", attrs: { points: star(5, 50, 21) } };

    case "star6":
      return { element: "polygon", attrs: { points: star(6, 50, 25) } };

    case "star8":
      return { element: "polygon", attrs: { points: star(8, 50, 20) } };

    case "star10":
      return { element: "polygon", attrs: { points: star(10, 50, 20) } };

    case "star12":
      return { element: "polygon", attrs: { points: star(12, 50, 25) } };

    case "star16":
      return { element: "polygon", attrs: { points: star(16, 50, 25) } };

    case "star24":
      return { element: "polygon", attrs: { points: star(24, 50, 30) } };

    case "star32":
      return { element: "polygon", attrs: { points: star(32, 50, 35) } };

    // ── Arrows ────────────────────────────────────────────────────────────
    case "rightArrow":
      return {
        element: "polygon",
        attrs: { points: "0,30 70,30 70,0 100,50 70,100 70,70 0,70" },
      };

    case "leftArrow":
      return {
        element: "polygon",
        attrs: { points: "100,30 30,30 30,0 0,50 30,100 30,70 100,70" },
      };

    case "upArrow":
      return {
        element: "polygon",
        attrs: { points: "30,100 30,30 0,30 50,0 100,30 70,30 70,100" },
      };

    case "downArrow":
      return {
        element: "polygon",
        attrs: { points: "30,0 30,70 0,70 50,100 100,70 70,70 70,0" },
      };

    case "leftRightArrow":
      return {
        element: "polygon",
        attrs: {
          points: "0,50 25,0 25,35 75,35 75,0 100,50 75,100 75,65 25,65 25,100",
        },
      };

    case "upDownArrow":
      return {
        element: "polygon",
        attrs: {
          points: "50,0 100,25 65,25 65,75 100,75 50,100 0,75 35,75 35,25 0,25",
        },
      };

    case "bentArrow":
    case "curvedRightArrow":
      return {
        element: "path",
        attrs: {
          d: "M 20 100 L 20 40 L 60 40 L 60 20 L 100 50 L 60 80 L 60 60 L 40 60 L 40 100 Z",
        },
      };

    // ── Callouts ──────────────────────────────────────────────────────────
    case "cloudCallout":
    case "wedgeRectCallout":
      return {
        element: "path",
        attrs: {
          d: "M 0 0 L 100 0 L 100 80 L 60 80 L 40 100 L 50 80 L 0 80 Z",
        },
      };

    // ── Process / flow ────────────────────────────────────────────────────
    case "flowChartProcess":
    case "process":
      return rect();

    case "flowChartDecision":
    case "decision":
      return { element: "polygon", attrs: { points: "50,0 100,50 50,100 0,50" } };

    case "flowChartTerminator":
    case "terminator":
      return { element: "rect", attrs: { x: 0, y: 0, width: 100, height: 100 }, rx: 50, ry: 50 };

    case "flowChartDocument":
      return {
        element: "path",
        attrs: { d: "M 0 0 L 100 0 L 100 85 Q 75 75 50 85 Q 25 95 0 85 Z" },
      };

    case "flowChartPredefinedProcess":
      return {
        element: "path",
        attrs: { d: "M 0 0 L 100 0 L 100 100 L 0 100 Z M 15 0 L 15 100 M 85 0 L 85 100" },
      };

    case "flowChartInternalStorage":
      return {
        element: "path",
        attrs: { d: "M 0 0 L 100 0 L 100 100 L 0 100 Z M 15 0 L 15 100 M 0 15 L 100 15" },
      };

    // ── Plus / cross ──────────────────────────────────────────────────────
    case "plus":
      return {
        element: "polygon",
        attrs: {
          points: "35,0 65,0 65,35 100,35 100,65 65,65 65,100 35,100 35,65 0,65 0,35 35,35",
        },
      };

    // ── Lines / connectors ────────────────────────────────────────────────
    case "line":
    case "straightConnector1":
      return { element: "line", attrs: { x1: 0, y1: 0, x2: 100, y2: 100 } };

    case "bentConnector2":
    case "bentConnector3":
      return {
        element: "path",
        attrs: { d: "M 0 0 L 50 0 L 50 100 L 100 100", fill: "none" },
      };

    case "curvedConnector2":
    case "curvedConnector3":
      return {
        element: "path",
        attrs: { d: "M 0 0 C 50 0 50 100 100 100", fill: "none" },
      };

    // ── Misc ──────────────────────────────────────────────────────────────
    case "cube":
      return {
        element: "path",
        attrs: {
          d: "M 15 0 L 100 0 L 100 85 L 85 100 L 0 100 L 0 15 Z M 0 15 L 85 15 L 100 0 M 85 15 L 85 100",
        },
      };

    case "can":
      return {
        element: "path",
        attrs: {
          d: "M 0 15 A 50 15 0 0 0 100 15 L 100 85 A 50 15 0 0 1 0 85 Z M 0 15 A 50 15 0 0 1 100 15",
        },
      };

    case "noSmoking":
      return {
        element: "path",
        attrs: {
          d: "M 50 0 A 50 50 0 1 1 50 100 A 50 50 0 0 1 50 0 M 15 85 L 85 15",
        },
      };

    case "custom":
    default:
      return rect();
  }
}

function rect(): ShapePathProps {
  return { element: "rect", attrs: { x: 0, y: 0, width: 100, height: 100 } };
}

function ellipse(): ShapePathProps {
  return { element: "ellipse", attrs: { cx: 50, cy: 50, rx: 50, ry: 50 } };
}

/**
 * Generate a star polygon with `n` points.
 * outerR and innerR are in the 0-100 coordinate space.
 */
function star(n: number, outerR: number, innerR: number): string {
  const points: string[] = [];
  for (let i = 0; i < n * 2; i++) {
    const r = i % 2 === 0 ? outerR : innerR;
    const angle = (Math.PI * i) / n - Math.PI / 2;
    points.push(`${50 + r * Math.cos(angle)},${50 + r * Math.sin(angle)}`);
  }
  return points.join(" ");
}
