// Angles are in degrees. Based on PPTXjs shapeArc().
export function shapeArc(
  cx: number,
  cy: number,
  rx: number,
  ry: number,
  startAngle: number,
  endAngle: number,
  isClose: boolean,
): string {
  const startRad = (startAngle * Math.PI) / 180;
  const endRad = (endAngle * Math.PI) / 180;

  const x1 = cx + rx * Math.cos(startRad);
  const y1 = cy + ry * Math.sin(startRad);
  const x2 = cx + rx * Math.cos(endRad);
  const y2 = cy + ry * Math.sin(endRad);

  // OOXML convention: always sweep clockwise from startAngle to endAngle.
  let sweepDeg = (((endAngle - startAngle) % 360) + 360) % 360;
  if (sweepDeg === 0 && startAngle !== endAngle) sweepDeg = 360;

  const largeArc = sweepDeg > 180 ? 1 : 0;
  const sweep = 1;

  let d = `M${x1},${y1} A${rx},${ry} 0 ${largeArc},${sweep} ${x2},${y2}`;
  if (isClose) {
    d += " Z";
  }
  return d;
}
