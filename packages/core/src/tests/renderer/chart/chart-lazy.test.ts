import type * as echarts from "echarts";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { buildPresentation } from "../../../model/presentation";
import { readPptx } from "../../../ooxml/zip";
import { BAR_CHART_XML, buildPptxWithChart } from "../../fixtures/chart-pptx";

const echartsChunk = {
  release: (): void => {},
  init: vi.fn((dom: HTMLElement) => ({
    setOption: vi.fn(),
    resize: vi.fn(),
    dispose: vi.fn(),
    isDisposed: () => false,
    getDom: () => dom,
  })),
};

async function renderChartSlide() {
  const { renderSlide } = await import("../../../renderer/slide");
  const buffer = await buildPptxWithChart(BAR_CHART_XML);
  const presentation = buildPresentation(await readPptx(buffer));
  const chartInstances = new Set<echarts.ECharts>();
  const handle = renderSlide(presentation, presentation.slides[0], {
    chartInstances,
  });
  document.body.appendChild(handle.element);
  return { handle, chartInstances };
}

describe("renderChart with ECharts loaded on demand", () => {
  beforeEach(() => {
    vi.resetModules();
    echartsChunk.init.mockClear();
    const loaded = new Promise<void>((resolve) => {
      echartsChunk.release = resolve;
    });
    vi.doMock("../../../renderer/chart/echarts", async () => {
      await loaded;
      return { init: echartsChunk.init };
    });
    vi.spyOn(HTMLElement.prototype, "offsetWidth", "get").mockReturnValue(640);
    vi.spyOn(HTMLElement.prototype, "offsetHeight", "get").mockReturnValue(480);
  });

  afterEach(() => {
    vi.doUnmock("../../../renderer/chart/echarts");
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it("skips a chart whose slide is disposed while ECharts loads", async () => {
    const { handle, chartInstances } = await renderChartSlide();
    await new Promise((resolve) => requestAnimationFrame(resolve));

    handle.dispose();
    echartsChunk.release();
    await handle.ready;

    expect(echartsChunk.init).not.toHaveBeenCalled();
    expect(chartInstances.size).toBe(0);
  });

  it("initializes the chart and clears the pending flag once ECharts loads", async () => {
    echartsChunk.release();
    const { handle, chartInstances } = await renderChartSlide();
    await handle.ready;

    expect(echartsChunk.init).toHaveBeenCalledTimes(1);
    expect(chartInstances.size).toBe(1);
    expect(
      handle.element.querySelector("[data-pptx-chart-pending]"),
    ).toBeNull();
    handle.dispose();
  });
});
