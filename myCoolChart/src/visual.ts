"use strict";

import "./../style/visual.less";
import * as d3 from "d3";
import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;

export class Visual implements IVisual {

    private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private text: d3.Selection<SVGTextElement, unknown, null, undefined>;

    constructor(options: VisualConstructorOptions) {

        // create svg inside the visual container
        this.svg = d3.select(options.element)
            .append("svg")
            .classed("myCoolChartSvg", true);

        // create text element
        this.text = this.svg
            .append("text")
            .style("fill", "blue")
            .style("font-size", "24px")
            .text("Hello from MyCoolChart!");
    }

    public update(options: VisualUpdateOptions): void {

        const width = options.viewport.width;
        const height = options.viewport.height;

        // resize svg to full visual size
        this.svg
            .attr("width", width)
            .attr("height", height);

        // center the text
        this.text
            .attr("x", width / 2)
            .attr("y", height / 2)
            .attr("text-anchor", "middle")
            .attr("dominant-baseline", "central");
    }
}