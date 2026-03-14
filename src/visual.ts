"use strict";

import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import PrimitiveValue = powerbi.PrimitiveValue;
import DataView = powerbi.DataView;
import DataViewCategorical = powerbi.DataViewCategorical;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;

import * as d3 from "d3";

import {
    createTooltipServiceWrapper,
    ITooltipServiceWrapper,
    TooltipEventArgs
} from "powerbi-visuals-utils-tooltiputils";

import { dataViewObjects } from "powerbi-visuals-utils-dataviewutils";

import { SettingsParser, VisualSettings } from "./settings";

type ChartRow = {
    group: string;
    series: string;
    actual: number;
    benchmark: number;
    actualHighlight?: number | null;
    benchmarkHighlight?: number | null;
    selectionId: any;
};

type SeriesStyle = {
    actualColor: string;
    actualOpacity: number;
    benchmarkColor: string;
    benchmarkOpacity: number;
};

export class Visual implements IVisual {
    private svg: d3.Selection<SVGSVGElement, unknown, null, undefined>;
    private root: d3.Selection<SVGGElement, unknown, null, undefined>;
    private host: IVisualHost;
    private selectionManager: powerbi.extensibility.ISelectionManager;
    private tooltipServiceWrapper: ITooltipServiceWrapper;

    private currentDataView: DataView | undefined;
    private settings: VisualSettings;
    private latestSeriesList: string[] = [];
    private seriesOverrides: Record<string, Partial<SeriesStyle>> = {};
    private seriesSelectors: Record<string, powerbi.data.Selector> = {};

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
        this.tooltipServiceWrapper = createTooltipServiceWrapper(
            options.host.tooltipService,
            options.element
        );

        this.svg = d3.select(options.element)
            .append("svg")
            .attr("tabindex", "0");

        this.root = this.svg.append("g");

        this.settings = {
            // Legend
            legendShow: true,
            legendPosition: "TopLeft",
            legendShowTitle: true,
            legendTitleText: "Category",
            legendLabelColor: "#666666",
            legendFontSize: 9,
            legendFontFamily: "Segoe UI",

            // X-axis
            xAxisShow: true,
            xAxisLabelOffset: 3,
            xAxisLabelColor: "#666666",
            xAxisFontSize: 9,
            xAxisFontFamily: "Segoe UI",
            xAxisShowTitle: true,
            xAxisTitleText: "",
            xAxisTitleColor: "#666666",
            xAxisTitleFontSize: 11,
            xAxisTitleFontFamily: "Segoe UI",

            // Y-axis
            yAxisShow: true,
            yAxisShowAxisLine: false,
            yAxisPosition: "Left",
            yAxisStart: null,
            yAxisEnd: null,
            yAxisLabelColor: "#666666",
            yAxisFontSize: 9,
            yAxisFontFamily: "Segoe UI",
            yAxisShowTitle: true,
            yAxisTitleText: "",
            yAxisTitleColor: "#666666",
            yAxisTitleFontSize: 11,
            yAxisTitleFontFamily: "Segoe UI",

            // Gridlines
            gridlinesShow: true,
            gridlinesColor: "#e6e6e6",
            gridlinesStrokeWidth: 1,
            gridlinesLineStyle: "dotted",

            // Columns
            seriesSelection: "All",
            actualColor: "",
            actualTransparency: 0,
            benchmarkColor: "",
            benchmarkTransparency: 70,
            actualWidth: 65,
            innerPadding: 20,
            groupPadding: 20,

            // Data labels
            dataLabelsShow: false,
            dataLabelsColor: "#333333",
            dataLabelsFontSize: 9,
            dataLabelsFontFamily: "Segoe UI",
            dataLabelsDisplayUnits: 0,
            dataLabelsValueDecimalPlaces: null,
            dataLabelsPosition: "Auto",

            // Benchmark labels
            showBenchmarkLabels: true,
            benchmarkLabelPrefix: "National Average: ",
            benchmarkLabelFontSize: 9,
            benchmarkLabelFontFamily: "Segoe UI"
        };
    }

    public update(options: VisualUpdateOptions): void {
        this.renderingStarted();

        this.currentDataView = options.dataViews?.[0];
        this.settings = SettingsParser.parse(this.currentDataView);

        const width = options.viewport.width;
        const height = options.viewport.height;

        this.svg
            .attr("width", width)
            .attr("height", height);

        this.root.selectAll("*").remove();

        if (!this.currentDataView?.categorical) {
            this.latestSeriesList = [];
            this.seriesOverrides = {};
            this.drawLandingMessage(width, height);
            this.renderingFinished();
            return;
        }

        const rows = this.getCategoricalRows(this.currentDataView);

        if (!rows.length) {
            this.latestSeriesList = [];
            this.seriesOverrides = {};
            this.seriesSelectors = {};
            this.drawLandingMessage(width, height);
            this.renderingFinished();
            return;
        }

        this.latestSeriesList = Array.from(new Set(rows.map(d => d.series))).sort();
        this.seriesOverrides = this.buildSeriesOverrides(this.currentDataView);

        this.drawChart(rows, width, height);
        this.renderingFinished();
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        const cards: powerbi.visuals.FormattingCard[] = [];

        // 1. Legend Card
        cards.push({
            displayName: "Legend",
            uid: "legend_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("legend", "show", this.settings.legendShow),
            groups: [{
                displayName: "Options",
                uid: "legend_options_group",
                slices: [
                    this.makeDropdownSlice("legend_position", "Position", "legend", "position", this.settings.legendPosition),
                    this.makeTextSlice("legend_titleText", "Title Text", "legend", "titleText", this.settings.legendTitleText),
                    this.makeColorSlice("legend_labelColor", "Color", "legend", "labelColor", this.settings.legendLabelColor),
                    this.makeNumericSlice("legend_fontSize", "Text Size", "legend", "fontSize", this.settings.legendFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("legend_fontFamily", "Font Family", "legend", "fontFamily", this.settings.legendFontFamily)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "legend", propertyName: "show" },
                { objectName: "legend", propertyName: "position" },
                { objectName: "legend", propertyName: "titleText" },
                { objectName: "legend", propertyName: "labelColor" },
                { objectName: "legend", propertyName: "fontSize" },
                { objectName: "legend", propertyName: "fontFamily" }
            ]
        } as any);

        // 2. X-axis Card (categoryAxis)
        cards.push({
            displayName: "X-axis",
            uid: "categoryAxis_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("categoryAxis", "show", this.settings.xAxisShow),
            groups: [{
                displayName: "Options",
                uid: "categoryAxis_options_group",
                slices: [
                    this.makeNumericSlice("categoryAxis_labelOffset", "Label offset", "categoryAxis", "labelOffset", this.settings.xAxisLabelOffset),
                    this.makeColorSlice("categoryAxis_labelColor", "Color", "categoryAxis", "labelColor", this.settings.xAxisLabelColor),
                    this.makeNumericSlice("categoryAxis_fontSize", "Text Size", "categoryAxis", "fontSize", this.settings.xAxisFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("categoryAxis_fontFamily", "Font Family", "categoryAxis", "fontFamily", this.settings.xAxisFontFamily),
                    this.makeBoolSlice("categoryAxis_showTitle", "Title", "categoryAxis", "showTitle", this.settings.xAxisShowTitle),
                    this.makeTextSlice("categoryAxis_titleText", "Title Text", "categoryAxis", "titleText", this.settings.xAxisTitleText),
                    this.makeColorSlice("categoryAxis_titleColor", "Title Color", "categoryAxis", "titleColor", this.settings.xAxisTitleColor),
                    this.makeNumericSlice("categoryAxis_titleFontSize", "Title Text Size", "categoryAxis", "titleFontSize", this.settings.xAxisTitleFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("categoryAxis_titleFontFamily", "Title Font Family", "categoryAxis", "titleFontFamily", this.settings.xAxisTitleFontFamily)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "categoryAxis", propertyName: "show" },
                { objectName: "categoryAxis", propertyName: "labelOffset" },
                { objectName: "categoryAxis", propertyName: "labelColor" },
                { objectName: "categoryAxis", propertyName: "fontSize" },
                { objectName: "categoryAxis", propertyName: "fontFamily" },
                { objectName: "categoryAxis", propertyName: "showTitle" },
                { objectName: "categoryAxis", propertyName: "titleText" },
                { objectName: "categoryAxis", propertyName: "titleColor" },
                { objectName: "categoryAxis", propertyName: "titleFontSize" },
                { objectName: "categoryAxis", propertyName: "titleFontFamily" }
            ]
        } as any);

        // 3. Y-axis Card (valueAxis)
        cards.push({
            displayName: "Y-axis",
            uid: "valueAxis_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("valueAxis", "show", this.settings.yAxisShow),
            groups: [{
                displayName: "Options",
                uid: "valueAxis_options_group",
                slices: [
                    this.makeDropdownSlice("valueAxis_position", "Position", "valueAxis", "position", this.settings.yAxisPosition),
                    this.makeBoolSlice("valueAxis_showAxisLine", "Show axis line", "valueAxis", "showAxisLine", this.settings.yAxisShowAxisLine),
                    this.makeNumericSlice("valueAxis_start", "Start", "valueAxis", "start", this.settings.yAxisStart),
                    this.makeNumericSlice("valueAxis_end", "End", "valueAxis", "end", this.settings.yAxisEnd),
                    this.makeColorSlice("valueAxis_labelColor", "Color", "valueAxis", "labelColor", this.settings.yAxisLabelColor),
                    this.makeNumericSlice("valueAxis_fontSize", "Text Size", "valueAxis", "fontSize", this.settings.yAxisFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("valueAxis_fontFamily", "Font Family", "valueAxis", "fontFamily", this.settings.yAxisFontFamily),
                    this.makeBoolSlice("valueAxis_showTitle", "Title", "valueAxis", "showTitle", this.settings.yAxisShowTitle),
                    this.makeTextSlice("valueAxis_titleText", "Title Text", "valueAxis", "titleText", this.settings.yAxisTitleText),
                    this.makeColorSlice("valueAxis_titleColor", "Title Color", "valueAxis", "titleColor", this.settings.yAxisTitleColor),
                    this.makeNumericSlice("valueAxis_titleFontSize", "Title Text Size", "valueAxis", "titleFontSize", this.settings.yAxisTitleFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("valueAxis_titleFontFamily", "Title Font Family", "valueAxis", "titleFontFamily", this.settings.yAxisTitleFontFamily)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "valueAxis", propertyName: "show" },
                { objectName: "valueAxis", propertyName: "position" },
                { objectName: "valueAxis", propertyName: "showAxisLine" },
                { objectName: "valueAxis", propertyName: "start" },
                { objectName: "valueAxis", propertyName: "end" },
                { objectName: "valueAxis", propertyName: "labelColor" },
                { objectName: "valueAxis", propertyName: "fontSize" },
                { objectName: "valueAxis", propertyName: "fontFamily" },
                { objectName: "valueAxis", propertyName: "showTitle" },
                { objectName: "valueAxis", propertyName: "titleText" },
                { objectName: "valueAxis", propertyName: "titleColor" },
                { objectName: "valueAxis", propertyName: "titleFontSize" },
                { objectName: "valueAxis", propertyName: "titleFontFamily" }
            ]
        } as any);

        // 4. Gridlines Card
        cards.push({
            displayName: "Gridlines",
            uid: "gridlines_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("gridlines", "show", this.settings.gridlinesShow),
            groups: [{
                displayName: "Options",
                uid: "gridlines_options_group",
                slices: [
                    this.makeColorSlice("gridlines_color", "Color", "gridlines", "color", this.settings.gridlinesColor),
                    this.makeNumericSlice("gridlines_strokeWidth", "Stroke width", "gridlines", "strokeWidth", this.settings.gridlinesStrokeWidth, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeDropdownSlice("gridlines_lineStyle", "Line Style", "gridlines", "lineStyle", this.settings.gridlinesLineStyle)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "gridlines", propertyName: "show" },
                { objectName: "gridlines", propertyName: "color" },
                { objectName: "gridlines", propertyName: "strokeWidth" },
                { objectName: "gridlines", propertyName: "lineStyle" }
            ]
        } as any);

        // 5. Columns Card
        const seriesItems: powerbi.IEnumMember[] = [{ value: "All", displayName: "All" }];
        
        this.latestSeriesList.forEach(series => {
            seriesItems.push({ value: series, displayName: series });
        });

        // Current selected series or "All"
        let selectedSeries = this.settings.seriesSelection;
        if (!seriesItems.find(i => i.value === selectedSeries)) {
            selectedSeries = "All";
        }

        const isAll = selectedSeries === "All";
        const targetColorActual = isAll ? this.settings.actualColor : this.getSeriesStyle(selectedSeries).actualColor;
        const targetTransparencyActual = isAll ? this.settings.actualTransparency : this.opacityToTransparency(this.getSeriesStyle(selectedSeries).actualOpacity);
        const targetColorBenchmark = isAll ? this.settings.benchmarkColor : this.getSeriesStyle(selectedSeries).benchmarkColor;
        const targetTransparencyBenchmark = isAll ? this.settings.benchmarkTransparency : this.opacityToTransparency(this.getSeriesStyle(selectedSeries).benchmarkOpacity);
        const targetSelector = isAll ? undefined : this.seriesSelectors[selectedSeries];

        const columnsSlices: any[] = [
            this.makeDropdownSlice("columns_seriesSelection", "Series", "columns", "seriesSelection", selectedSeries, seriesItems),
            this.makeColorSlice(
                isAll ? "columns_actualColor" : `columns_actualColor_${selectedSeries}`,
                "Actual colour",
                "columns",
                "actualColor",
                targetColorActual,
                targetSelector
            ),
            this.makeNumericSlice(
                isAll ? "columns_actualTransparency" : `columns_actualTransparency_${selectedSeries}`,
                "Actual transparency",
                "columns",
                "actualTransparency",
                targetTransparencyActual,
                targetSelector,
                { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 100 } }
            ),
            this.makeColorSlice(
                isAll ? "columns_benchmarkColor" : `columns_benchmarkColor_${selectedSeries}`,
                "Benchmark colour",
                "columns",
                "benchmarkColor",
                targetColorBenchmark,
                targetSelector
            ),
            this.makeNumericSlice(
                isAll ? "columns_benchmarkTransparency" : `columns_benchmarkTransparency_${selectedSeries}`,
                "Benchmark transparency",
                "columns",
                "benchmarkTransparency",
                targetTransparencyBenchmark,
                targetSelector,
                { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 100 } }
            )
        ];

        const layoutSlices: any[] = [
            this.makeNumericSlice("columns_actualWidth", "Actual width %", "columns", "actualWidth", this.settings.actualWidth, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 100 } }),
            this.makeNumericSlice("columns_innerPadding", "Inner padding", "columns", "innerPadding", this.settings.innerPadding, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
            this.makeNumericSlice("columns_groupPadding", "Group padding", "columns", "groupPadding", this.settings.groupPadding, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } })
        ];

        cards.push({
            displayName: "Columns",
            uid: "columns_card",
            analyticsPane: false,
            groups: [{
                displayName: "Apply settings to",
                uid: "columns_options_group",
                slices: columnsSlices
            }, {
                displayName: "Layout",
                uid: "columns_layout_group",
                slices: layoutSlices
            }],
            revertToDefaultDescriptors: [
                { objectName: "columns", propertyName: "seriesSelection" },
                { objectName: "columns", propertyName: "actualWidth" },
                { objectName: "columns", propertyName: "actualColor" },
                { objectName: "columns", propertyName: "actualTransparency" },
                { objectName: "columns", propertyName: "benchmarkColor" },
                { objectName: "columns", propertyName: "benchmarkTransparency" },
                { objectName: "columns", propertyName: "innerPadding" },
                { objectName: "columns", propertyName: "groupPadding" }
            ]
        } as any);

        // 6. Data Labels Card
        cards.push({
            displayName: "Data labels",
            uid: "dataLabels_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("dataLabels", "show", this.settings.dataLabelsShow),
            groups: [{
                displayName: "Options",
                uid: "dataLabels_options_group",
                slices: [
                    this.makeColorSlice("dataLabels_color", "Color", "dataLabels", "color", this.settings.dataLabelsColor),
                    this.makeNumericSlice("dataLabels_fontSize", "Text Size", "dataLabels", "fontSize", this.settings.dataLabelsFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("dataLabels_fontFamily", "Font Family", "dataLabels", "fontFamily", this.settings.dataLabelsFontFamily),
                    this.makeNumericSlice("dataLabels_displayUnits", "Display Units", "dataLabels", "displayUnits", this.settings.dataLabelsDisplayUnits),
                    this.makeNumericSlice("dataLabels_valueDecimalPlaces", "Value Decimal Places", "dataLabels", "valueDecimalPlaces", this.settings.dataLabelsValueDecimalPlaces, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeDropdownSlice("dataLabels_position", "Position", "dataLabels", "position", this.settings.dataLabelsPosition)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "dataLabels", propertyName: "show" },
                { objectName: "dataLabels", propertyName: "color" },
                { objectName: "dataLabels", propertyName: "fontSize" },
                { objectName: "dataLabels", propertyName: "fontFamily" },
                { objectName: "dataLabels", propertyName: "displayUnits" },
                { objectName: "dataLabels", propertyName: "valueDecimalPlaces" },
                { objectName: "dataLabels", propertyName: "position" }
            ]
        } as any);

        // 7. Benchmark labels Card
        cards.push({
            displayName: "Benchmark labels",
            uid: "benchmarkLabels_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("benchmarkLabels", "show", this.settings.showBenchmarkLabels),
            groups: [{
                displayName: "Options",
                uid: "benchmarkLabels_options_group",
                slices: [
                    this.makeTextSlice("benchmarkLabels_labelPrefix", "Label prefix", "benchmarkLabels", "labelPrefix", this.settings.benchmarkLabelPrefix),
                    this.makeNumericSlice("benchmarkLabels_fontSize", "Font size", "benchmarkLabels", "fontSize", this.settings.benchmarkLabelFontSize, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } }),
                    this.makeFontSlice("benchmarkLabels_fontFamily", "Font family", "benchmarkLabels", "fontFamily", this.settings.benchmarkLabelFontFamily)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "benchmarkLabels", propertyName: "show" },
                { objectName: "benchmarkLabels", propertyName: "labelPrefix" },
                { objectName: "benchmarkLabels", propertyName: "fontSize" },
                { objectName: "benchmarkLabels", propertyName: "fontFamily" }
            ]
        } as any);

        return { cards };
    }

    private getCategoricalRows(dataView: DataView): ChartRow[] {
        const categorical = dataView.categorical as DataViewCategorical;
        const categoryColumn = categorical.categories?.[0];
        const groupValues = (categoryColumn?.values || []).map(v => String(v ?? ""));
        const valueColumns = categorical.values || [];

        const rowMap = new Map<string, ChartRow>();

        for (let i = 0; i < valueColumns.length; i++) {
            const col = valueColumns[i];
            const seriesName = String(col.source.groupName ?? "Unknown");
            const measureName = String(col.source.displayName ?? "").toLowerCase();

            for (let r = 0; r < groupValues.length; r++) {
                const groupName = groupValues[r];
                const key = `${groupName}|||${seriesName}`;

                if (!rowMap.has(key)) {
                    const selectionId = this.host
                        .createSelectionIdBuilder()
                        .withCategory(categoryColumn!, r)
                        .createSelectionId();

                    rowMap.set(key, {
                        group: groupName,
                        series: seriesName,
                        actual: 0,
                        benchmark: 0,
                        actualHighlight: null,
                        benchmarkHighlight: null,
                        selectionId
                    });
                }

                const row = rowMap.get(key)!;
                const rawValue = Number((col.values?.[r] as PrimitiveValue) ?? 0);
                const rawHighlight = col.highlights
                    ? Number((col.highlights[r] as PrimitiveValue) ?? 0)
                    : null;

                if (measureName.includes("actual")) {
                    row.actual = rawValue;
                    row.actualHighlight = rawHighlight;
                } else if (measureName.includes("benchmark")) {
                    row.benchmark = rawValue;
                    row.benchmarkHighlight = rawHighlight;
                }
            }
        }

        return Array.from(rowMap.values());
    }

    private buildSeriesOverrides(dataView: DataView): Record<string, Partial<SeriesStyle>> {
        const overrides: Record<string, Partial<SeriesStyle>> = {};
        this.seriesSelectors = {};
        
        const categorical = dataView.categorical;
        if (!categorical || !categorical.values) return overrides;

        const grouped = categorical.values.grouped ? categorical.values.grouped() : [];

        // 1. Properly construct generic series selectors that the Format Pane requires
        grouped.forEach(group => {
            const seriesName = String(group.name ?? "Unknown");
            const selectionId = this.host.createSelectionIdBuilder()
                .withSeries(categorical.values, group)
                .createSelectionId();
            
            this.seriesSelectors[seriesName] = selectionId.getSelector() as powerbi.data.Selector;
            overrides[seriesName] = {};

            // Fallback: Sometimes objects are attached directly to the group
            if (group.objects) {
                this.applyStyleFromObjects(overrides[seriesName], group.objects);
            }
        });

        // 2. Extract properties from the column metadata where Power BI physically injects them
        for (let i = 0; i < categorical.values.length; i++) {
            const col = categorical.values[i];
            const seriesName = String(col.source.groupName ?? "Unknown");
            
            if (overrides[seriesName] && col.source.objects) {
                this.applyStyleFromObjects(overrides[seriesName], col.source.objects);
            }
        }

        return overrides;
    }

    private applyStyleFromObjects(style: Partial<SeriesStyle>, objects: powerbi.DataViewObjects): void {
        const actColor = dataViewObjects.getFillColor(objects, { objectName: "columns", propertyName: "actualColor" });
        const actTrans = dataViewObjects.getValue<number>(objects, { objectName: "columns", propertyName: "actualTransparency" });
        const benColor = dataViewObjects.getFillColor(objects, { objectName: "columns", propertyName: "benchmarkColor" });
        const benTrans = dataViewObjects.getValue<number>(objects, { objectName: "columns", propertyName: "benchmarkTransparency" });

        if (actColor) style.actualColor = actColor;
        if (typeof actTrans === "number") style.actualOpacity = this.transparencyToOpacity(actTrans);
        if (benColor) style.benchmarkColor = benColor;
        if (typeof benTrans === "number") style.benchmarkOpacity = this.transparencyToOpacity(benTrans);
    }

    private drawChart(data: ChartRow[], width: number, height: number): void {
        const margin = { top: 50, right: 20, bottom: 75, left: 60 };

        // Adjust margins based on Legend and Axis titles (Simplified approximation)
        if (this.settings.legendShow && this.settings.legendPosition.includes("Top")) margin.top += 30;
        if (this.settings.legendShow && this.settings.legendPosition.includes("Bottom")) margin.bottom += 30;
        if (this.settings.xAxisShowTitle) margin.bottom += 20;
        if (this.settings.yAxisShowTitle) margin.left += 20;

        const plotWidth = width - margin.left - margin.right;
        const plotHeight = height - margin.top - margin.bottom;

        this.root.attr("transform", `translate(${margin.left},${margin.top})`);

        const groups = Array.from(new Set(data.map(d => d.group)));
        const seriesList = Array.from(new Set(data.map(d => d.series)));

        const x0 = d3.scaleBand<string>()
            .domain(groups)
            .range([0, plotWidth])
            .paddingInner(Math.max(0, Math.min(0.99, this.settings.groupPadding / 100)));

        const x1 = d3.scaleBand<string>()
            .domain(seriesList)
            .range([0, x0.bandwidth()])
            .padding(Math.max(0, Math.min(0.99, this.settings.innerPadding / 100)));

        const maxActual = d3.max(data, d => this.getRenderedActual(d)) ?? 0;
        const maxBench = d3.max(data, d => this.getRenderedBenchmark(d)) ?? 0;
        const maxValueMax = Math.max(maxActual, maxBench) || 1;

        const yStart = this.settings.yAxisStart ?? 0;
        const yEnd = this.settings.yAxisEnd ?? (maxValueMax * 1.15);

        const y = d3.scaleLinear()
            .domain([yStart, yEnd])
            .nice()
            .range([plotHeight, 0]);

        // Gridlines
        if (this.settings.gridlinesShow) {
            const yAxisGrid = d3.axisLeft(y).tickSize(-plotWidth).tickFormat("" as any);
            const gridGroup = this.root.append("g")
                .attr("class", "gridlines")
                .call(yAxisGrid);

            gridGroup.selectAll("line")
                .style("stroke", this.settings.gridlinesColor)
                .style("stroke-width", `${this.settings.gridlinesStrokeWidth}px`)
                .style("stroke-dasharray", this.settings.gridlinesLineStyle === "dashed" ? "5,5" : this.settings.gridlinesLineStyle === "dotted" ? "2,2" : "none");
            gridGroup.select(".domain").remove();
        }

        // X Axis
        const xAxisGen = d3.axisBottom(x0).tickPadding(this.settings.xAxisLabelOffset);
        const xAxis = this.root.append("g")
            .attr("transform", `translate(0,${plotHeight})`)
            .call(xAxisGen);

        if (!this.settings.xAxisShow) {
            xAxis.style("display", "none");
        } else {
            xAxis.selectAll("text")
                .style("fill", this.settings.xAxisLabelColor)
                .style("font-size", `${this.settings.xAxisFontSize}px`)
                .style("font-family", this.settings.xAxisFontFamily);

            if (this.settings.xAxisShowTitle) {
                this.root.append("text")
                    .attr("x", plotWidth / 2)
                    .attr("y", plotHeight + margin.bottom - 10)
                    .attr("text-anchor", "middle")
                    .style("fill", this.settings.xAxisTitleColor)
                    .style("font-size", `${this.settings.xAxisTitleFontSize}px`)
                    .style("font-family", this.settings.xAxisTitleFontFamily)
                    .text(this.settings.xAxisTitleText || "Base Category");
            }
        }

        // Y Axis
        const yAxis = this.root.append("g")
            .attr("transform", this.settings.yAxisPosition === "Right" ? `translate(${plotWidth},0)` : "translate(0,0)")
            .call(this.settings.yAxisPosition === "Right" ? d3.axisRight(y) : d3.axisLeft(y));

        if (!this.settings.yAxisShow) {
            yAxis.style("display", "none");
        } else {
            if (!this.settings.yAxisShowAxisLine) {
                yAxis.select(".domain").remove();
            }

            yAxis.selectAll("text")
                .style("fill", this.settings.yAxisLabelColor)
                .style("font-size", `${this.settings.yAxisFontSize}px`)
                .style("font-family", this.settings.yAxisFontFamily);

            if (this.settings.yAxisShowTitle) {
                this.root.append("text")
                    .attr("transform", "rotate(-90)")
                    .attr("x", -plotHeight / 2)
                    .attr("y", this.settings.yAxisPosition === "Right" ? plotWidth + 40 : -margin.left + 15)
                    .attr("text-anchor", "middle")
                    .style("fill", this.settings.yAxisTitleColor)
                    .style("font-size", `${this.settings.yAxisTitleFontSize}px`)
                    .style("font-family", this.settings.yAxisTitleFontFamily)
                    .text(this.settings.yAxisTitleText || "Value");
            }
        }

        const actualWidthFactor = Math.max(0.1, Math.min(1, this.settings.actualWidth / 100));
        const labelFontSize = Math.max(8, this.settings.benchmarkLabelFontSize);
        const isHighContrast = !!this.host.colorPalette.isHighContrast;
        const foreground = this.host.colorPalette.foreground.value;
        const background = this.host.colorPalette.background.value;

        const groupLayer = this.root
            .selectAll<SVGGElement, string>(".groupLayer")
            .data(groups)
            .enter()
            .append("g")
            .attr("class", "groupLayer")
            .attr("transform", d => `translate(${x0(d) ?? 0},0)`);

        groupLayer.each((groupName, i, nodes) => {
            const groupData = data.filter(d => d.group === groupName);
            const g = d3.select(nodes[i]);

            const benchmarkBars = g.selectAll<SVGRectElement, ChartRow>(".benchmark")
                .data(groupData)
                .enter()
                .append("rect")
                .attr("class", "benchmark")
                .attr("x", d => x1(d.series) ?? 0)
                .attr("width", x1.bandwidth())
                .attr("y", d => y(this.getRenderedBenchmark(d)))
                .attr("height", d => Math.max(0, plotHeight - y(this.getRenderedBenchmark(d))))
                .attr("fill", d => isHighContrast ? background : this.getSeriesStyle(d.series).benchmarkColor)
                .attr("stroke", isHighContrast ? foreground : "none")
                .attr("stroke-width", isHighContrast ? 1.5 : 0)
                .attr("opacity", d => isHighContrast ? 1 : this.getSeriesStyle(d.series).benchmarkOpacity)
                .attr("tabindex", 0)
                .style("cursor", "pointer");

            const actualBars = g.selectAll<SVGRectElement, ChartRow>(".actual")
                .data(groupData)
                .enter()
                .append("rect")
                .attr("class", "actual")
                .attr("x", d => {
                    const bandX = x1(d.series) ?? 0;
                    const actualBarWidth = x1.bandwidth() * actualWidthFactor;
                    return bandX + (x1.bandwidth() - actualBarWidth) / 2;
                })
                .attr("width", x1.bandwidth() * actualWidthFactor)
                .attr("y", d => y(this.getRenderedActual(d)))
                .attr("height", d => Math.max(0, plotHeight - y(this.getRenderedActual(d))))
                .attr("fill", d => isHighContrast ? foreground : this.getSeriesStyle(d.series).actualColor)
                .attr("stroke", isHighContrast ? foreground : "none")
                .attr("stroke-width", isHighContrast ? 1 : 0)
                .attr("opacity", d => isHighContrast ? 1 : this.getSeriesStyle(d.series).actualOpacity)
                .attr("tabindex", 0)
                .style("cursor", "pointer");

            actualBars
                .on("click", async (event, d) => {
                    const multi = !!(event.ctrlKey || event.metaKey);
                    await this.selectionManager.select(d.selectionId, multi);
                    await this.applySelectionState();
                })
                .on("contextmenu", (event, d) => {
                    event.preventDefault();
                    this.selectionManager.showContextMenu(d.selectionId, { x: event.clientX, y: event.clientY });
                });

            this.tooltipServiceWrapper.addTooltip(actualBars as any, (args: TooltipEventArgs<ChartRow>) => this.getTooltipData(args.data, "Actual"), () => null);
            this.tooltipServiceWrapper.addTooltip(benchmarkBars as any, (args: TooltipEventArgs<ChartRow>) => this.getTooltipData(args.data, "Benchmark"), () => null);

            // Benchmark Labels (top of thick bar)
            if (this.settings.showBenchmarkLabels) {
                g.selectAll<SVGTextElement, ChartRow>(".benchmarkLabel")
                    .data(groupData)
                    .enter()
                    .append("text")
                    .attr("class", "benchmarkLabel")
                    .attr("x", d => (x1(d.series) ?? 0) + x1.bandwidth() / 2)
                    .attr("y", d => y(this.getRenderedBenchmark(d)) - 8)
                    .attr("text-anchor", "middle")
                    .style("font-size", `${labelFontSize}px`)
                    .style("font-family", this.settings.benchmarkLabelFontFamily)
                    .style("font-weight", "600")
                    .style("fill", isHighContrast ? foreground : null)
                    .text(d => `${this.settings.benchmarkLabelPrefix}${this.getRenderedBenchmark(d).toFixed(1)}`);
            }

            // Data Labels
            if (this.settings.dataLabelsShow) {
                g.selectAll<SVGTextElement, ChartRow>(".dataLabel")
                    .data(groupData)
                    .enter()
                    .append("text")
                    .attr("class", "dataLabel")
                    .attr("x", d => (x1(d.series) ?? 0) + x1.bandwidth() / 2)
                    .attr("y", d => {
                        const yPos = y(this.getRenderedActual(d));
                        if (this.settings.dataLabelsPosition === "InsideEnd") return yPos + 15;
                        if (this.settings.dataLabelsPosition === "InsideBase") return plotHeight - 5;
                        return yPos - 5; // OutsideEnd / Auto
                    })
                    .attr("text-anchor", "middle")
                    .style("font-size", `${this.settings.dataLabelsFontSize}px`)
                    .style("font-family", this.settings.dataLabelsFontFamily)
                    .style("fill", this.settings.dataLabelsColor)
                    .text(d => {
                        let val = this.getRenderedActual(d);
                        if (this.settings.dataLabelsDisplayUnits > 1) val = val / this.settings.dataLabelsDisplayUnits;
                        if (this.settings.dataLabelsValueDecimalPlaces !== null) return val.toFixed(this.settings.dataLabelsValueDecimalPlaces);
                        return val.toString();
                    });
            }

            // Series Labels under X-axis, used if not using legend natively or if requested by old code
            // Only draw if x-axis is showing and we're not using a standard legend
            if (this.settings.xAxisShow && !this.settings.legendShow) {
                g.selectAll<SVGTextElement, ChartRow>(".seriesLabel")
                    .data(groupData)
                    .enter()
                    .append("text")
                    .attr("class", "seriesLabel")
                    .attr("x", d => (x1(d.series) ?? 0) + x1.bandwidth() / 2)
                    .attr("y", plotHeight + 20)
                    .attr("text-anchor", "middle")
                    .style("font-size", `${Math.max(8, labelFontSize - 1)}px`)
                    .style("fill", isHighContrast ? foreground : null)
                    .text(d => d.series);
            }
        });

        // Draw Legend
        if (this.settings.legendShow && seriesList.length) {
            const legendGroup = this.root.append("g").attr("class", "legend");
            
            let yOffset = 10;
            if (this.settings.legendPosition.includes("Bottom")) {
                yOffset = plotHeight + margin.bottom - 20;
            } else {
                yOffset = -margin.top + 10;
            }

            let xAnchor = 0;
            if (this.settings.legendPosition.includes("Right")) {
                 legendGroup.attr("transform", `translate(${plotWidth - (seriesList.length * 80)}, ${yOffset})`);
            } else {
                 legendGroup.attr("transform", `translate(0, ${yOffset})`);
            }

            if (this.settings.legendShowTitle) {
                legendGroup.append("text")
                    .attr("x", 0)
                    .attr("y", 10)
                    .text(this.settings.legendTitleText)
                    .style("font-size", `${this.settings.legendFontSize}px`)
                    .style("font-family", this.settings.legendFontFamily)
                    .style("font-weight", "bold")
                    .style("fill", this.settings.legendLabelColor);
            }

            let xOffset = this.settings.legendShowTitle ? 50 : 0;
            seriesList.forEach(series => {
                const style = this.getSeriesStyle(series);

                legendGroup.append("rect")
                    .attr("x", xOffset)
                    .attr("y", 1)
                    .attr("width", 10)
                    .attr("height", 10)
                    .attr("fill", style.actualColor);

                legendGroup.append("text")
                    .attr("x", xOffset + 15)
                    .attr("y", 10)
                    .text(series)
                    .style("font-size", `${this.settings.legendFontSize}px`)
                    .style("font-family", this.settings.legendFontFamily)
                    .style("fill", this.settings.legendLabelColor);

                xOffset += 80;
            });
        }

        this.svg.on("click", async (event: MouseEvent) => {
            if (event.target === this.svg.node()) {
                await this.selectionManager.clear();
                await this.applySelectionState();
            }
        });

        void this.applySelectionState();
    }

    private getTooltipData(d: ChartRow, kind: "Actual" | "Benchmark"): VisualTooltipDataItem[] {
        const value = kind === "Actual" ? this.getRenderedActual(d) : this.getRenderedBenchmark(d);

        return [
            { displayName: "Base Category", value: d.group },
            { displayName: "Sub-Category", value: d.series },
            { displayName: kind === "Actual" ? "Actual Target Value" : "Benchmark Comparison Value", value: value.toString() }
        ];
    }

    private getRenderedActual(d: ChartRow): number {
        return d.actualHighlight != null ? d.actualHighlight : d.actual;
    }

    private getRenderedBenchmark(d: ChartRow): number {
        return d.benchmarkHighlight != null ? d.benchmarkHighlight : d.benchmark;
    }

    private getGlobalStyle(): SeriesStyle {
        return {
            actualColor: this.settings.actualColor,
            actualOpacity: this.transparencyToOpacity(this.settings.actualTransparency),
            benchmarkColor: this.settings.benchmarkColor,
            benchmarkOpacity: this.transparencyToOpacity(this.settings.benchmarkTransparency)
        };
    }

    private getSeriesStyle(series: string): SeriesStyle {
        const globalStyle = this.getGlobalStyle();
        const override = this.seriesOverrides[series] || {};

        const baseColor = this.host.colorPalette.getColor(series).value;

        return {
            actualColor: override.actualColor || globalStyle.actualColor || baseColor,
            actualOpacity: typeof override.actualOpacity === "number" ? override.actualOpacity : globalStyle.actualOpacity,
            benchmarkColor: override.benchmarkColor || globalStyle.benchmarkColor || baseColor,
            benchmarkOpacity: typeof override.benchmarkOpacity === "number" ? override.benchmarkOpacity : globalStyle.benchmarkOpacity
        };
    }

    private makeColorSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: string,
        selector?: powerbi.data.Selector
    ): any {
        const properties: any = {
            descriptor: {
                objectName: objectName,
                propertyName: propertyName,
                selector: selector
            }
        };

        if (value) {
            properties.value = { value: value };
        }

        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.ColorPicker,
                properties: properties
            }
        };
    }

    private makeNumericSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: number | null,
        selector?: powerbi.data.Selector,
        options?: any
    ): any {
        const properties: any = {
            descriptor: {
                objectName: objectName,
                propertyName: propertyName,
                selector: selector
            }
        };

        if (value !== null && value !== undefined) {
            properties.value = value;
        }

        if (options) {
            properties.options = options;
        }

        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.NumUpDown,
                properties: properties
            }
        };
    }

    private makeBoolSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: boolean
    ): any {
        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                properties: {
                    descriptor: {
                        objectName: objectName,
                        propertyName: propertyName
                    },
                    value: value
                }
            }
        };
    }

    private makeDropdownSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: string,
        items?: powerbi.IEnumMember[]
    ): any {
        // AutoDropdown requires primitive string value, ItemDropdown requires an IEnumMember object match.
        let finalValue: any = value;

        if (items) {
            const foundItem = items.find(i => i.value === value);
            if (foundItem) {
                finalValue = foundItem;
            } else {
                finalValue = { value: value, displayName: value };
            }
        }

        const properties: any = {
            descriptor: {
                objectName: objectName,
                propertyName: propertyName
            },
            value: finalValue
        };

        if (items) {
            properties.items = items;
        }

        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.Dropdown,
                properties: properties
            }
        };
    }

    private makeTextSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: string
    ): any {
        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.TextInput,
                properties: {
                    descriptor: {
                        objectName: objectName,
                        propertyName: propertyName
                    },
                    placeholder: "",
                    value: value
                }
            }
        };
    }

    private makeFontSlice(
        uid: string,
        displayName: string,
        objectName: string,
        propertyName: string,
        value: string
    ): any {
        return {
            uid: uid,
            displayName: displayName,
            control: {
                type: powerbi.visuals.FormattingComponent.FontPicker,
                properties: {
                    descriptor: {
                        objectName: objectName,
                        propertyName: propertyName
                    },
                    value: value
                }
            }
        };
    }

    private makeTopLevelToggle(
        objectName: string,
        propertyName: string,
        value: boolean
    ): any {
        return {
            uid: `${objectName}_${propertyName}_toggle`,
            suppressDisplayName: true,
            control: {
                type: powerbi.visuals.FormattingComponent.ToggleSwitch,
                properties: {
                    descriptor: {
                        objectName: objectName,
                        propertyName: propertyName
                    },
                    value: value
                }
            }
        };
    }

    private getFillColor(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string
    ): string | undefined {
        return (objects as any)?.[objectName]?.[propertyName]?.solid?.color;
    }

    private getNumericOptional(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string
    ): number | undefined {
        const value = (objects as any)?.[objectName]?.[propertyName];
        return typeof value === "number" ? value : undefined;
    }

    private transparencyToOpacity(transparencyPercent: number): number {
        return Math.max(0, Math.min(1, 1 - transparencyPercent / 100));
    }

    private opacityToTransparency(opacity: number): number {
        return Math.max(0, Math.min(100, Math.round((1 - opacity) * 100)));
    }

    private selectionIdToKey(id: any): string {
        if (!id) return "";
        if (typeof id.getKey === "function") return id.getKey();
        if (typeof id.getSelector === "function") return JSON.stringify(id.getSelector());
        if (id.key) return String(id.key);
        if (id.selector) return JSON.stringify(id.selector);
        return JSON.stringify(id);
    }

    private async applySelectionState(): Promise<void> {
        const ids = await this.selectionManager.getSelectionIds();
        const selectedKeys = ids.map(id => this.selectionIdToKey(id));
        const hasSelection = selectedKeys.length > 0;

        this.root
            .selectAll<SVGRectElement, ChartRow>(".benchmark")
            .attr("opacity", d => {
                const seriesOpacity = this.getSeriesStyle(d.series).benchmarkOpacity;

                if (!hasSelection) {
                    return seriesOpacity;
                }

                const rowKey = this.selectionIdToKey(d.selectionId);
                return selectedKeys.includes(rowKey) ? seriesOpacity : 0.25;
            });

        this.root
            .selectAll<SVGRectElement, ChartRow>(".actual")
            .attr("opacity", d => {
                const seriesOpacity = this.getSeriesStyle(d.series).actualOpacity;

                if (!hasSelection) {
                    return seriesOpacity;
                }

                const rowKey = this.selectionIdToKey(d.selectionId);
                return selectedKeys.includes(rowKey) ? seriesOpacity : 0.25;
            });
    }

    private drawLandingMessage(width: number, height: number): void {
        const g = this.root
            .append("g")
            .attr("transform", `translate(${width / 2},${height / 2})`);

        g.append("text")
            .attr("text-anchor", "middle")
            .style("font-size", "18px")
            .style("font-weight", "600")
            .text("Overlay Benchmark Bar");

        g.append("text")
            .attr("text-anchor", "middle")
            .attr("y", 26)
            .style("font-size", "12px")
            .text("Add Base Category, Sub-Category, Actual Target Value and Benchmark Comparison Value");
    }

    private toPropertySafeName(value: string): string {
        return value.replace(/[^a-zA-Z0-9_]/g, "_");
    }

    private renderingStarted(): void {
        try {
            this.host.eventService.renderingStarted({ name: "overlayBenchmarkBar" });
        } catch {
            // no-op
        }
    }

    private renderingFinished(): void {
        try {
            this.host.eventService.renderingFinished({ name: "overlayBenchmarkBar" });
        } catch {
            // no-op
        }
    }
}

function actualBarWidthFactor(value: number): number {
    return value;
}