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
    private latestGroupList: string[] = [];
    private groupOverrides: Record<string, Partial<SeriesStyle>> = {};
    private groupSelectors: Record<string, powerbi.data.Selector> = {};
    private hasSeries: boolean = false;

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
            legendBold: false,
            legendItalic: false,
            legendUnderline: false,

            // X-axis
            xAxisShow: true,
            xAxisLabelOffset: 3,
            xAxisPadding: 0,
            xAxisShowValues: true,
            xAxisLabelColor: "#666666",
            xAxisFontSize: 9,
            xAxisFontFamily: "Segoe UI",
            xAxisBold: false,
            xAxisItalic: false,
            xAxisUnderline: false,
            xAxisShowTitle: true,
            xAxisTitleText: "",
            xAxisTitleColor: "#666666",
            xAxisTitleFontSize: 11,
            xAxisTitleFontFamily: "Segoe UI",
            xAxisTitleBold: false,
            xAxisTitleItalic: false,
            xAxisTitleUnderline: false,

            // Y-axis
            yAxisShow: true,
            yAxisShowAxisLine: false,
            yAxisPosition: "Left",
            yAxisStart: null,
            yAxisEnd: null,
            yAxisPadding: 0,
            yAxisShowValues: true,
            yAxisLabelColor: "#666666",
            yAxisFontSize: 9,
            yAxisFontFamily: "Segoe UI",
            yAxisBold: false,
            yAxisItalic: false,
            yAxisUnderline: false,
            yAxisShowTitle: true,
            yAxisTitleText: "",
            yAxisTitleColor: "#666666",
            yAxisTitleFontSize: 11,
            yAxisTitleFontFamily: "Segoe UI",
            yAxisTitleBold: false,
            yAxisTitleItalic: false,
            yAxisTitleUnderline: false,

            // Gridlines
            gridlinesShow: true,
            gridlinesColor: "#e6e6e6",
            gridlinesTransparency: 0,
            gridlinesStrokeWidth: 1,
            gridlinesScaleByWidth: false,
            gridlinesLineStyle: "dotted",

            // Columns
            seriesSelection: "All",
            fill: "",
            showBorder: false,
            borderColor: "#ffffff",
            borderWidth: 1,
            actualColor: "",
            actualTransparency: 0,
            benchmarkColor: "",
            benchmarkTransparency: 70,
            actualWidth: 65,
            innerPadding: 20,
            groupPadding: 20,

            // Data labels
            dataLabelsShow: false,
            dataLabelsShowValues: true,
            dataLabelsColor: "#333333",
            dataLabelsFontSize: 9,
            dataLabelsFontFamily: "Segoe UI",
            dataLabelsBold: false,
            dataLabelsItalic: false,
            dataLabelsUnderline: false,
            dataLabelsDisplayUnits: 0,
            dataLabelsValueDecimalPlaces: null,
            dataLabelsPosition: "Auto",

            // Benchmark labels
            showBenchmarkLabels: true,
            benchmarkLabelPrefix: "National Average: ",
            benchmarkLabelFontSize: 9,
            benchmarkLabelFontFamily: "Segoe UI",
            benchmarkLabelBold: false,
            benchmarkLabelItalic: false,
            benchmarkLabelUnderline: false,

            // Chart padding
            chartPaddingTop: 50,
            chartPaddingRight: 20,
            chartPaddingBottom: 75,
            chartPaddingLeft: 60
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
            this.latestGroupList = [];
            this.seriesOverrides = {};
            this.groupOverrides = {};
            this.drawLandingMessage(width, height);
            this.renderingFinished();
            return;
        }

        const rows = this.getCategoricalRows(this.currentDataView);

        if (!rows.length) {
            this.latestSeriesList = [];
            this.latestGroupList = [];
            this.seriesOverrides = {};
            this.seriesSelectors = {};
            this.groupOverrides = {};
            this.groupSelectors = {};
            this.hasSeries = false;
            this.drawLandingMessage(width, height);
            this.renderingFinished();
            return;
        }

        this.latestSeriesList = Array.from(new Set(rows.map(d => d.series))).sort();
        this.latestGroupList = Array.from(new Set(rows.map(d => d.group)));
        this.hasSeries = this.latestSeriesList.length > 1 || (this.latestSeriesList.length === 1 && this.latestSeriesList[0] !== "Unknown" && this.latestSeriesList[0] !== "");

        this.seriesOverrides = this.buildSeriesOverrides(this.currentDataView);
        this.groupOverrides = this.buildGroupOverrides(this.currentDataView);

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
                    this.makeDropdownSlice("legend_position", "Position", "legend", "position", this.settings.legendPosition)
                ]
            }, {
                displayName: "Text",
                uid: "legend_text_group",
                slices: [
                    this.makeColorSlice("legend_labelColor", "Color", "legend", "labelColor", this.settings.legendLabelColor),
                    this.makeFontControlSlice("legend_font", "legend",
                        this.settings.legendFontFamily, "fontFamily",
                        this.settings.legendFontSize, "fontSize",
                        this.settings.legendBold, "bold",
                        this.settings.legendItalic, "italic",
                        this.settings.legendUnderline, "underline")
                ]
            }, {
                displayName: "Title",
                uid: "legend_title_group",
                topLevelToggle: this.makeTopLevelToggle("legend", "showTitle", this.settings.legendShowTitle),
                slices: [
                    this.makeTextSlice("legend_titleText", "Title Text", "legend", "titleText", this.settings.legendTitleText)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "legend", propertyName: "show" },
                { objectName: "legend", propertyName: "position" },
                { objectName: "legend", propertyName: "showTitle" },
                { objectName: "legend", propertyName: "titleText" },
                { objectName: "legend", propertyName: "labelColor" },
                { objectName: "legend", propertyName: "fontSize" },
                { objectName: "legend", propertyName: "fontFamily" },
                { objectName: "legend", propertyName: "bold" },
                { objectName: "legend", propertyName: "italic" },
                { objectName: "legend", propertyName: "underline" }
            ]
        } as any);

        // 2. X-axis Card (categoryAxis)
        cards.push({
            displayName: "X-axis",
            uid: "categoryAxis_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("categoryAxis", "show", this.settings.xAxisShow),
            groups: [{
                displayName: "Values",
                uid: "categoryAxis_values_group",
                topLevelToggle: this.makeTopLevelToggle("categoryAxis", "showValues", this.settings.xAxisShowValues),
                slices: [
                    this.makeColorSlice("categoryAxis_labelColor", "Color", "categoryAxis", "labelColor", this.settings.xAxisLabelColor),
                    this.makeFontControlSlice("categoryAxis_font", "categoryAxis",
                        this.settings.xAxisFontFamily, "fontFamily",
                        this.settings.xAxisFontSize, "fontSize",
                        this.settings.xAxisBold, "bold",
                        this.settings.xAxisItalic, "italic",
                        this.settings.xAxisUnderline, "underline"),
                    this.makeNumericSlice("categoryAxis_labelOffset", "Label offset", "categoryAxis", "labelOffset", this.settings.xAxisLabelOffset),
                    this.makeNumericSlice("categoryAxis_padding", "Padding", "categoryAxis", "padding", this.settings.xAxisPadding)
                ]
            }, {
                displayName: "Title",
                uid: "categoryAxis_title_group",
                topLevelToggle: this.makeTopLevelToggle("categoryAxis", "showTitle", this.settings.xAxisShowTitle),
                slices: [
                    this.makeTextSlice("categoryAxis_titleText", "Title Text", "categoryAxis", "titleText", this.settings.xAxisTitleText),
                    this.makeColorSlice("categoryAxis_titleColor", "Title Color", "categoryAxis", "titleColor", this.settings.xAxisTitleColor),
                    this.makeFontControlSlice("categoryAxis_titleFont", "categoryAxis",
                        this.settings.xAxisTitleFontFamily, "titleFontFamily",
                        this.settings.xAxisTitleFontSize, "titleFontSize",
                        this.settings.xAxisTitleBold, "titleBold",
                        this.settings.xAxisTitleItalic, "titleItalic",
                        this.settings.xAxisTitleUnderline, "titleUnderline")
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "categoryAxis", propertyName: "show" },
                { objectName: "categoryAxis", propertyName: "showValues" },
                { objectName: "categoryAxis", propertyName: "labelOffset" },
                { objectName: "categoryAxis", propertyName: "padding" },
                { objectName: "categoryAxis", propertyName: "labelColor" },
                { objectName: "categoryAxis", propertyName: "fontSize" },
                { objectName: "categoryAxis", propertyName: "fontFamily" },
                { objectName: "categoryAxis", propertyName: "bold" },
                { objectName: "categoryAxis", propertyName: "italic" },
                { objectName: "categoryAxis", propertyName: "underline" },
                { objectName: "categoryAxis", propertyName: "showTitle" },
                { objectName: "categoryAxis", propertyName: "titleText" },
                { objectName: "categoryAxis", propertyName: "titleColor" },
                { objectName: "categoryAxis", propertyName: "titleFontSize" },
                { objectName: "categoryAxis", propertyName: "titleFontFamily" },
                { objectName: "categoryAxis", propertyName: "titleBold" },
                { objectName: "categoryAxis", propertyName: "titleItalic" },
                { objectName: "categoryAxis", propertyName: "titleUnderline" }
            ]
        } as any);

        // 3. Y-axis Card (valueAxis)
        cards.push({
            displayName: "Y-axis",
            uid: "valueAxis_card",
            analyticsPane: false,
            topLevelToggle: this.makeTopLevelToggle("valueAxis", "show", this.settings.yAxisShow),
            groups: [{
                displayName: "Values",
                uid: "valueAxis_values_group",
                topLevelToggle: this.makeTopLevelToggle("valueAxis", "showValues", this.settings.yAxisShowValues),
                slices: [
                    this.makeColorSlice("valueAxis_labelColor", "Color", "valueAxis", "labelColor", this.settings.yAxisLabelColor),
                    this.makeFontControlSlice("valueAxis_font", "valueAxis",
                        this.settings.yAxisFontFamily, "fontFamily",
                        this.settings.yAxisFontSize, "fontSize",
                        this.settings.yAxisBold, "bold",
                        this.settings.yAxisItalic, "italic",
                        this.settings.yAxisUnderline, "underline")
                ]
            }, {
                displayName: "Title",
                uid: "valueAxis_title_group",
                topLevelToggle: this.makeTopLevelToggle("valueAxis", "showTitle", this.settings.yAxisShowTitle),
                slices: [
                    this.makeTextSlice("valueAxis_titleText", "Title Text", "valueAxis", "titleText", this.settings.yAxisTitleText),
                    this.makeColorSlice("valueAxis_titleColor", "Title Color", "valueAxis", "titleColor", this.settings.yAxisTitleColor),
                    this.makeFontControlSlice("valueAxis_titleFont", "valueAxis",
                        this.settings.yAxisTitleFontFamily, "titleFontFamily",
                        this.settings.yAxisTitleFontSize, "titleFontSize",
                        this.settings.yAxisTitleBold, "titleBold",
                        this.settings.yAxisTitleItalic, "titleItalic",
                        this.settings.yAxisTitleUnderline, "titleUnderline")
                ]
            }, {
                displayName: "Options",
                uid: "valueAxis_options_group",
                slices: [
                    this.makeDropdownSlice("valueAxis_position", "Position", "valueAxis", "position", this.settings.yAxisPosition),
                    this.makeBoolSlice("valueAxis_showAxisLine", "Show axis line", "valueAxis", "showAxisLine", this.settings.yAxisShowAxisLine),
                    this.makeNumericSlice("valueAxis_start", "Start", "valueAxis", "start", this.settings.yAxisStart, undefined, { placeholder: "AUTO" }),
                    this.makeNumericSlice("valueAxis_end", "End", "valueAxis", "end", this.settings.yAxisEnd, undefined, { placeholder: "AUTO" }),
                    this.makeNumericSlice("valueAxis_padding", "Padding", "valueAxis", "padding", this.settings.yAxisPadding)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "valueAxis", propertyName: "show" },
                { objectName: "valueAxis", propertyName: "showValues" },
                { objectName: "valueAxis", propertyName: "position" },
                { objectName: "valueAxis", propertyName: "showAxisLine" },
                { objectName: "valueAxis", propertyName: "start" },
                { objectName: "valueAxis", propertyName: "end" },
                { objectName: "valueAxis", propertyName: "padding" },
                { objectName: "valueAxis", propertyName: "labelColor" },
                { objectName: "valueAxis", propertyName: "fontSize" },
                { objectName: "valueAxis", propertyName: "fontFamily" },
                { objectName: "valueAxis", propertyName: "bold" },
                { objectName: "valueAxis", propertyName: "italic" },
                { objectName: "valueAxis", propertyName: "underline" },
                { objectName: "valueAxis", propertyName: "showTitle" },
                { objectName: "valueAxis", propertyName: "titleText" },
                { objectName: "valueAxis", propertyName: "titleColor" },
                { objectName: "valueAxis", propertyName: "titleFontSize" },
                { objectName: "valueAxis", propertyName: "titleFontFamily" },
                { objectName: "valueAxis", propertyName: "titleBold" },
                { objectName: "valueAxis", propertyName: "titleItalic" },
                { objectName: "valueAxis", propertyName: "titleUnderline" }
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
                    this.makeNumericSlice("gridlines_transparency", "Transparency", "gridlines", "transparency", this.settings.gridlinesTransparency, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, maxValue: { type: powerbi.visuals.ValidatorType.Max, value: 100 } }),
                    this.makeNumericSlice("gridlines_strokeWidth", "Stroke width", "gridlines", "strokeWidth", this.settings.gridlinesStrokeWidth, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeBoolSlice("gridlines_scaleByWidth", "Scale by width", "gridlines", "scaleByWidth", this.settings.gridlinesScaleByWidth),
                    this.makeDropdownSlice("gridlines_lineStyle", "Line Style", "gridlines", "lineStyle", this.settings.gridlinesLineStyle)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "gridlines", propertyName: "show" },
                { objectName: "gridlines", propertyName: "color" },
                { objectName: "gridlines", propertyName: "transparency" },
                { objectName: "gridlines", propertyName: "strokeWidth" },
                { objectName: "gridlines", propertyName: "scaleByWidth" },
                { objectName: "gridlines", propertyName: "lineStyle" }
            ]
        } as any);

        // 5. Columns Card
        const seriesItems: powerbi.IEnumMember[] = [{ value: "All", displayName: "All" }];
        
        const targetList = this.hasSeries ? this.latestSeriesList : this.latestGroupList;

        targetList.forEach(item => {
            seriesItems.push({ value: item, displayName: item });
        });

        // Current selected series or "All"
        let selectedSeries = this.settings.seriesSelection;
        if (!seriesItems.find(i => i.value === selectedSeries)) {
            selectedSeries = "All";
        }

        const isAll = selectedSeries === "All";
        const style = this.hasSeries ? this.getSeriesStyle(selectedSeries) : this.getGroupStyle(selectedSeries);

        const targetColorActual = isAll ? this.settings.actualColor : style.actualColor;
        const targetTransparencyActual = isAll ? this.settings.actualTransparency : this.opacityToTransparency(style.actualOpacity);
        const targetColorBenchmark = isAll ? this.settings.benchmarkColor : style.benchmarkColor;
        const targetTransparencyBenchmark = isAll ? this.settings.benchmarkTransparency : this.opacityToTransparency(style.benchmarkOpacity);
        
        const targetSelector = isAll ? undefined : (this.hasSeries ? this.seriesSelectors[selectedSeries] : this.groupSelectors[selectedSeries]);
        const dropdownDisplayName = this.hasSeries ? "Series" : "Base Category";

        const columnsSlices: any[] = [
            this.makeDropdownSlice("columns_seriesSelection", dropdownDisplayName, "columns", "seriesSelection", selectedSeries, seriesItems)
        ];

        const colourSlices: any[] = [
            this.makeColorSlice(
                isAll ? "columns_fill" : `columns_fill_${selectedSeries}`,
                "Fill",
                "columns",
                "fill",
                this.settings.fill,
                targetSelector
            ),
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

        const borderSlices: any[] = [
            this.makeBoolSlice("columns_showBorder", "Show border", "columns", "showBorder", this.settings.showBorder),
            this.makeColorSlice("columns_borderColor", "Border colour", "columns", "borderColor", this.settings.borderColor),
            this.makeNumericSlice("columns_borderWidth", "Border width", "columns", "borderWidth", this.settings.borderWidth, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } })
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
                displayName: "Apply settings to series",
                uid: "columns_options_group",
                slices: columnsSlices
            }, {
                displayName: "Colour",
                uid: "columns_colour_group",
                slices: colourSlices
            }, {
                displayName: "Border",
                uid: "columns_border_group",
                slices: borderSlices
            }, {
                displayName: "Layout",
                uid: "columns_layout_group",
                slices: layoutSlices
            }],
            revertToDefaultDescriptors: [
                { objectName: "columns", propertyName: "seriesSelection" },
                { objectName: "columns", propertyName: "fill" },
                { objectName: "columns", propertyName: "showBorder" },
                { objectName: "columns", propertyName: "borderColor" },
                { objectName: "columns", propertyName: "borderWidth" },
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
                displayName: "Values",
                uid: "dataLabels_values_group",
                topLevelToggle: this.makeTopLevelToggle("dataLabels", "showValues", this.settings.dataLabelsShowValues),
                slices: [
                    this.makeColorSlice("dataLabels_color", "Color", "dataLabels", "color", this.settings.dataLabelsColor),
                    this.makeFontControlSlice("dataLabels_font", "dataLabels",
                        this.settings.dataLabelsFontFamily, "fontFamily",
                        this.settings.dataLabelsFontSize, "fontSize",
                        this.settings.dataLabelsBold, "bold",
                        this.settings.dataLabelsItalic, "italic",
                        this.settings.dataLabelsUnderline, "underline")
                ]
            }, {
                displayName: "Options",
                uid: "dataLabels_options_group",
                slices: [
                    this.makeNumericSlice("dataLabels_displayUnits", "Display Units", "dataLabels", "displayUnits", this.settings.dataLabelsDisplayUnits),
                    this.makeNumericSlice("dataLabels_valueDecimalPlaces", "Value Decimal Places", "dataLabels", "valueDecimalPlaces", this.settings.dataLabelsValueDecimalPlaces, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 }, placeholder: "AUTO" }),
                    this.makeDropdownSlice("dataLabels_position", "Position", "dataLabels", "position", this.settings.dataLabelsPosition)
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "dataLabels", propertyName: "show" },
                { objectName: "dataLabels", propertyName: "showValues" },
                { objectName: "dataLabels", propertyName: "color" },
                { objectName: "dataLabels", propertyName: "fontSize" },
                { objectName: "dataLabels", propertyName: "fontFamily" },
                { objectName: "dataLabels", propertyName: "bold" },
                { objectName: "dataLabels", propertyName: "italic" },
                { objectName: "dataLabels", propertyName: "underline" },
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
                    this.makeFontControlSlice("benchmarkLabels_font", "benchmarkLabels",
                        this.settings.benchmarkLabelFontFamily, "fontFamily",
                        this.settings.benchmarkLabelFontSize, "fontSize",
                        this.settings.benchmarkLabelBold, "bold",
                        this.settings.benchmarkLabelItalic, "italic",
                        this.settings.benchmarkLabelUnderline, "underline")
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "benchmarkLabels", propertyName: "show" },
                { objectName: "benchmarkLabels", propertyName: "labelPrefix" },
                { objectName: "benchmarkLabels", propertyName: "fontSize" },
                { objectName: "benchmarkLabels", propertyName: "fontFamily" },
                { objectName: "benchmarkLabels", propertyName: "bold" },
                { objectName: "benchmarkLabels", propertyName: "italic" },
                { objectName: "benchmarkLabels", propertyName: "underline" }
            ]
        } as any);

        // 8. Chart Padding Card
        cards.push({
            displayName: "Chart padding",
            uid: "chartPadding_card",
            analyticsPane: false,
            groups: [{
                displayName: "Padding",
                uid: "chartPadding_group",
                slices: [
                    this.makeNumericSlice("chartPadding_top", "Top", "chartPadding", "top", this.settings.chartPaddingTop, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeNumericSlice("chartPadding_right", "Right", "chartPadding", "right", this.settings.chartPaddingRight, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeNumericSlice("chartPadding_bottom", "Bottom", "chartPadding", "bottom", this.settings.chartPaddingBottom, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } }),
                    this.makeNumericSlice("chartPadding_left", "Left", "chartPadding", "left", this.settings.chartPaddingLeft, undefined, { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 0 } })
                ]
            }],
            revertToDefaultDescriptors: [
                { objectName: "chartPadding", propertyName: "top" },
                { objectName: "chartPadding", propertyName: "right" },
                { objectName: "chartPadding", propertyName: "bottom" },
                { objectName: "chartPadding", propertyName: "left" }
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

                if (col.source.roles?.["actual"]) {
                    row.actual = rawValue;
                    row.actualHighlight = rawHighlight;
                } else if (col.source.roles?.["benchmark"]) {
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

    private buildGroupOverrides(dataView: DataView): Record<string, Partial<SeriesStyle>> {
        const overrides: Record<string, Partial<SeriesStyle>> = {};
        this.groupSelectors = {};
        
        const categorical = dataView.categorical;
        if (!categorical || !categorical.categories || categorical.categories.length === 0) return overrides;

        const categoryColumn = categorical.categories[0];
        const values = categoryColumn.values || [];
        const objects = categoryColumn.objects || [];

        for (let r = 0; r < values.length; r++) {
            const groupName = String(values[r] ?? "");
            if (!this.groupSelectors[groupName]) {
                const selectionId = this.host.createSelectionIdBuilder()
                    .withCategory(categoryColumn, r)
                    .createSelectionId();
                this.groupSelectors[groupName] = selectionId.getSelector() as powerbi.data.Selector;
                overrides[groupName] = {};
            }
            if (objects[r]) {
                this.applyStyleFromObjects(overrides[groupName], objects[r]);
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
        const margin = {
            top: Math.max(0, this.settings.chartPaddingTop),
            right: Math.max(0, this.settings.chartPaddingRight),
            bottom: Math.max(0, this.settings.chartPaddingBottom),
            left: Math.max(0, this.settings.chartPaddingLeft)
        };

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
            const yAxisGrid = d3.axisLeft(y)
                .tickSize(this.settings.gridlinesScaleByWidth ? -plotWidth : -this.settings.gridlinesStrokeWidth)
                .tickFormat("" as any);
            const gridGroup = this.root.append("g")
                .attr("class", "gridlines")
                .call(yAxisGrid);

            gridGroup.selectAll("line")
                .style("stroke", this.settings.gridlinesColor)
                .style("stroke-width", `${this.settings.gridlinesStrokeWidth}px`)
                .style("opacity", this.transparencyToOpacity(this.settings.gridlinesTransparency))
                .style("stroke-dasharray", this.settings.gridlinesLineStyle === "dashed" ? "5,5" : this.settings.gridlinesLineStyle === "dotted" ? "2,2" : "none");
            gridGroup.select(".domain").remove();
        }

        // X Axis
        const xAxisGen = d3.axisBottom(x0)
            .tickPadding(this.settings.xAxisLabelOffset + this.settings.xAxisPadding);
            
        const xAxis = this.root.append("g")
            .attr("transform", `translate(0,${plotHeight})`)
            .call(xAxisGen);

        if (!this.settings.xAxisShow || !this.settings.xAxisShowValues) {
            xAxis.selectAll("text").style("display", "none");
        } else {
            const maxWidth = x0.bandwidth();
            xAxis.selectAll("text")
                .style("fill", this.settings.xAxisLabelColor)
                .style("font-size", `${this.settings.xAxisFontSize}px`)
                .style("font-family", this.settings.xAxisFontFamily)
                .style("font-weight", this.settings.xAxisBold ? "bold" : "normal")
                .style("font-style", this.settings.xAxisItalic ? "italic" : "normal")
                .style("text-decoration", this.settings.xAxisUnderline ? "underline" : "none")
                .call(this.wrapText, maxWidth);
        }

        if (!this.settings.xAxisShow) {
            xAxis.select(".domain").style("display", "none");
            xAxis.selectAll("line").style("display", "none");
        }

        if (this.settings.xAxisShow && this.settings.xAxisShowTitle) {
            this.root.append("text")
                .attr("x", plotWidth / 2)
                .attr("y", plotHeight + margin.bottom - 10)
                .attr("text-anchor", "middle")
                .style("fill", this.settings.xAxisTitleColor)
                .style("font-size", `${this.settings.xAxisTitleFontSize}px`)
                .style("font-family", this.settings.xAxisTitleFontFamily)
                .style("font-weight", this.settings.xAxisTitleBold ? "bold" : "normal")
                .style("font-style", this.settings.xAxisTitleItalic ? "italic" : "normal")
                .style("text-decoration", this.settings.xAxisTitleUnderline ? "underline" : "none")
                .text(this.settings.xAxisTitleText || "Base Category");
        }

        // Y Axis
        const yAxisGen = this.settings.yAxisPosition === "Right" ? d3.axisRight(y) : d3.axisLeft(y);
        yAxisGen.tickPadding(this.settings.yAxisPadding);

        const yAxis = this.root.append("g")
            .attr("transform", this.settings.yAxisPosition === "Right" ? `translate(${plotWidth},0)` : "translate(0,0)")
            .call(yAxisGen);

        if (!this.settings.yAxisShowAxisLine || !this.settings.yAxisShow) {
            yAxis.select(".domain").remove();
        }

        if (!this.settings.yAxisShow || !this.settings.yAxisShowValues) {
            yAxis.selectAll("text").style("display", "none");
        } else {
            yAxis.selectAll("text")
                .style("fill", this.settings.yAxisLabelColor)
                .style("font-size", `${this.settings.yAxisFontSize}px`)
                .style("font-family", this.settings.yAxisFontFamily)
                .style("font-weight", this.settings.yAxisBold ? "bold" : "normal")
                .style("font-style", this.settings.yAxisItalic ? "italic" : "normal")
                .style("text-decoration", this.settings.yAxisUnderline ? "underline" : "none");
        }

        if (this.settings.yAxisShow && this.settings.yAxisShowTitle) {
            this.root.append("text")
                .attr("transform", "rotate(-90)")
                .attr("x", -plotHeight / 2)
                .attr("y", this.settings.yAxisPosition === "Right" ? plotWidth + 40 : -margin.left + 15)
                .attr("text-anchor", "middle")
                .style("fill", this.settings.yAxisTitleColor)
                .style("font-size", `${this.settings.yAxisTitleFontSize}px`)
                .style("font-family", this.settings.yAxisTitleFontFamily)
                .style("font-weight", this.settings.yAxisTitleBold ? "bold" : "normal")
                .style("font-style", this.settings.yAxisTitleItalic ? "italic" : "normal")
                .style("text-decoration", this.settings.yAxisTitleUnderline ? "underline" : "none")
                .text(this.settings.yAxisTitleText || "Value");
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
                .attr("fill", d => isHighContrast ? background : (this.hasSeries ? this.getSeriesStyle(d.series).benchmarkColor : this.getGroupStyle(d.group).benchmarkColor))
                .attr("stroke", isHighContrast ? foreground : (this.settings.showBorder ? this.settings.borderColor : "none"))
                .attr("stroke-width", isHighContrast ? 1.5 : (this.settings.showBorder ? this.settings.borderWidth : 0))
                .attr("opacity", d => isHighContrast ? 1 : (this.hasSeries ? this.getSeriesStyle(d.series).benchmarkOpacity : this.getGroupStyle(d.group).benchmarkOpacity))
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
                .attr("fill", d => isHighContrast ? foreground : (this.hasSeries ? this.getSeriesStyle(d.series).actualColor : this.getGroupStyle(d.group).actualColor))
                .attr("stroke", isHighContrast ? foreground : (this.settings.showBorder ? this.settings.borderColor : "none"))
                .attr("stroke-width", isHighContrast ? 1 : (this.settings.showBorder ? this.settings.borderWidth : 0))
                .attr("opacity", d => isHighContrast ? 1 : (this.hasSeries ? this.getSeriesStyle(d.series).actualOpacity : this.getGroupStyle(d.group).actualOpacity))
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
                    .style("font-weight", this.settings.benchmarkLabelBold ? "bold" : "normal")
                    .style("font-style", this.settings.benchmarkLabelItalic ? "italic" : "normal")
                    .style("text-decoration", this.settings.benchmarkLabelUnderline ? "underline" : "none")
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
                    .style("font-weight", this.settings.dataLabelsBold ? "bold" : "normal")
                    .style("font-style", this.settings.dataLabelsItalic ? "italic" : "normal")
                    .style("text-decoration", this.settings.dataLabelsUnderline ? "underline" : "none")
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
                const validSeries = groupData.filter(d => d.series !== "Unknown");
                if (validSeries.length > 0) {
                    g.selectAll<SVGTextElement, ChartRow>(".seriesLabel")
                        .data(validSeries)
                        .enter()
                        .append("text")
                        .attr("class", "seriesLabel")
                        .attr("x", d => (x1(d.series) ?? 0) + x1.bandwidth() / 2)
                        .attr("y", plotHeight + 35)
                        .attr("text-anchor", "middle")
                        .style("font-size", `${Math.max(8, labelFontSize - 1)}px`)
                        .style("fill", isHighContrast ? foreground : null)
                        .text(d => d.series);
                }
            }
        });

        // Draw Legend
        if (this.settings.legendShow && this.hasSeries && seriesList.length) {
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

    private getGroupStyle(groupName: string): SeriesStyle {
        const globalStyle = this.getGlobalStyle();
        const baseColor = this.host.colorPalette.getColor("Unknown").value;
        const override = this.groupOverrides[groupName] || {};

        return {
            actualColor: override.actualColor || globalStyle.actualColor || baseColor,
            actualOpacity: typeof override.actualOpacity === "number" ? override.actualOpacity : globalStyle.actualOpacity,
            benchmarkColor: override.benchmarkColor || globalStyle.benchmarkColor || baseColor,
            benchmarkOpacity: typeof override.benchmarkOpacity === "number" ? override.benchmarkOpacity : globalStyle.benchmarkOpacity
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

    private makeFontControlSlice(
        uid: string,
        objectName: string,
        fontFamily: string, fontFamilyProp: string,
        fontSize: number, fontSizeProp: string,
        bold: boolean, boldProp: string,
        italic: boolean, italicProp: string,
        underline: boolean, underlineProp: string,
        selector?: powerbi.data.Selector
    ): any {
        const desc = (prop: string) => ({
            descriptor: { objectName, propertyName: prop, selector }
        });
        return {
            uid,
            displayName: "",
            control: {
                type: powerbi.visuals.FormattingComponent.FontControl,
                properties: {
                    fontFamily: { ...desc(fontFamilyProp), value: fontFamily },
                    fontSize: { ...desc(fontSizeProp), value: fontSize, options: { minValue: { type: powerbi.visuals.ValidatorType.Min, value: 8 } } },
                    bold: { ...desc(boldProp), value: bold },
                    italic: { ...desc(italicProp), value: italic },
                    underline: { ...desc(underlineProp), value: underline }
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
                const seriesOpacity = this.hasSeries ? this.getSeriesStyle(d.series).benchmarkOpacity : this.getGroupStyle(d.group).benchmarkOpacity;

                if (!hasSelection) {
                    return seriesOpacity;
                }

                const rowKey = this.selectionIdToKey(d.selectionId);
                return selectedKeys.includes(rowKey) ? seriesOpacity : 0.25;
            });

        this.root
            .selectAll<SVGRectElement, ChartRow>(".actual")
            .attr("opacity", d => {
                const seriesOpacity = this.hasSeries ? this.getSeriesStyle(d.series).actualOpacity : this.getGroupStyle(d.group).actualOpacity;

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

    private wrapText(textGroup: d3.Selection<d3.BaseType, any, SVGGElement, any>, width: number): void {
        textGroup.each(function() {
            const text = d3.select(this);
            const words = text.text().split(/\s+/).reverse();
            let word: string | undefined;
            let line: string[] = [];
            let lineNumber = 0;
            const lineHeight = 1.1; // ems
            const y = text.attr("y");
            const dy = parseFloat(text.attr("dy")) || 0;
            let targetTspan = text.text(null).append("tspan").attr("x", 0).attr("y", y).attr("dy", dy + "em");

            while (word = words.pop()) {
                line.push(word);
                targetTspan.text(line.join(" "));
                if (targetTspan.node()!.getComputedTextLength() > width && line.length > 1) {
                    line.pop();
                    targetTspan.text(line.join(" "));
                    line = [word];
                    targetTspan = text.append("tspan").attr("x", 0).attr("y", y).attr("dy", ++lineNumber * lineHeight + dy + "em").text(word);
                }
            }
        });
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
