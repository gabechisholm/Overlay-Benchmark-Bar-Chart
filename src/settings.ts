import powerbi from "powerbi-visuals-api";

export interface VisualSettings {
    // Legend
    legendShow: boolean;
    legendPosition: string;
    legendShowTitle: boolean;
    legendTitleText: string;
    legendLabelColor: string;
    legendFontSize: number;
    legendFontFamily: string;
    legendBold: boolean;
    legendItalic: boolean;
    legendUnderline: boolean;

    // X-axis
    xAxisShow: boolean;
    xAxisLabelOffset: number;
    xAxisPadding: number;
    xAxisShowValues: boolean;
    xAxisLabelColor: string;
    xAxisFontSize: number;
    xAxisFontFamily: string;
    xAxisBold: boolean;
    xAxisItalic: boolean;
    xAxisUnderline: boolean;
    xAxisShowTitle: boolean;
    xAxisTitleText: string;
    xAxisTitleColor: string;
    xAxisTitleFontSize: number;
    xAxisTitleFontFamily: string;
    xAxisTitleBold: boolean;
    xAxisTitleItalic: boolean;
    xAxisTitleUnderline: boolean;
    xAxisLabelAngle: number;

    // Y-axis
    yAxisShow: boolean;
    yAxisShowAxisLine: boolean;
    yAxisPosition: string;
    yAxisStart: number | null;
    yAxisEnd: number | null;
    yAxisPadding: number;
    yAxisShowValues: boolean;
    yAxisLabelColor: string;
    yAxisFontSize: number;
    yAxisFontFamily: string;
    yAxisBold: boolean;
    yAxisItalic: boolean;
    yAxisUnderline: boolean;
    yAxisShowTitle: boolean;
    yAxisTitleText: string;
    yAxisTitleColor: string;
    yAxisTitleFontSize: number;
    yAxisTitleFontFamily: string;
    yAxisTitleBold: boolean;
    yAxisTitleItalic: boolean;
    yAxisTitleUnderline: boolean;

    // Gridlines
    gridlinesShow: boolean;
    gridlinesColor: string;
    gridlinesTransparency: number;
    gridlinesStrokeWidth: number;
    gridlinesScaleByWidth: boolean;
    gridlinesLineStyle: string;

    // Columns
    seriesSelection: string;
    fill: string;
    showBorder: boolean;
    borderColor: string;
    borderWidth: number;
    actualColor: string;
    actualTransparency: number;
    benchmarkColor: string;
    benchmarkTransparency: number;
    actualWidth: number;
    benchmarkWidth: number;
    innerPadding: number;
    groupPadding: number;

    // Data labels
    dataLabelsShow: boolean;
    dataLabelsShowValues: boolean;
    dataLabelsColor: string;
    dataLabelsFontSize: number;
    dataLabelsFontFamily: string;
    dataLabelsBold: boolean;
    dataLabelsItalic: boolean;
    dataLabelsUnderline: boolean;
    dataLabelsDisplayUnits: number;
    dataLabelsValueDecimalPlaces: number | null;
    dataLabelsPosition: string;

    // Benchmark labels
    showBenchmarkLabels: boolean;
    benchmarkLabelPrefix: string;
    benchmarkLabelDecimalPlaces: number | null;
    benchmarkLabelColor: string;
    benchmarkLabelFontSize: number;
    benchmarkLabelFontFamily: string;
    benchmarkLabelBold: boolean;
    benchmarkLabelItalic: boolean;
    benchmarkLabelUnderline: boolean;

    // Chart padding
    chartPaddingTop: number;
    chartPaddingRight: number;
    chartPaddingBottom: number;
    chartPaddingLeft: number;
}

export class SettingsParser {
    public static parse(dataView?: powerbi.DataView): VisualSettings {
        const objects = dataView?.metadata?.objects;

        return {
            // Legend
            legendShow: this.getBool(objects, "legend", "show", true),
            legendPosition: this.getString(objects, "legend", "position", "TopLeft"),
            legendShowTitle: this.getBool(objects, "legend", "showTitle", true),
            legendTitleText: this.getString(objects, "legend", "titleText", "Category"),
            legendLabelColor: this.getFillColor(objects, "legend", "labelColor", "#666666"),
            legendFontSize: Math.max(8, this.getNumber(objects, "legend", "fontSize", 9)),
            legendFontFamily: this.getString(objects, "legend", "fontFamily", "Segoe UI"),
            legendBold: this.getBool(objects, "legend", "bold", false),
            legendItalic: this.getBool(objects, "legend", "italic", false),
            legendUnderline: this.getBool(objects, "legend", "underline", false),

            // X-axis
            xAxisShow: this.getBool(objects, "categoryAxis", "show", true),
            xAxisLabelOffset: this.getNumber(objects, "categoryAxis", "labelOffset", 3),
            xAxisPadding: Math.max(0, this.getNumber(objects, "categoryAxis", "padding", 0)),
            xAxisShowValues: this.getBool(objects, "categoryAxis", "showValues", true),
            xAxisLabelColor: this.getFillColor(objects, "categoryAxis", "labelColor", "#666666"),
            xAxisFontSize: Math.max(8, this.getNumber(objects, "categoryAxis", "fontSize", 9)),
            xAxisFontFamily: this.getString(objects, "categoryAxis", "fontFamily", "Segoe UI"),
            xAxisBold: this.getBool(objects, "categoryAxis", "bold", false),
            xAxisItalic: this.getBool(objects, "categoryAxis", "italic", false),
            xAxisUnderline: this.getBool(objects, "categoryAxis", "underline", false),
            xAxisShowTitle: this.getBool(objects, "categoryAxis", "showTitle", true),
            xAxisTitleText: this.getString(objects, "categoryAxis", "titleText", ""),
            xAxisTitleColor: this.getFillColor(objects, "categoryAxis", "titleColor", "#666666"),
            xAxisTitleFontSize: Math.max(8, this.getNumber(objects, "categoryAxis", "titleFontSize", 11)),
            xAxisTitleFontFamily: this.getString(objects, "categoryAxis", "titleFontFamily", "Segoe UI"),
            xAxisTitleBold: this.getBool(objects, "categoryAxis", "titleBold", false),
            xAxisTitleItalic: this.getBool(objects, "categoryAxis", "titleItalic", false),
            xAxisTitleUnderline: this.getBool(objects, "categoryAxis", "titleUnderline", false),
            xAxisLabelAngle: this.getNumber(objects, "categoryAxis", "labelAngle", -45),

            // Y-axis
            yAxisShow: this.getBool(objects, "valueAxis", "show", true),
            yAxisShowAxisLine: this.getBool(objects, "valueAxis", "showAxisLine", false),
            yAxisPosition: this.getString(objects, "valueAxis", "position", "Left"),
            yAxisStart: this.getNumberOptional(objects, "valueAxis", "start"),
            yAxisEnd: this.getNumberOptional(objects, "valueAxis", "end"),
            yAxisPadding: Math.max(0, this.getNumber(objects, "valueAxis", "padding", 0)),
            yAxisShowValues: this.getBool(objects, "valueAxis", "showValues", true),
            yAxisLabelColor: this.getFillColor(objects, "valueAxis", "labelColor", "#666666"),
            yAxisFontSize: Math.max(8, this.getNumber(objects, "valueAxis", "fontSize", 9)),
            yAxisFontFamily: this.getString(objects, "valueAxis", "fontFamily", "Segoe UI"),
            yAxisBold: this.getBool(objects, "valueAxis", "bold", false),
            yAxisItalic: this.getBool(objects, "valueAxis", "italic", false),
            yAxisUnderline: this.getBool(objects, "valueAxis", "underline", false),
            yAxisShowTitle: this.getBool(objects, "valueAxis", "showTitle", true),
            yAxisTitleText: this.getString(objects, "valueAxis", "titleText", ""),
            yAxisTitleColor: this.getFillColor(objects, "valueAxis", "titleColor", "#666666"),
            yAxisTitleFontSize: Math.max(8, this.getNumber(objects, "valueAxis", "titleFontSize", 11)),
            yAxisTitleFontFamily: this.getString(objects, "valueAxis", "titleFontFamily", "Segoe UI"),
            yAxisTitleBold: this.getBool(objects, "valueAxis", "titleBold", false),
            yAxisTitleItalic: this.getBool(objects, "valueAxis", "titleItalic", false),
            yAxisTitleUnderline: this.getBool(objects, "valueAxis", "titleUnderline", false),

            // Gridlines
            gridlinesShow: this.getBool(objects, "gridlines", "show", true),
            gridlinesColor: this.getFillColor(objects, "gridlines", "color", "#e6e6e6"),
            gridlinesTransparency: Math.max(0, Math.min(100, this.getNumber(objects, "gridlines", "transparency", 0))),
            gridlinesStrokeWidth: this.getNumber(objects, "gridlines", "strokeWidth", 1),
            gridlinesScaleByWidth: this.getBool(objects, "gridlines", "scaleByWidth", false),
            gridlinesLineStyle: this.getString(objects, "gridlines", "lineStyle", "dotted"),

            // Columns
            seriesSelection: this.getString(objects, "columns", "seriesSelection", "All"),
            fill: this.getFillColor(objects, "columns", "fill", ""),
            showBorder: this.getBool(objects, "columns", "showBorder", false),
            borderColor: this.getFillColor(objects, "columns", "borderColor", "#ffffff"),
            borderWidth: Math.max(0, this.getNumber(objects, "columns", "borderWidth", 1)),
            actualColor: this.getFillColor(objects, "columns", "actualColor", ""),
            actualTransparency: Math.max(0, Math.min(100, this.getNumber(objects, "columns", "actualTransparency", 0))),
            benchmarkColor: this.getFillColor(objects, "columns", "benchmarkColor", ""),
            benchmarkTransparency: Math.max(0, Math.min(100, this.getNumber(objects, "columns", "benchmarkTransparency", 70))),
            actualWidth: Math.max(0, Math.min(100, SettingsParser.getNumber(objects, "columns", "actualWidth", 40))),
            benchmarkWidth: Math.max(0, Math.min(100, SettingsParser.getNumber(objects, "columns", "benchmarkWidth", 70))),
            innerPadding: Math.max(0, this.getNumber(objects, "columns", "innerPadding", 20)),
            groupPadding: Math.max(0, this.getNumber(objects, "columns", "groupPadding", 20)),

            // Data labels
            dataLabelsShow: this.getBool(objects, "dataLabels", "show", false),
            dataLabelsShowValues: this.getBool(objects, "dataLabels", "showValues", true),
            dataLabelsColor: this.getFillColor(objects, "dataLabels", "color", "#333333"),
            dataLabelsFontSize: Math.max(8, this.getNumber(objects, "dataLabels", "fontSize", 9)),
            dataLabelsFontFamily: this.getString(objects, "dataLabels", "fontFamily", "Segoe UI"),
            dataLabelsBold: this.getBool(objects, "dataLabels", "bold", false),
            dataLabelsItalic: this.getBool(objects, "dataLabels", "italic", false),
            dataLabelsUnderline: this.getBool(objects, "dataLabels", "underline", false),
            dataLabelsDisplayUnits: this.getNumber(objects, "dataLabels", "displayUnits", 0),
            dataLabelsValueDecimalPlaces: this.getNumberOptional(objects, "dataLabels", "valueDecimalPlaces"),
            dataLabelsPosition: this.getString(objects, "dataLabels", "position", "Auto"),

            // Benchmark labels
            showBenchmarkLabels: SettingsParser.getBool(objects, "benchmarkLabels", "show", false),
            benchmarkLabelPrefix: this.getString(objects, "benchmarkLabels", "labelPrefix", "National Average: "),
            benchmarkLabelDecimalPlaces: this.getNumberOptional(objects, "benchmarkLabels", "valueDecimalPlaces"),
            benchmarkLabelColor: this.getFillColor(objects, "benchmarkLabels", "color", ""),
            benchmarkLabelFontSize: Math.max(8, this.getNumber(objects, "benchmarkLabels", "fontSize", 9)),
            benchmarkLabelFontFamily: this.getString(objects, "benchmarkLabels", "fontFamily", "Segoe UI"),
            benchmarkLabelBold: this.getBool(objects, "benchmarkLabels", "bold", false),
            benchmarkLabelItalic: this.getBool(objects, "benchmarkLabels", "italic", false),
            benchmarkLabelUnderline: this.getBool(objects, "benchmarkLabels", "underline", false),

            // Chart padding
            chartPaddingTop: this.getNumber(objects, "chartPadding", "top", 50),
            chartPaddingRight: this.getNumber(objects, "chartPadding", "right", 20),
            chartPaddingBottom: this.getNumber(objects, "chartPadding", "bottom", 75),
            chartPaddingLeft: this.getNumber(objects, "chartPadding", "left", 60)
        };
    }

    private static getFillColor(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string,
        fallback: string
    ): string {
        const value = (objects as any)?.[objectName]?.[propertyName]?.solid?.color;
        return typeof value === "string" ? value : fallback;
    }

    private static getNumber(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string,
        fallback: number
    ): number {
        const value = (objects as any)?.[objectName]?.[propertyName];
        return typeof value === "number" ? value : fallback;
    }

    private static getNumberOptional(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string
    ): number | null {
        const value = (objects as any)?.[objectName]?.[propertyName];
        return typeof value === "number" ? value : null;
    }

    private static getBool(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string,
        fallback: boolean
    ): boolean {
        const value = (objects as any)?.[objectName]?.[propertyName];
        return typeof value === "boolean" ? value : fallback;
    }

    private static getString(
        objects: powerbi.DataViewObjects | undefined,
        objectName: string,
        propertyName: string,
        fallback: string
    ): string {
        const value = (objects as any)?.[objectName]?.[propertyName];
        return typeof value === "string" ? value : fallback;
    }
}
