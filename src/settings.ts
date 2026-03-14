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

    // X-axis
    xAxisShow: boolean;
    xAxisLabelOffset: number;
    xAxisLabelColor: string;
    xAxisFontSize: number;
    xAxisFontFamily: string;
    xAxisShowTitle: boolean;
    xAxisTitleText: string;
    xAxisTitleColor: string;
    xAxisTitleFontSize: number;
    xAxisTitleFontFamily: string;

    // Y-axis
    yAxisShow: boolean;
    yAxisShowAxisLine: boolean;
    yAxisPosition: string;
    yAxisStart: number | null;
    yAxisEnd: number | null;
    yAxisLabelColor: string;
    yAxisFontSize: number;
    yAxisFontFamily: string;
    yAxisShowTitle: boolean;
    yAxisTitleText: string;
    yAxisTitleColor: string;
    yAxisTitleFontSize: number;
    yAxisTitleFontFamily: string;

    // Gridlines
    gridlinesShow: boolean;
    gridlinesColor: string;
    gridlinesStrokeWidth: number;
    gridlinesLineStyle: string;

    // Columns
    seriesSelection: string;
    actualColor: string;
    actualTransparency: number;
    benchmarkColor: string;
    benchmarkTransparency: number;
    actualWidth: number;
    innerPadding: number;
    groupPadding: number;

    // Data labels
    dataLabelsShow: boolean;
    dataLabelsColor: string;
    dataLabelsFontSize: number;
    dataLabelsFontFamily: string;
    dataLabelsDisplayUnits: number;
    dataLabelsValueDecimalPlaces: number | null;
    dataLabelsPosition: string;

    // Benchmark labels
    showBenchmarkLabels: boolean;
    benchmarkLabelPrefix: string;
    benchmarkLabelFontSize: number;
    benchmarkLabelFontFamily: string;
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

            // X-axis
            xAxisShow: this.getBool(objects, "categoryAxis", "show", true),
            xAxisLabelOffset: this.getNumber(objects, "categoryAxis", "labelOffset", 3),
            xAxisLabelColor: this.getFillColor(objects, "categoryAxis", "labelColor", "#666666"),
            xAxisFontSize: Math.max(8, this.getNumber(objects, "categoryAxis", "fontSize", 9)),
            xAxisFontFamily: this.getString(objects, "categoryAxis", "fontFamily", "Segoe UI"),
            xAxisShowTitle: this.getBool(objects, "categoryAxis", "showTitle", true),
            xAxisTitleText: this.getString(objects, "categoryAxis", "titleText", ""),
            xAxisTitleColor: this.getFillColor(objects, "categoryAxis", "titleColor", "#666666"),
            xAxisTitleFontSize: Math.max(8, this.getNumber(objects, "categoryAxis", "titleFontSize", 11)),
            xAxisTitleFontFamily: this.getString(objects, "categoryAxis", "titleFontFamily", "Segoe UI"),

            // Y-axis
            yAxisShow: this.getBool(objects, "valueAxis", "show", true),
            yAxisShowAxisLine: this.getBool(objects, "valueAxis", "showAxisLine", false),
            yAxisPosition: this.getString(objects, "valueAxis", "position", "Left"),
            yAxisStart: this.getNumberOptional(objects, "valueAxis", "start"),
            yAxisEnd: this.getNumberOptional(objects, "valueAxis", "end"),
            yAxisLabelColor: this.getFillColor(objects, "valueAxis", "labelColor", "#666666"),
            yAxisFontSize: Math.max(8, this.getNumber(objects, "valueAxis", "fontSize", 9)),
            yAxisFontFamily: this.getString(objects, "valueAxis", "fontFamily", "Segoe UI"),
            yAxisShowTitle: this.getBool(objects, "valueAxis", "showTitle", true),
            yAxisTitleText: this.getString(objects, "valueAxis", "titleText", ""),
            yAxisTitleColor: this.getFillColor(objects, "valueAxis", "titleColor", "#666666"),
            yAxisTitleFontSize: Math.max(8, this.getNumber(objects, "valueAxis", "titleFontSize", 11)),
            yAxisTitleFontFamily: this.getString(objects, "valueAxis", "titleFontFamily", "Segoe UI"),

            // Gridlines
            gridlinesShow: this.getBool(objects, "gridlines", "show", true),
            gridlinesColor: this.getFillColor(objects, "gridlines", "color", "#e6e6e6"),
            gridlinesStrokeWidth: this.getNumber(objects, "gridlines", "strokeWidth", 1),
            gridlinesLineStyle: this.getString(objects, "gridlines", "lineStyle", "dotted"),

            // Columns
            seriesSelection: this.getString(objects, "columns", "seriesSelection", "All"),
            actualColor: this.getFillColor(objects, "columns", "actualColor", ""),
            actualTransparency: Math.max(0, Math.min(100, this.getNumber(objects, "columns", "actualTransparency", 0))),
            benchmarkColor: this.getFillColor(objects, "columns", "benchmarkColor", ""),
            benchmarkTransparency: Math.max(0, Math.min(100, this.getNumber(objects, "columns", "benchmarkTransparency", 70))),
            actualWidth: Math.max(0, Math.min(100, this.getNumber(objects, "columns", "actualWidth", 65))),
            innerPadding: Math.max(0, this.getNumber(objects, "columns", "innerPadding", 20)),
            groupPadding: Math.max(0, this.getNumber(objects, "columns", "groupPadding", 20)),

            // Data labels
            dataLabelsShow: this.getBool(objects, "dataLabels", "show", false),
            dataLabelsColor: this.getFillColor(objects, "dataLabels", "color", "#333333"),
            dataLabelsFontSize: Math.max(8, this.getNumber(objects, "dataLabels", "fontSize", 9)),
            dataLabelsFontFamily: this.getString(objects, "dataLabels", "fontFamily", "Segoe UI"),
            dataLabelsDisplayUnits: this.getNumber(objects, "dataLabels", "displayUnits", 0),
            dataLabelsValueDecimalPlaces: this.getNumberOptional(objects, "dataLabels", "valueDecimalPlaces"),
            dataLabelsPosition: this.getString(objects, "dataLabels", "position", "Auto"),

            // Benchmark labels
            showBenchmarkLabels: this.getBool(objects, "benchmarkLabels", "show", true),
            benchmarkLabelPrefix: this.getString(objects, "benchmarkLabels", "labelPrefix", "National Average: "),
            benchmarkLabelFontSize: Math.max(8, this.getNumber(objects, "benchmarkLabels", "fontSize", 9)),
            benchmarkLabelFontFamily: this.getString(objects, "benchmarkLabels", "fontFamily", "Segoe UI")
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