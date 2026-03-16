# Overlay Benchmark Bar (Power BI Custom Visual)

A Power BI custom visual that overlays a narrow **Actual Value** bar in front of a wider **Benchmark** bar.  
This design provides a clear comparison between performance and a benchmark (for example national averages, targets, or previous periods) without cluttering the chart with side-by-side columns.

---

## Overview

The **Overlay Benchmark Bar** visual is designed for scenarios where comparing a value to its benchmark is more important than comparing multiple categories.

Instead of displaying separate columns, the benchmark acts as a background reference while the actual value sits in the foreground. This makes it immediately obvious whether performance is above or below the benchmark.

Typical use cases include:

- Performance vs national averages  
- KPI tracking against targets  
- Department performance vs organisational benchmarks  
- Product performance vs market averages  

---

## Features

### Dual-Bar Overlay

Displays a wide **Benchmark** bar with a narrower **Actual** bar layered on top.

This design allows users to quickly determine whether a value exceeds or falls below the benchmark.

### Flexible Data Roles

Supports hierarchical grouping for detailed analysis.

| Data Role | Example |
|-----------|--------|
| Base Category | Region, Department |
| Sub Category | Product, Month |
| Actual Value | Performance value |
| Benchmark Value | National average / target |

### Formatting Controls

#### Columns
- Actual colour  
- Benchmark colour  
- Transparency (0–100%)  
- Adjustable actual bar width  
- Group and inner padding  

#### Axes
- Auto-scaling or manual min/max  
- Axis titles and fonts  
- Custom axis styling  

#### Labels
- Actual value labels  
- Benchmark labels  
- Custom prefix text (for example `National Average:`)  
- Minimum accessible font sizes  

#### Gridlines & Legend
- Toggleable gridlines  
- Solid, dashed, or dotted styles  
- Configurable legend placement  

### Interactivity

- Full Power BI cross-filtering support  
- Multi-selection  
- Native Power BI tooltips  

---

## Installation

### 1. Download the visual

Download the compiled `.pbiviz` file from the **Releases** section of this repository.

### 2. Import into Power BI

1. Open **Power BI Desktop**  
2. In the **Visualizations** pane select `...`  
3. Click **Import a visual from a file**  
4. Select the downloaded `.pbiviz` file  

### 3. Add the visual to a report

1. Click the **Overlay Benchmark Bar** visual icon  
2. Add fields to the following data roles  

| Field | Example |
|------|--------|
| Base Category | Region |
| Sub Category | Product |
| Actual Value | Sales |
| Benchmark Value | National Average |

3. Use the **Format Pane** to adjust colours, widths, labels, and axes.

---

## License

This project is licensed under the **MIT License**.

See the `LICENSE` file for details.

---

## Support

If you encounter issues, have feature requests, or need help using the visual:

- Email me at gabe.chisholm@synapsysiq.co.uk
- Provide a short description of the issue and screenshots where possible

Feedback and suggestions for improvements are welcome.

---

## Author

Created by **Gabe Chisholm**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?logo=linkedin)](https://www.linkedin.com/in/gabe-chisholm/)
