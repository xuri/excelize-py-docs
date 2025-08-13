# Chart

## Add chart {#AddChart}

```python
def add_chart(sheet: str, cell: str, chart: Chart, **combo: Chart) -> None
```

Add a chart in a worksheet by given chart format set (such as offset, scale, aspect ratio setting, and print settings) and properties set.

The following shows the `type` of chart supported by excelize:

ID|Enumeration|Chart
---|---|---
0  | Area                        | 2D area chart
1  | AreaStacked                 | 2D stacked area chart
2  | AreaPercentStacked          | 2D 100% stacked area chart
3  | Area3D                      | 3D area chart
4  | Area3DStacked               | 3D stacked area chart
5  | Area3DPercentStacked        | 3D 100% stacked area chart
6  | Bar                         | 2D clustered bar chart
7  | BarStacked                  | 2D stacked bar chart
8  | BarPercentStacked           | 2D 100% stacked bar chart
9  | Bar3DClustered              | 3D clustered bar chart
10 | Bar3DStacked                | 3D stacked bar chart
11 | Bar3DPercentStacked         | 3D 100% stacked bar chart
12 | Bar3DConeClustered          | 3D cone clustered bar chart
13 | Bar3DConeStacked            | 3D cone stacked bar chart
14 | Bar3DConePercentStacked     | 3D 100% cone bar chart
15 | Bar3DPyramidClustered       | 3D pyramid clustered bar chart
16 | Bar3DPyramidStacked         | 3D pyramid stacked bar chart
17 | Bar3DPyramidPercentStacked  | 3D 100% pyramid stacked bar chart
18 | Bar3DCylinderClustered      | 3D cylinder clustered bar chart
19 | Bar3DCylinderStacked        | 3D cylinder stacked bar chart
20 | Bar3DCylinderPercentStacked | 3D 100% cylinder stacked bar chart
21 | Col                         | 2D clustered column chart
22 | ColStacked                  | 2D stacked column chart
23 | ColPercentStacked           | 2D 100% stacked column chart
24 | Col3DClustered              | 3D clustered column chart
25 | Col3D                       | 3D column chart
26 | Col3DStacked                | 3D stacked column chart
27 | Col3DPercentStacked         | 3D 100% stacked column chart
28 | Col3DCone                   | 3D cone column chart
29 | Col3DConeClustered          | 3D cone clustered column chart
30 | Col3DConeStacked            | 3D cone stacked column chart
31 | Col3DConePercentStacked     | 3D 100% cone stacked column chart
32 | Col3DPyramid                | 3D pyramid column chart
33 | Col3DPyramidClustered       | 3D pyramid clustered column chart
34 | Col3DPyramidStacked         | 3D pyramid stacked column chart
35 | Col3DPyramidPercentStacked  | 3D 100% pyramid stacked column chart
36 | Col3DCylinder               | 3D cylinder column chart
37 | Col3DCylinderClustered      | 3D cylinder clustered column chart
38 | Col3DCylinderStacked        | 3D cylinder stacked column chart
39 | Col3DCylinderPercentStacked | 3D 100% cylinder stacked column chart
40 | Doughnut                    | doughnut chart
41 | Line                        | line chart
42 | Line3D                      | 3D line chart
43 | Pie                         | pie chart
44 | Pie3D                       | 3D pie chart
45 | PieOfPie                    | pie of pie chart
46 | BarOfPie                    | bar of pie chart
47 | Radar                       | radar chart
48 | Scatter                     | scatter chart
49 | Surface3D                   | 3D surface chart
50 | WireframeSurface3D          | 3D wireframe surface chart
51 | Contour                     | contour chart
52 | WireframeContour            | wireframe contour chart
53 | Bubble                      | bubble chart
54 | Bubble3D                    | 3D bubble chart

In the Office Excel chart data range, `series` specifies the set of information for which data to draw, the legend item (series), and the horizontal (category) axis label.

The `series` options that can be set are:

Parameter|Explanation
---|---
name                | Legend item (series), displayed in the chart legend and formula bar. The `name` parameter is optional. If you don't specify this value, the default will be `Series 1 .. n`. `name` support for formula representation, for example: `Sheet1!$A$1`.
categories          | Horizontal (category) axis label. The `categories` parameter is optional in most chart types, the default is a contiguous sequence of the form `1..n`.
values              | The chart data area, which is the most important parameter in `series`, is also the only required parameter when creating a chart. This option links the chart to the worksheet data it displays.
fill                | This sets the format for the data series fill.
legend              | This set the font of legend text for a data series. The `legend` property is optional.
line                | This sets the line format of the line chart. The `line` property is optional and if it isn't supplied it will default style. The options that can be set is `width`. The range of `width` is 0.25pt - 999pt. If the value of width is outside the range, the default width of the line is 2pt.
marker              | This sets the marker of the line chart and scatter chart. The range of the optional field `size` is 2-72 (default value is `5`). The enumeration value of optional field `symbol` are (default value is `auto`): `circle`, `dash`, `diamond`, `dot`, `none`, `picture`, `plus`, `square`, `star`, `triangle`, `x` and `auto`.
data_label_position | This sets the position of the chart series data label.

Set properties of the chart legend. The options that can be set are:

Parameter|Type|Explanation
---|---|---
position        | `str`  | The position of the chart legend
show_legend_key | `bool` | Set the legend keys shall be shown in data labels
font            | `Font` | Set the font properties of the chart legend text. The properties that can be set are the same as the font object that is used for cell formatting. The font family, size, color, bold, italic, underline, and strike properties can be set

Set the `position` of the chart legend. The default legend position is `right`. This parameter only takes effect when `none` is `False`. The available positions are:

Parameter|Explanation
---|---
none      | Disable legend
top       | On top
bottom    | On bottom
left      | On left
right     | On right
top_right | On top right

The `show_legend_key` parameter set the legend keys shall be shown in data labels. The default value is `False`.

The chart title is set by selecting the `name` parameter of the `title` object, and the title will be displayed above the chart. The parameter `name` supports the use of formula representations, such as `Sheet1!$A$1`, if you do not specify an icon title, the default value is null.

The parameter `show_blanks_as` provides the "Hide and empty cells" setting. The default value is: `gap`. In the Excel application "empty cell is displayed as": "space". The following are optional values for this parameter:

Parameter|Explanation
---|---
gap  | Space
span | Connect data points with straight lines
zero | Zero value

Set chart legend for all data series by `legend` property. The `legend` property is optional.

Set the bubble size in all data series for the bubble chart or 3D bubble chart by `bubble_sizes` property. The `bubble_sizes` property is optional. The default width is `100`, and the value should be great than 0 and less or equal than 300.

Set the doughnut hole size in all data series for the doughnut chart by `hole_size` property. The `hole_size` property is optional. The default width is `75`, and the value should be great than 0 and less or equal than 90.

Specifies that each data marker in the series has a different color by `vary_colors`. The default value is `True`.

The parameter `format` provides settings for parameters such as chart offset, scale, aspect ratio settings, and print properties, as well as those used in the [`add_picture`](image.md#AddPicture) function.

Set the position of the chart plot area by plot area. The properties that can be set are:

Parameter|Type|Default|Explanation
---|---|---|---
second_plot_values   | `int`         | `0`     | Specifies the values in second plot for the `PieOfPie` and `BarOfPie` chart.
show_bubble_size     | `bool`        | `False` | Specifies the bubble size shall be shown in a data label.
show_cat_name        | `bool`        | `True`  | Specifies that the category name shall be shown in the data label. The `show_cat_name` property is optional.
show_data_table      | `bool`        | `False` | Used for add data table under chart, depending on the chart type, only available for area, bar, column and line series type charts.
show_data_table_keys | `bool`        | `False` | Used for add legend key in data table, only works on `show_data_table` is enabled. The `show_data_table_keys` property is optional.
show_leader_lines    | `bool`        | `False` | Specifies that the category name shall be shown in the data label.
show_percent         | `bool`        | `False` | Specifies that the percentage shall be shown in a data label.
show_ser_name        | `bool`        | `False` | Specifies that the series name shall be shown in a data label.
show_val             | `bool`        | `False` | Specifies that the value shall be shown in a data label.
num_fmt              | `ChartNumFmt` | N/A     | Specifies that if linked to source and set custom number format code for data labels. The `num_fmt` property is optional. The default format code is `General`.

Set the primary horizontal and vertical axis options by `x_axis` and `y_axis`.

The properties of `x_axis` that can be set are:

Parameter|Type|Default|Explanation
---|---|---|---
none             | `bool`                        | `False` | Disable axes.
major_grid_lines | `bool`                        | `False` | Specifies major grid lines.
minor_grid_lines | `bool`                        | `False` | Specifies minor grid lines.
tick_label_skip  | `int`                         | `1`     | Specifies how many tick labels to skip between label that is drawn. The `tick_label_skip` property is optional. The default value is auto.
reverse_order    | `bool`                        | `False` | Specifies that the categories or values in reverse order (orientation of the chart). The `reverse_order` property is optional.
maximum          | `Optional[float]`             | `0`     | Specifies that the fixed maximum, 0 is auto. The maximum property is optional.
minimum          | `Optional[float]`             | `0`     | Specifies that the fixed minimum, 0 is auto. The minimum property is optional. The default value is auto.
alignment        | `Alignment`                   | N/A     | Specifies that the alignment of the horizontal and vertical axis. The properties of font that can be set are: `text_rotation` and `vertical`
font             | `Font`                        | N/A     | Specifies that the font of the horizontal axis.
num_fmt          | `ChartNumFmt`                 | N/A     | Specifies that if linked to source and set custom number format code for axis.
title            | `Optional[List[RichTextRun]]` | N/A     | Specifies that the primary horizontal axis title and resize chart.

The properties of `YAxis` that can be set are:

Parameter|Type|Default|Explanation
---|---|---|---
none             | `bool`                        | `False` | Disable axes.
major_grid_lines | `bool`                        | `False` | Specifies major grid lines.
minor_grid_lines | `bool`                        | `False` | Specifies minor grid lines.
major_unit       | `float`                       | `0`     | Specifies the distance between major ticks. Shall contain a positive floating-point number. The `MajorUnit` property is optional. The default value is auto.
reverse_order    | `bool`                        | `False` | Specifies that the categories or values in reverse order (orientation of the chart). The `reverse_order` property is optional.
maximum          | `Optional[float]`             | `0`     | Specifies that the fixed maximum, 0 is auto. The maximum property is optional.
minimum          | `Optional[float]`             | `0`     | Specifies that the fixed minimum, 0 is auto. The minimum property is optional. The default value is auto.
alignment        | `Alignment`                   | N/A     | Specifies that the alignment of the horizontal and vertical axis. The properties of font that can be set are: `text_rotation` and `vertical`
font             | `Font`                        | N/A     | Specifies that the font of the vertical axis.
log_base         | `float64`                     | N/A     | Specifies logarithmic scale base number of the vertical axis.
num_fmt          | `ChartNumFmt`                 | N/A     | Specifies that if linked to source and set custom number format code for axis.
title            | `Optional[List[RichTextRun]]` | N/A     | Specifies that the primary vertical axis title and resize chart.

The value of `text_rotation` that can be set from -90 to 90.

The value of `vertical` that can be set are: `horz`, `vert`, `vert270`, `wordArtVert`, `eaVert`, `mongolianVert` and `wordArtVertRtl`.

Set the chart size by `dimension` property. The dimension property is optional. The properties that can be set are:

Parameter|Type|Default|Explanation
---|---|---|---
height | `int` | 260 | Height
width  | `int` | 480 | Width

The parameter `combo` specifies the create a chart that combines two or more chart types in a single chart. For example, create a clustered column - line chart with data `Sheet1!$E$1:$L$15`:

```python
import excelize

f = excelize.new_file()
try:
    for idx, row in enumerate(
        [
            [None, "Apple", "Orange", "Pear"],
            ["Small", 2, 3, 3],
            ["Normal", 5, 2, 4],
            ["Large", 6, 7, 8],
        ]
    ):
        cell = excelize.coordinates_to_cell_name(1, idx + 1)
        f.set_sheet_row("Sheet1", cell, row)

    f.add_chart(
        "Sheet1",
        "E1",
        excelize.Chart(
            type=excelize.ChartType.Col,
            series=[
                excelize.ChartSeries(
                    name="Sheet1!$A$2",
                    categories="Sheet1!$B$1:$D$1",
                    values="Sheet1!$B$2:$D$2",
                )
            ],
            format=excelize.GraphicOptions(
                scale_x=1,
                scale_y=1,
                offset_x=15,
                offset_y=10,
                print_object=True,
                lock_aspect_ratio=False,
                locked=False,
            ),
            title=[
                excelize.RichTextRun(
                    text="Clustered Column - Line Chart",
                )
            ],
            legend=excelize.ChartLegend(position="left"),
            plot_area=excelize.ChartPlotArea(
                show_cat_name=False,
                show_leader_lines=False,
                show_percent=True,
                show_ser_name=True,
                show_val=True,
            ),
        ),
        combo=excelize.Chart(
            type=excelize.ChartType.Line,
            series=[
                excelize.ChartSeries(
                    name="Sheet1!$A$4",
                    categories="Sheet1!$B$1:$D$1",
                    values="Sheet1!$B$4:$D$4",
                    marker=excelize.ChartMarker(
                        symbol="none",
                        size=10,
                    ),
                )
            ],
            format=excelize.GraphicOptions(
                scale_x=1,
                scale_y=1,
                offset_x=15,
                offset_y=10,
                print_object=True,
                lock_aspect_ratio=False,
                locked=False,
            ),
            legend=excelize.ChartLegend(position="right"),
            plot_area=excelize.ChartPlotArea(
                show_cat_name=False,
                show_leader_lines=False,
                show_percent=True,
                show_ser_name=True,
                show_val=True,
            ),
        ),
    )
    # Save the spreadsheet by the given path.
    f.save_as("Book1.xlsx")
except (RuntimeError, TypeError) as err:
    print(err)
finally:
    err = f.close()
    if err:
        print(err)
```

## Add chart sheet {#AddChartSheet}

```python
def add_chart_sheet(sheet: str, chart: Chart, **combo: Chart) -> None
```

Create a chartsheet by given chart format set (such as offset, scale, aspect ratio setting and print settings) and properties set. In Excel a chartsheet is a worksheet that only contains a chart.

## Delete chart {#DeleteChart}

```python
def delete_chart(sheet: str, cell: str) -> None
```

Delete chart in spreadsheet by given worksheet name and cell reference.
