# Workbook

`Options` defines the options for reading and writing spreadsheets.

```python
class Options:
    max_calc_iterations: int = 0
    password: str = ""
    raw_cell_value: bool = False
    unzip_size_limit: int = 0
    unzip_xml_size_limit: int = 0
    tmp_dir: str = ""
    short_date_pattern: str = ""
    long_date_pattern: str = ""
    long_time_pattern: str = ""
    culture_info: CultureName = CultureName.CultureNameUnknown
```

`max_calc_iterations` specifies the maximum iterations for iterative calculation, the default value is 0.

`password` specifies the password of the spreadsheet in plain text.

`raw_cell_value` specifies if apply the number format for the cell value or get the raw value.

`unzip_size_limit` specifies the unzip size limit in bytes on open the spreadsheet, this value should be greater than or equal to `unzip_xml_size_limit`, the default size limit is 16GB.

`unzip_xml_size_limit` specifies the memory limit on unzipping worksheet and shared string table in bytes, worksheet XML will be extracted to system temporary directory when the file size is over this value, this value should be less than or equal to `unzip_size_limit`, the default value is 16MB.

`tmp_dir` specifies the temporary directory for creating temporary files, if the value is empty, the system default temporary directory will be used.

`short_date_pattern` specifies the short date number format code. In the spreadsheet applications, date formats display date and time serial numbers as date values. Date formats that begin with an asterisk (\*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings. The `short_date_pattern` used for specifies apply date formats that begin with an asterisk.

`long_date_pattern` specifies the long date number format code.

`long_time_pattern` specifies the long time number format code.

`culture_info` specifies the country code for applying built-in language number format code these effect by the system's local language settings.

## Create Excel document {#NewFile}

```python
def new_file() -> File
```

Create new workbook by default template. The newly created workbook will by default contain a worksheet named `Sheet1`. For example:

## Open {#OpenFile}

```python
def open_file(filename: str, *opts: Options) -> File
```

This function takes the name of a spreadsheet file and returns a populated spreadsheet file struct for it. For example, open a spreadsheet with password protection:

```python
try:
    f = excelize.open_file("Book1.xlsx", excelize.Options(password="password"))
except RuntimeError as err:
    print(err)
```

Close the file by [`close()`](workbook.md#Close) after opening the spreadsheet.

## Open data stream {#OpenReader}

```python
def open_reader(buffer: bytes, *opts: Options) -> Optional[File]
```

Read data stream from `bytes` and return a populated spreadsheet file.

## Save {#Save}

```python
def save(*opts: Options) -> None
```

Override the spreadsheet file with the origin path.

## Save as {#SaveAs}

```python
def save_as(filename: str, *opts: Options) -> None
```

Create or update the spreadsheet file at the provided path.

## Close workbook {#Close}

```python
def close() -> (Exception | None)
```

Closes and cleanup the open temporary file for the spreadsheet.

## Create worksheet {#NewSheet}

```python
def new_sheet(sheet: str) -> int
```

Create a new sheet by given a worksheet name and returns the index of the sheets in the workbook (spreadsheet) after it appended. Note that when creating a new spreadsheet file, the default worksheet named `Sheet1` will be created.

## Delete worksheet {#DeleteSheet}

```python
def delete_sheet(sheet: str) -> None
```

Delete worksheet in a workbook by given worksheet name. Use this method with caution, which will affect changes in references such as formulas, charts, and so on. If there is any referenced value of the deleted worksheet, it will cause a file error when you open it. This function will be invalid when only one worksheet is left.

## Move worksheet {#MoveSheet}

```python
def move_sheet(source: str, target: str) -> None
```

Moves a sheet to a specified position in the workbook. The function moves the source sheet before the target sheet. After moving, other sheets will be shifted to the left or right. If the sheet is already at the target position, the function will not perform any action. Not that this function will be ungroup all sheets after moving. For example, move `Sheet2` before `Sheet1`:

```python
f.move_sheet("Sheet2", "Sheet1")
```

## Copy worksheet {#CopySheet}

```python
def copy_sheet(src: int, to: int) -> None
```

Duplicate a worksheet by gave source and target worksheet index. Note that currently doesn't support duplicate workbooks that contain tables, charts or pictures. For example:

```python
try:
    # Sheet1 already exists...
    index = f.new_sheet("Sheet2")
    f.copy_sheet(0, index)
except RuntimeError as err:
    print(err)
```

## Group worksheets {#GroupSheets}

```python
def group_sheets(sheets: List[str]) -> None
```

Group worksheets by given worksheets names. Group worksheets must contain an active worksheet.

## Ungroup worksheets {#UngroupSheets}

```python
def ungroup_sheets() -> None
```

Ungroup worksheets.

## Set worksheet background {#SetSheetBackground}

```python
def set_sheet_background(sheet: str, picture: str) -> None
```

Set background picture by given worksheet name and file path. Supported image types: BMP, EMF, EMZ, GIF, JPEG, JPG, PNG, SVG, TIF, TIFF, WMF, and WMZ.

```python
def set_sheet_background_from_bytes(sheet: str, extension: str, picture: bytes) -> None
```

Set background picture by given worksheet name, extension name and image data. Supported image types: BMP, EMF, EMZ, GIF, JPEG, JPG, PNG, SVG, TIF, TIFF, WMF, and WMZ.

## Set default worksheet {#SetActiveSheet}

```python
def set_active_sheet(index: int) -> None
```

Set the default active sheet of the workbook by a given index. Note that the active index is different from the ID returned by function [`get_sheet_map`](sheet.md#GetSheetMap). It should be greater or equal to 0 and less than the total worksheet numbers.

## Get active sheet index {#GetActiveSheetIndex}

```python
def get_active_sheet_index() -> int
```

Get an active worksheet of the workbook. If not found the active sheet will return integer `0`.

## Set worksheet visible {#SetSheetVisible}

```python
def set_sheet_visible(sheet: str, visible: bool, *very_hidden: bool) -> None
```

Set worksheet visible by given worksheet name. A workbook must contain at least one visible worksheet. If the given worksheet has been activated, this setting will be invalidated. The third optional `very_hidden` parameter only works when `visible` was `False`.

For example, hide `Sheet1`:

```python
f.set_sheet_visible("Sheet1", False)
```

## Get worksheet visible {#GetSheetVisible}

```python
def get_sheet_visible(sheet: str) -> bool
```

Get visible state of the sheet by given sheet name. For example, get the visible state of `Sheet1`:

```python
try:
    visible = f.get_sheet_visible("Sheet1")
except RuntimeError as err:
    print(err)
```

## Set worksheet properties {#SetSheetProps}

```python
def set_sheet_props(sheet: str, opts: SheetPropsOptions) -> None
```

Set worksheet properties. The properties that can be set are:

Options|Type|Description
---|---|---
code_name                            | `Optional[str]`   | Specifies a stable name of the sheet, which should not change over time, and does not change from user input. This name should be used by code to reference a particular sheet
enable_format_conditions_calculation | `Optional[bool]`  | Indicating whether the conditional formatting calculations shall be evaluated. If set to false, then the min/max values of color scales or data bars or threshold values in Top N rules shall not be updated. Essentially the conditional formatting "calc" is off
published                            | `Optional[bool]`  | Indicating whether the worksheet is published, the default value is `True`
auto_page_breaks                     | `Optional[bool]`  | Indicating whether the sheet displays Automatic Page Breaks, the default value is `True`
fit_to_page                          | `Optional[bool]`  | Indicating whether the Fit to Page print option is enabled, the default value is `False`
tab_color_indexed                    | `Optional[int]`   | Represents the indexed color value
tab_color_rgb                        | `Optional[str]`   | Represents the standard ARGB (Alpha Red Green Blue) color value
tab_color_theme                      | `Optional[int]`   | Represents the zero-based index into the collection, referencing a particular value expressed in the Theme part
tab_color_tint                       | `Optional[float]` | Specifies the tint value applied to the color, the default value is `0.0`
outline_summary_below                | `Optional[bool]`  | Indicating whether summary rows appear below detail in an outline, when applying an outline, the default value is `True`
outline_summary_right                | `Optional[bool]`  | Indicating whether summary columns appear to the right of detail in an outline, when applying an outline, the default value is `True`
base_col_width                       | `Optional[int]`   | Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for grid lines. It is only the number of characters, the default value is `8`
default_col_width                    | `Optional[float]` | Specifies the default column width measured as the number of characters of the maximum digit width of the normal style's font
default_row_height                   | `Optional[float]` | Specifies the default row height measured in point size. Optimization so we don't have to write the height on all rows. This can be written out if most rows have custom height, to achieve the optimization
custom_height                        | `Optional[bool]`  | Specifies the custom height, the default value is `False`
zero_height                          | `Optional[bool]`  | Specifies if rows are hidden, the default value is `False`
thick_top                            | `Optional[bool]`  | Specifies if rows have a thick top border by default, the default value is `False`
thick_bottom                         | `Optional[bool]`  | Specifies if rows have a thick bottom border by default, the default value is `False`

For example, make worksheet rows default as hidden:

<p align="center"><img width="612" src="https://xuri.me/excelize/en/images/sheet_format_pr_01.png" alt="Set worksheet properties"></p>

```python
f = excelize.new_file()
f.set_sheet_props("Sheet1", excelize.SheetPropsOptions(
    zero_height=True
))
f.set_row_visible("Sheet1", 10, True)
f.save_as("Book1.xlsx")
```

There 4 kinds of presets "Custom Scaling Options" in the spreadsheet applications, if you need to set those kind of scaling options, please using the `set_sheet_props` and `set_page_layout` functions to approach these 4 scaling options:

1. No Scaling (Print sheets at their actual size):

    ```python
    f.set_sheet_props("Sheet1", excelize.SheetPropsOptions(
        fit_to_page=False
    ))
    ```

2. Fit Sheet on One Page (Shrink the printout so that it fits on one page):

    ```python
    f.set_sheet_props("Sheet1", excelize.SheetPropsOptions(
        fit_to_page=True
    ))
    ```

3. Fit All Columns on One Page (Shrink the printout so that it is one page wide):

    ```python
    f.set_sheet_props("Sheet1", excelize.SheetPropsOptions(
        fit_to_page=True
    ))
    f.set_page_layout("Sheet1", excelize.PageLayoutOptions(
        fit_to_height=0,
    ))
    ```

4. Fit All Rows on One Page (Shrink the printout so that it is one page high):

    ```python
    f.set_sheet_props("Sheet1", excelize.SheetPropsOptions(
        fit_to_page=True
    ))
    f.set_page_layout("Sheet1", excelize.PageLayoutOptions(
        fit_to_width=0,
    ))
    ```

## Get worksheet properties {#GetSheetProps}

```python
def get_sheet_props(sheet_name: str) -> Optional[SheetPropsOptions]
```

Get worksheet properties.

## Set worksheet view properties {#SetSheetView}

```python
def set_sheet_view(sheet: str, view_index: int, opts: ViewOptions) -> None
```

Sets sheet view properties. The `view_index` may be negative and if so is counted backward (`-1` is the last view). The properties that can be set are:

Options|Type|Description
---|---|---
default_grid_color   | `Optional[bool]`  | Indicating that the consuming application should use the default grid lines color(system dependent). Overrides any color specified in colorId, the default value is `True`
right_to_left        | `Optional[bool]`  | Indicating whether the sheet is in "right to left" display mode. When in this mode, Column A is on the far right, Column B; is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format, the default value is `false`
show_formulas        | `Optional[bool]`  | Indicating whether this sheet should display formulas, the default value is `False`
show_grid_lines      | `Optional[bool]`  | Indicating whether this sheet should display grid lines, the default value is `True`
show_row_col_headers | `Optional[bool]`  | Indicating whether the sheet should display row and column headings, the default value is `True`
show_ruler           | `Optional[bool]`  | Indicating this sheet should display ruler, the default value is `True`
show_zeros           | `Optional[bool]`  | Indicating whether to "show a zero in cells that have zero value". When using a formula to reference another cell which is empty, the referenced value becomes `0` when the flag is `True`, the default value is `True`
top_left_cell        | `Optional[str]`   | Specifies a location of the top left visible cell Location of the top left visible cell in the bottom right pane (when in Left-to-Right mode)
view                 | `Optional[str]`   | Indicating how sheet is displayed, by default it uses empty string, available options: `normal`，`pageBreakPreview` and `pageLayout`
zoom_scale           | `Optional[float]` | Specifies a window zoom magnification for current view representing percent values. This attribute is restricted to values ranging from `10` to `400`. Horizontal & Vertical scale together, the default value is `100`

## Set worksheet page layout {#SetPageLayout}

```python
def set_page_layout(sheet: str, opts: PageLayoutOptions) -> None
```

Sets worksheet page layout. Available options:

`size` specified the worksheet paper size, the default paper size of worksheet is "Letter paper (8.5 in. by 11 in.)". The following shows the paper size sorted by Excelize index number:

Index|Paper Size
---|---
1   | Letter paper (8.5 in. × 11 in.)
2   | Letter small paper (8.5 in. × 11 in.)
3   | Tabloid paper (11 in. × 17 in.)
4   | Ledger paper (17 in. × 11 in.)
5   | Legal paper (8.5 in. × 14 in.)
6   | Statement paper (5.5 in. × 8.5 in.)
7   | Executive paper (7.25 in. × 10.5 in.)
8   | A3 paper (297 mm × 420 mm)
9   | A4 paper (210 mm × 297 mm)
10  | A4 small paper (210 mm × 297 mm)
11  | A5 paper (148 mm × 210 mm)
12  | B4 paper (250 mm × 353 mm)
13  | B5 paper (176 mm × 250 mm)
14  | Folio paper (8.5 in. × 13 in.)
15  | Quarto paper (215 mm × 275 mm)
16  | Standard paper (10 in. × 14 in.)
17  | Standard paper (11 in. × 17 in.)
18  | Note paper (8.5 in. × 11 in.)
19  | #9 envelope (3.875 in. × 8.875 in.)
20  | #10 envelope (4.125 in. × 9.5 in.)
21  | #11 envelope (4.5 in. × 10.375 in.)
22  | #12 envelope (4.75 in. × 11 in.)
23  | #14 envelope (5 in. × 11.5 in.)
24  | C paper (17 in. × 22 in.)
25  | D paper (22 in. × 34 in.)
26  | E paper (34 in. × 44 in.)
27  | DL envelope (110 mm × 220 mm)
28  | C5 envelope (162 mm × 229 mm)
29  | C3 envelope (324 mm × 458 mm)
30  | C4 envelope (229 mm × 324 mm)
31  | C6 envelope (114 mm × 162 mm)
32  | C65 envelope (114 mm × 229 mm)
33  | B4 envelope (250 mm × 353 mm)
34  | B5 envelope (176 mm × 250 mm)
35  | B6 envelope (176 mm × 125 mm)
36  | Italy envelope (110 mm × 230 mm)
37  | Monarch envelope (3.875 in. × 7.5 in.)
38  | 6¾ envelope (3.625 in. × 6.5 in.)
39  | US standard fanfold (14.875 in. × 11 in.)
40  | German standard fanfold (8.5 in. × 12 in.)
41  | German legal fanfold (8.5 in. × 13 in.)
42  | ISO B4 (250 mm × 353 mm)
43  | Japanese postcard (100 mm × 148 mm)
44  | Standard paper (9 in. × 11 in.)
45  | Standard paper (10 in. × 11 in.)
46  | Standard paper (15 in. × 11 in.)
47  | Invite envelope (220 mm × 220 mm)
50  | Letter extra paper (9.275 in. × 12 in.)
51  | Legal extra paper (9.275 in. × 15 in.)
52  | Tabloid extra paper (11.69 in. × 18 in.)
53  | A4 extra paper (236 mm × 322 mm)
54  | Letter transverse paper (8.275 in. × 11 in.)
55  | A4 transverse paper (210 mm × 297 mm)
56  | Letter extra transverse paper (9.275 in. × 12 in.)
57  | SuperA/SuperA/A4 paper (227 mm × 356 mm)
58  | SuperB/SuperB/A3 paper (305 mm × 487 mm)
59  | Letter plus paper (8.5 in. × 12.69 in.)
60  | A4 plus paper (210 mm × 330 mm)
61  | A5 transverse paper (148 mm × 210 mm)
62  | JIS B5 transverse paper (182 mm × 257 mm)
63  | A3 extra paper (322 mm × 445 mm)
64  | A5 extra paper (174 mm × 235 mm)
65  | ISO B5 extra paper (201 mm × 276 mm)
66  | A2 paper (420 mm × 594 mm)
67  | A3 transverse paper (297 mm × 420 mm)
68  | A3 extra transverse paper (322 mm × 445 mm)
69  | Japanese Double Postcard (200 mm × 148 mm)
70  | A6 (105 mm × 148 mm)
71  | Japanese Envelope Kaku #2
72  | Japanese Envelope Kaku #3
73  | Japanese Envelope Chou #3
74  | Japanese Envelope Chou #4
75  | Letter Rotated (11 in. × 8½ in.)
76  | A3 Rotated (420 mm × 297 mm)
77  | A4 Rotated (297 mm × 210 mm)
78  | A5 Rotated (210 mm × 148 mm)
79  | B4 (JIS) Rotated (364 mm × 257 mm)
80  | B5 (JIS) Rotated (257 mm × 182 mm)
81  | Japanese Postcard Rotated (148 mm × 100 mm)
82  | Double Japanese Postcard Rotated (148 mm × 200 mm)
83  | A6 Rotated (148 mm × 105 mm)
84  | Japanese Envelope Kaku #2 Rotated
85  | Japanese Envelope Kaku #3 Rotated
86  | Japanese Envelope Chou #3 Rotated
87  | Japanese Envelope Chou #4 Rotated
88  | B6 (JIS) (128 mm × 182 mm)
89  | B6 (JIS) Rotated (182 mm × 128 mm)
90  | 12 in. × 11 in.
91  | Japanese Envelope You #4
92  | Japanese Envelope You #4 Rotated
93  | PRC 16K (146 mm × 215 mm)
94  | PRC 32K (97 mm × 151 mm)
95  | PRC 32K(Big) (97 mm × 151 mm)
96  | PRC Envelope #1 (102 mm × 165 mm)
97  | PRC Envelope #2 (102 mm × 176 mm)
98  | PRC Envelope #3 (125 mm × 176 mm)
99  | PRC Envelope #4 (110 mm × 208 mm)
100 | PRC Envelope #5 (110 mm × 220 mm)
101 | PRC Envelope #6 (120 mm × 230 mm)
102 | PRC Envelope #7 (160 mm × 230 mm)
103 | PRC Envelope #8 (120 mm × 309 mm)
104 | PRC Envelope #9 (229 mm × 324 mm)
105 | PRC Envelope #10 (324 mm × 458 mm)
106 | PRC 16K Rotated
107 | PRC 32K Rotated
108 | PRC 32K(Big) Rotated
109 | PRC Envelope #1 Rotated (165 mm × 102 mm)
110 | PRC Envelope #2 Rotated (176 mm × 102 mm)
111 | PRC Envelope #3 Rotated (176 mm × 125 mm)
112 | PRC Envelope #4 Rotated (208 mm × 110 mm)
113 | PRC Envelope #5 Rotated (220 mm × 110 mm)
114 | PRC Envelope #6 Rotated (230 mm × 120 mm)
115 | PRC Envelope #7 Rotated (230 mm × 160 mm)
116 | PRC Envelope #8 Rotated (309 mm × 120 mm)
117 | PRC Envelope #9 Rotated (324 mm × 229 mm)
118 | PRC Envelope #10 Rotated (458 mm × 324 mm)

`orientation` specified worksheet orientation, the default orientation is `portrait`. The possible values for this field is `portrait` and `landscape`.

`first_page_number` specified the first printed page number. If no value is specified, then "automatic" is assumed.

`adjust_to` specified the print scaling. This attribute is restricted to values ranging from 10 (10%) to 400 (400%). This setting is overridden when `fit_to_width` and/or `fit_to_height` are in use.

`fit_to_height` specified the number of vertical pages to fit on.

`fit_to_width` specified the number of horizontal pages to fit on.

`black_and_white` specified print black and white.

`page_order` specifies the ordering of multiple pages. Values accepted: `overThenDown` and `downThenOver`.

For example, set page layout for `Sheet1` with print black and white, first printed page number from `2`, landscape A4 small paper (210 mm by 297 mm), 2 vertical pages to fit on, and 2 horizontal pages to fit:

```python
f = excelize.new_file()
f.set_page_layout("Sheet1", excelize.PageLayoutOptions(
    size=10,
    orientation="landscape",
    first_page_number=2,
    adjust_to=100,
    fit_to_height=2,
    fit_to_width=2,
    black_and_white=True,
))
```

## Set worksheet page margins {#SetPageMargins}

```python
def set_page_margins(sheet: str, opts: PageLayoutMarginsOptions) -> None
```

Set worksheet page margins. Available options:

Options|Type|Description
---|---|---
bottom       | `Optional[float]` | Bottom
footer       | `Optional[float]` | Footer
header       | `Optional[float]` | Header
left         | `Optional[float]` | Left
right        | `Optional[float]` | Right
top          | `Optional[float]` | Top
horizontally | `Optional[bool]`  | Center on page: Horizontally
vertically   | `Optional[bool]`  | Center on page: Vertically

## Set workbook properties {#SetWorkbookProps}

```python
def set_workbook_props(opts: WorkbookPropsOptions) -> None
```

Sets workbook properties. Available options:

Options|Type|Description
---|---|---
date1904       | `Optional[bool]` | Indicates whether to use a 1900 or 1904 date system when converting serial date-times in the workbook to dates.
filter_privacy | `Optional[bool]` | Specifies a boolean value that indicates whether the application has inspected the workbook for personally identifying information (PII). If this flag is set, the application warns the user any time the user performs an action that will insert PII into the document.
code_name      | `Optional[str]`  | Specifies the codename of the application that created this workbook. Use this attribute to track file content in incremental releases of the application.

## Get workbook properties {#GetWorkbookProps}

```python
def get_workbook_props() -> WorkbookPropsOptions
```

Get all tables in a worksheet by given worksheet name.

## Set header and footer {#SetHeaderFooter}

```python
def set_header_footer(sheet: str, opts: HeaderFooterOptions) -> None
```

Set headers and footers by given worksheet name and the control characters.

Headers and footers are specified using the following settings fields:

Fields             | Description
---|---
align_with_margins | Align header footer margins with page margins
different_first    | Different first-page header and footer indicator
different_odd_even | Different odd and even page headers and footers indicator
scale_with_doc     | Scale header and footer with document scaling
odd_header         | Odd Page Footer, or primary Page Footer if `different_odd_even` is `False`
odd_footer         | Odd Header, or primary Page Header if `different_odd_even` is `False`
even_header        | Even Page Footer
even_footer        | Even Page Header
first_header       | First Page Footer
first_footer       | First Page Header

The following formatting codes can be used in 6 string type fields: `odd_header`, `odd_footer`, `even_header`, `even_footer`, `first_header`, `first_footer`

<table>
    <thead>
        <tr>
            <th>Formatting Code</th>
            <th>Description</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><code>&amp;&amp;</code></td>
            <td>The character &quot;&amp;&quot;</td>
        </tr>
        <tr>
            <td><code>&amp;font-size</code></td>
            <td>Size of the text font, where font-size is a decimal font size in points</td>
        </tr>
        <tr>
            <td><code>&amp;&quot;font name,font type&quot;</code></td>
            <td>A text font-name string, font name, and a text font-type string, font type</td>
        </tr>
        <tr>
            <td><code>&amp;&quot;-,Regular&quot;</code></td>
            <td>Regular text format. Toggles bold and italic modes to off</td>
        </tr>
        <tr>
            <td><code>&amp;A</code></td>
            <td>Current worksheet&#39;s tab name</td>
        </tr>
        <tr>
            <td><code>&amp;B</code> or <code>&amp;&quot;-,Bold&quot;</code></td>
            <td>Bold text format, from off to on, or vice versa. The default mode is off</td>
        </tr>
        <tr>
            <td><code>&amp;D</code></td>
            <td>Current date</td>
        </tr>
        <tr>
            <td><code>&amp;C</code></td>
            <td>Center section</td>
        </tr>
        <tr>
            <td><code>&amp;E</code></td>
            <td>Double-underline text format</td>
        </tr>
        <tr>
            <td><code>&amp;F</code></td>
            <td>Current workbook&#39;s file name</td>
        </tr>
        <tr>
            <td><code>&amp;G</code></td>
            <td>Drawing object as background (Use AddHeaderFooterImage)</td>
        </tr>
        <tr>
            <td><code>&amp;H</code></td>
            <td>Shadow text format</td>
        </tr>
        <tr>
            <td><code>&amp;I</code> or <code>&amp;&quot;-,Italic&quot;</code></td>
            <td>Italic text format</td>
        </tr>
        <tr>
            <td><code>&amp;K</code></td>
            <td>Text font color<br>An RGB Color is specified as RRGGBB<br>A Theme Color is specified as TTSNNN where TT is the theme color Id, S is either &quot;+&quot; or &quot;-&quot; of the tint/shade value, and NNN is the tint/shade value</td>
        </tr>
        <tr>
            <td><code>&amp;L</code></td>
            <td>Left section</td>
        </tr>
        <tr>
            <td><code>&amp;N</code></td>
            <td>Total number of pages</td>
        </tr>
        <tr>
            <td><code>&amp;O</code></td>
            <td>Outline text format</td>
        </tr>
        <tr>
            <td><code>&amp;P[[+\|-]n]</code></td>
            <td>Without the optional suffix, the current page number in decimal</td>
        </tr>
        <tr>
            <td><code>&amp;R</code></td>
            <td>Right section</td>
        </tr>
        <tr>
            <td><code>&amp;S</code></td>
            <td>Strikethrough text format</td>
        </tr>
        <tr>
            <td><code>&amp;T</code></td>
            <td>Current time</td>
        </tr>
        <tr>
            <td><code>&amp;U</code></td>
            <td>Single-underline text format. If double-underline mode is on, the next occurrence in a section specifier toggles double-underline mode to off; otherwise, it toggles single-underline mode, from off to on, or vice versa. The default mode is off</td>
        </tr>
        <tr>
            <td><code>&amp;X</code></td>
            <td>Superscript text format</td>
        </tr>
        <tr>
            <td><code>&amp;Y</code></td>
            <td>Subscript text format</td>
        </tr>
        <tr>
            <td><code>&amp;Z</code></td>
            <td>Current workbook&#39;s file path</td>
        </tr>
    </tbody>
</table>

For example:

```python
f.set_header_footer("Sheet1", excelize.HeaderFooterOptions(
    different_first=True,
    different_odd_even=True,
    odd_header="&R&P",
    odd_footer="&C&F",
    even_header="&L&P",
    even_footer="&L&D&R&T",
    first_header="&CCenter &\"-,Bold\"Bold&\"-,Regular\"HeaderU+000A&D"
))
```

This example shows:

- The first page has its own header and footer
- Odd and even-numbered pages have different headers and footers
- Current page number in the right section of odd-page headers
- Current workbook's file name in the center section of odd-page footers
- Current page number in the left section of even-page headers
- Current date in the left section and the current time in the right section of even-page footers
- The text "Center Bold Header" on the first line of the center section of the first page, and the date on the second line of the center section of that same page
- No footer on the first page

## Set defined name {#SetDefinedName}

```python
def set_defined_name(defined_name: DefinedName) -> None
```

Set the defined names of the workbook or worksheet. If not specified scope, the default scope is the workbook. For example:

```python
f.set_defined_name(excelize.DefinedName(
    name="Amount",
    refers_to="Sheet1!$A$2:$D$5",
    comment="defined name comment",
    scope="Sheet2",
))
```

Print area and print titles settings for the worksheet:

<p align="center"><img width="628" src="https://xuri.me/excelize/en/images/page_setup_01.png" alt="Print area and print titles settings for the worksheet"></p>

```python
f.set_defined_name(excelize.DefinedName(
    name="_xlnm.Print_Area",
    refers_to="Sheet1!$A$1:$Z$100",
    scope="Sheet1"
))
f.set_defined_name(excelize.DefinedName(
    name="_xlnm.Print_Titles",
    refers_to="Sheet1!$A:$A,Sheet1!$1:$1",
    scope="Sheet1"
))
```

If you fill the `refers_to` property with only one columns range without a comma, it will work as "Columns to repeat at left" only. For example:

```python
f.set_defined_name(excelize.DefinedName(
    name="_xlnm.Print_Titles",
    refers_to="Sheet1!$A:$A",
    scope="Sheet1"
))
```

If you fill the `refers_to` property with only one rows range without a comma, it will work as "Rows to repeat at top" only. For example:

```python
f.set_defined_name(excelize.DefinedName(
    name="_xlnm.Print_Titles",
    refers_to="Sheet1!$1:$1",
    scope="Sheet1"
))
```

## Delete defined name {#DeleteDefinedName}

```python
def delete_defined_name(defined_name: DefinedName) -> None
```

Delete the defined names of the workbook or worksheet. If not specified scope, the default scope is workbook. For example:

```python
try:
    f.delete_defined_name(excelize.DefinedName(
        name="Amount",
        scope="Sheet2",
    ))
except RuntimeError as err:
    print(err)
```

## Get application properties {#GetAppProps}

```python
def get_app_props() -> Optional[AppProperties]
```

Get document application properties.

## Set document properties {#SetDocProps}

```python
def set_doc_props(doc_properties: DocProperties) -> None
```

Set document core properties. The properties that can be set are:

Property         | Description
---|---
category         | A categorization of the content of this package.
content_status   | The status of the content. For example: Values might include "Draft", "Reviewed" and "Final"
created          | The created time of the content of the resource which represent in ISO 8601 UTC format, for example `2019-06-04T22:00:10Z`.
creator          | An entity primarily responsible for making the content of the resource.
description      | An explanation of the content of the resource.
identifier       | An unambiguous reference to the resource within a given context.
keywords         | A delimited set of keywords to support searching and indexing. This is typically a list of terms that are not available elsewhere in the properties.
last_modified_by | The language of the intellectual content of the resource.
modified         | The user who performed the last modification. The identification is environment-specific.
revision         | The modified time of the content of the resource which represent in ISO 8601 UTC format, for example `2019-06-04T22:00:10Z`.
subject          | The revision number of the content of the resource.
title            | The topic of the content of the resource.
language         | The name given to the resource.
version          | The version number. This value is set by the user or by the application.

For example:

```python
try:
    f.set_doc_props(
        excelize.DocProperties(
            category="category",
            content_status="Draft",
            created="2019-06-04T22:00:10Z",
            creator="Excelize for Python",
            description="This file created by Excelize for Python",
            identifier="xlsx",
            keywords="Spreadsheet",
            last_modified_by="Author Name",
            modified="2019-06-04T22:00:10Z",
            revision="0",
            subject="Test Subject",
            title="Test Title",
            language="en-US",
            version="1.0.0",
        )
    )
except RuntimeError as err:
    print(err)
```

## Protect workbook {#ProtectWorkbook}

```python
def protect_workbook(opts: WorkbookProtectionOptions) -> None
```

Prevent other users from accidentally or deliberately changing, moving, or deleting data in a workbook. The optional field `algorithm_name` specified hash algorithm, support XOR, MD4, MD5, SHA-1, SHA2-56, SHA-384, and SHA-512 currently, if no hash algorithm specified, will be using the XOR algorithm as default. For example, protect workbook with protection settings:

```python
try:
    f.protect_workbook(excelize.WorkbookProtectionOptions(
        password="password",
        lock_structure=True,
    ))
except RuntimeError as err:
    print(err)
```

WorkbookProtectionOptions directly maps the settings of workbook protection.

```python
class WorkbookProtectionOptions:
    algorithm_name: str = ""
    password: str = ""
    lock_structure: bool = False
    lock_windows: bool = False
```

## Unprotect workbook {#UnprotectWorkbook}

```python
def unprotect_workbook(*password: str) -> None
```

Remove protection for workbook, specified the optional password parameter to remove workbook protection with password verification.
