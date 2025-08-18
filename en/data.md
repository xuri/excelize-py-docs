# Data

## Add slicer {#AddSlicer}

`SlicerOptions` represents the settings of the slicer.

```python
class SlicerOptions:
    name: str = ""
    cell: str = ""
    table_sheet: str = ""
    table_name: str = ""
    caption: str = ""
    macro: str = ""
    width: int = 0
    height: int = 0
    display_header: Optional[bool] = None
    item_desc: bool = False
    format: GraphicOptions = GraphicOptions
```

`name` specifies the slicer name, should be an existing field name of the given table or pivot table, this setting is required.

`table` specifies the name of the table or pivot table, this setting is required.

`cell` specifies the left top cell coordinates the position for inserting the slicer, this setting is required.

`caption` specifies the caption of the slicer, this setting is optional.

`macro` used for set macro for the slicer, the workbook extension should be XLSM or XLTM.

`width` specifies the width of the slicer, this setting is optional.

`height` specifies the height of the slicer, this setting is optional.

`display_header` specifies if display header of the slicer, this setting is optional, the default setting is display.

`item_desc` specifies descending (Z-A) item sorting, this setting is optional, and the default setting is `false` (represents ascending).

`format` specifies the format of the slicer, this setting is optional.

```python
def add_slicer(self, sheet: str, opts: SlicerOptions) -> None
```

Inserts a slicer by giving the worksheet name and slicer settings. For example, insert a slicer on the `Sheet1!E1` with field `Column1` for the table named `Table1`:

```python
try:
    f.add_slicer(
        "Sheet1",
        excelize.SlicerOptions(
            name="Column1",
            cell="E1",
            table_sheet="Sheet1",
            table_name="Table1",
            caption="Column1",
            width=200,
            height=200,
        ),
    )
except (RuntimeError, TypeError) as err:
    print(err)
```

## Delete slicer {#DeleteSlicer}

```python
def delete_slicer(self, name: str) -> None
```

Delete a slicer by a given slicer name.
