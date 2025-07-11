# Picture

## Add Picture {#AddPicture}

```python
def add_picture(
    sheet: str, cell: str, name: str, opts: Optional[GraphicOptions]
) -> None
```

Add picture in a sheet by given picture format set (such as offset, scale, aspect ratio setting, and print settings) and file path. Supported image types: GIF, JPEG, JPG, PNG, TIF and TIFF. Note that this function only supports adding pictures placed over the cells currently, and doesn't support adding pictures placed in cells or creating the Kingsoft WPS Office embedded image cells.

For example:

```python
import excelize

try:
    f = excelize.new_file()
except RuntimeError as err:
    print(err)
    exit()
try:
    # Insert a picture.
    f.add_picture("Sheet1", "A2", "image.png", None)
    # Insert a picture to worksheet with scaling.
    f.add_picture("Sheet1", "D2", "image.jpg", excelize.GraphicOptions(
        scale_x=0.5,
        scale_y=0.5,
    ))
    # Insert a picture offset in the cell with printing support.
    f.add_picture("Sheet1", "H2", "image.gif", excelize.GraphicOptions(
        offset_x=15,
        offset_y=10,
        hyperlink="https://github.com/xuri/excelize",
        hyperlink_type="External",
        print_object=True,
        lock_aspect_ratio=False,
        locked=False,
        positioning="oneCell",
    ))
    # Save the spreadsheet with the origin path.
    f.save_as("Book1.xlsx")
except RuntimeError as err:
    print(err)
finally:
    # Close the spreadsheet.
    err = f.close()
    if err:
        print(err)
```

The optional parameter `alt_ext` is used to add alternative text to a graph object.

The optional parameter `print_object` indicates whether the graph object is printed when the worksheet is printed, the default value of that is `True`.

The optional parameter `locked` indicates whether lock the graph object. Locking an object has no effect unless the sheet is protected.

The optional parameter `lock_aspect_ratio` indicates whether lock aspect ratio for the graph object, the default value of that is `False`.

The optional parameter `auto_fit` specifies if make graph object size auto fits the cell, the default value of that is `False`.

The optional parameter `offset_x` specifies the horizontal offset of the graph object with the cell, the default value of that is 0.

The optional parameter `offset_y` specifies the vertical offset of the graph object with the cell, the default value of that is 0.

The optional parameter `scale_x` specifies the horizontal scale of graph object, the default value of that is 1.0 which presents 100%.

The optional parameter `scale_y` specifies the vertical scale of graph object, the default value of that is 1.0 which presents 100%.

The optional parameter `hyperlink` specifies the hyperlink of the graph object.

The optional parameter `hyperlink_type` defines two types of hyperlink `External` for the website or `Location` for moving to one of the cells in this workbook. When the `hyperlink_type` is `Location`, coordinates need to start with `#`.

The optional parameter `positioning` defines 3 types of the position of a graph object in a spreadsheet: `oneCell` (Move but don't size with cells), `twoCell` (Move and size with cells), and `absolute` (Don't move or size with cells). If you don't set this parameter, the default positioning is to move and size with cells.

```python
def add_picture_from_bytes(sheet: str, cell: str, picture: Picture) -> None
```

Add a picture in a sheet by given picture format set (such as offset, scale, aspect ratio setting and print settings), alt text description, extension name and file content in `byte` type. Supported image types:  GIF, JPEG, JPG, PNG, TIF and TIFF. Note that this function only supports adding pictures placed over the cells currently, and doesn't support adding pictures placed in cells or creating the Kingsoft WPS Office embedded image cells.

For example:

```python
import excelize

try:
    f = excelize.new_file()
except RuntimeError as err:
    print(err)
    exit()
try:
    with open("image.jpg", "rb") as file:
        f.add_picture_from_bytes(
            "Sheet1",
            "A2",
            excelize.Picture(
                extension=".jpg",
                file=file.read(),
                format=excelize.GraphicOptions(alt_text="Excel Logo"),
            ),
        )
    f.save_as("Book1.xlsx")
except RuntimeError as err:
    print(err)
finally:
    err = f.close()
    if err:
        print(err)
```

## Delete Picture {#DeletePicture}

```python
def delete_picture(sheet: str, cell: str) -> None
```

Delete all pictures in a cell by given worksheet name and cell reference.
