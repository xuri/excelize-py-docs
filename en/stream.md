# Streaming write

## Get stream writer {#NewStreamWriter}

```python
def new_stream_writer(sheet: str) -> StreamWriter
```

NewStreamWriter returns stream writer struct by given worksheet name used for writing data on a new existing empty worksheet with large amounts of data. Note that after writing data with the stream writer for the worksheet, you must call the [`flush`](stream.md#Flush) method to end the streaming writing process, ensure that the order of row numbers is ascending when set rows, and the normal mode functions and stream mode functions can not be work mixed to writing data on the worksheets. The stream writer will try to use temporary files on disk to reduce the memory usage when in-memory chunks data over 16MB, and you can't get cell value at this time. For example, set data for worksheet of size `102400` rows x `50` columns with numbers and style:

```python
import excelize, random

f = excelize.new_file()
try:
    sw = f.new_stream_writer("Sheet1")
    for r in range(2, 102401):
        row = [random.randrange(640000) for _ in range(1, 51)]
        cell = excelize.coordinates_to_cell_name(1, r, False)
        sw.set_row(cell, row)
    sw.flush()
    f.save_as("Book1.xlsx")
except RuntimeError as err:
    print(err)
finally:
    err = f.close()
    if err:
        print(err)
```

## Write sheet row in stream {#SetRow}

```python
def set_row(
    cell: str,
    values: List[Union[None, int, str, bool, datetime, date]],
) -> None
```

Writes an array to stream rows by giving starting cell reference and a pointer to an array of values. Note that you must call the [`flush`](stream.md#Flush) function to end the streaming writing process.

## Add table in stream {#AddTable}

```python
def add_table(table: Table) -> None
```

Creates an Excel table for the stream writer using the given cell range and format set.

Note that the table must be at least two lines including the header. The header cells must contain strings and must be unique. Currently, only one table is allowed for a stream writer. The function must be called after the rows are written but before `flush`.

Example 1, create a table of `A1:D5`:

```python
try:
    sw.add_table(excelize.Table(range="A1:D5"))
except RuntimeError as err:
    print(err)
```

Example 2, create a table of `F2:H6` with format set:

```python
try:
    sw.add_table(
        excelize.Table(
            range="F2:H6,
            name="table",
            style_name="TableStyleMedium2",
            show_first_column=True,
            show_last_column=True,
            show_row_stripes=False,
            show_column_stripes=True,
        )
    )
except RuntimeError as err:
    print(err)
```

Note that the table must be at least two lines including the header. The header cells must contain strings and must be unique. Currently only one table is allowed for a `StreamWriter`. [`add_table`](stream.md#AddTable) must be called after the rows are written but before `flush`. See [`add_able`](utils.md#AddTable) for details on the table format.

## Insert page break in stream {#InsertPageBreak}

```python
def insert_page_break(cell: str) -> None
```

Creates a page break to determine where the printed page ends and where begins the next one by a given cell reference, the content before the page break will be printed on one page and after the page break on another.

## Set panes in stream {#SetPanes}

```python
def set_panes(opts: Panes) -> None
```

Create and remove freeze panes and split panes by giving panes options for the `StreamWriter`. Note that you must call the `set_panes` function before the [`set_row`](stream.md#SetRow) function.

## Merge cell in stream {#MergeCell}

```python
def merge_cell(top_left_cell: str, bottom_right_cell: str) -> None
```

Merge cells by a given range reference for the `StreamWriter`. Don't create a merged cell that overlaps with another existing merged cell.

## Set column width in stream {#SetColWidth}

```python
def set_col_width(start_col: int, end_col: int, width: float) -> None
```

Set the width of a single column or multiple columns for the `StreamWriter`. Note that you must call the `set_col_width` function before the [`set_row`](stream.md#SetRow) function. For example set the width column `B:C` as `20`:

```python
try:
    sw.set_col_width(2, 3, 20)
except RuntimeError as err:
    print(err)
```

## Flush stream {#Flush}

```python
def flush() -> None
```

Ending the streaming writing process.
