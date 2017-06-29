XlsxWriter
----------------
XlsxWriter is a Python module that can be used to write text, numbers, formulas and hyperlinks to multiple worksheets in an Excel 2007+ XLSX file.


To install
----------------
`pip install XlsxWriter`


`Workbook()` takes one, non-optional, argument which is the filename that we want to create:

Throughout XlsxWriter, rows and columns are zero indexed. The first cell in a worksheet, A1, is (0, 0).

XlsxWriter supports Excels worksheet limits of 1,048,576 rows by 16,384 columns.

Format Class:  ---> The properties of a cell that can be formatted include: fonts, colors, patterns, borders, alignment and number formatting.

Excel treats different types of input data, such as strings and numbers, differently although it generally does it transparently to the user. XlsxWriter tries to emulate this in the `worksheet.write()` method by mapping Python data types to types that Excel supports. Thus the write() method acts as a general alias for several more specific methods:

`set_column()` method to adjust the width of column


Workbook class
----------------

The Workbook class is the main class exposed by the XlsxWriter module, represents the entire spreadsheet as you see it in Excel and internally it represents the Excel file as it is written on disk.

To avoid the use of temporary files in the assembly of the final XLSX file, for example on servers that don’t allow temp files such as the Google APP Engine, set the in_memory constructor option to True:

`workbook.add_worksheet()`: Add a new worksheet to a workbook.

`workbook.add_chart()` : Create a chart object that can be added to a worksheet

`workbook.add_chartsheet()` Add a new add_chartsheet to a workbook.

`workbook.set_size()`: Set the size of a workbook window. width (int) – Width of the window in pixels.
height (int) – Height of the window in pixels. - workbook,set_size(1200, 800)

`workbook.set_properties()`: Set the document properties such as Title, Author etc.

`workbook.set_custom_property()`: Set a custom document property.

`workbook.define_name()`: Create a defined name in the workbook to use as a variable.

`worksheet.set_row()`: set properties for a row of cells

`worksheet.insert_image()`: insert an image in a worksheet cell

`worksheet.add_sparkline()`: Sparklines are small charts that fit in a single cell and are used to show trends in data.


Page set-up methods
----------------
Page set-up methods affect the way that a worksheet looks when it is printed. They control features such as paper size, orientation, page headers and margins.

`worksheet.set_portrait()`: - Set page potrait

`worksheet.set_landscape()`: - Set page landscape

`worksheet.set_page_view():` Set the page view mode.

`worksheet.set_margins()`:

`worksheet.set_header()`:
eg $worksheet->set_header('&CHello');

`worksheet.set_footer()`: Set the printed page footer caption and options.

`worksheet.repeat_rows()`
