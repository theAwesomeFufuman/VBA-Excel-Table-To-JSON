# VBA-Excel-Table-To-JSON
### What is this?
A small VBA script that converts Excel tables to JSON format and exports the data to a .json file at the location of your choice. Use the script by importing the .bas and .frm & .frx files to your Excel VBA editor.

### Installation
You can use this script by following these steps:
1. Open up Microsoft Excel
2. Go to the **Developer** tab (For more information on how to show the developer tab, go [here](https://support.office.com/en-us/article/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45?omkt=en-001&ui=en-US&rs=en-001&ad=US))
3. Click on **Visual Basic**, in the upper left corner of the window
4. In the toolbar at the top of the window that appears, click on **file 👉 Import file...**
5. Select **ExcelToJSON.bas** and click on **Open**
6. Click on **file 👉 Import file...** for a second time
7. Select **ExcelToJSONForm.frm** and click on **Open** (make sure that **ExcelToJSONForm.frx** is located in the same folder, or this step will not work)
8. Congratulation, you have successfully installed the script **🎉🥳**

### Usage
To use the script, you need an Excel file with at least one table in it. Once you do, follow these instructions:
1. Go to the **Developer** tab
2. Click on **Macros**
3. Select **PERSONAL.XLSB!ExcelToJSON.ExcelToJSON**
4. Click on **Run**
5. In the window that appears, select which tables that you would like to export, and then click on **Submit**
6. Finally select the name for the JSON file that will be selected as well as the location that you would like to save the file in

Please note that the script ['escapes'](https://en.wikipedia.org/wiki/Escape_character#JavaScript) all double quotes that exists in the table cells to stay compatible with the JSON format. This means that a cell that contains the text `Lorem ipsum "dolor sit" amet, consectetur` will be edited to look like `Lorem ipsum \"dolor sit\" amet, consectetur`.

The script prints out the cell values of all cell contents as strings, and will not print out the formulas used

The script does not support unicode characters yet, but this will be implemented in a future version.

### Contact
You can reach me at otto[dot]wretling[at]gmail[dot]com

### License
> MIT License
> 
> Copyright (c) 2020 theAwesomeFufuman
> 
> Permission is hereby granted, free of charge, to any person obtaining
> a copy of this software and associated documentation files (the
> "Software"), to deal in the Software without restriction, including
> without limitation the rights to use, copy, modify, merge, publish,
> distribute, sublicense, and/or sell copies of the Software, and to
> permit persons to whom the Software is furnished to do so, subject to
> the following conditions:
> 
> The above copyright notice and this permission notice shall be
> included in all copies or substantial portions of the Software.
> 
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
> EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
> MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
> IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
> CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
> TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
> SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
