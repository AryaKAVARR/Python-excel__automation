# Python-excel__automation

### This Python code uses the `openpyxl` library to process an Excel workbook (`transactions.xlsx`). It performs the following tasks: 

#### 1. Import necessary modules:
   - It imports the `openpyxl` library for working with Excel files.
   - It also imports specific classes required for creating and customizing charts within Excel (`BarChart` and `Reference`).

#### 2. Define a function `process_workbook(filename)`:
   - The function is defined but not used in the provided code.

#### 3. Load the Excel workbook:
   - It loads the Excel workbook named 'transactions.xlsx' into a variable `wb`.
   - It specifies that it wants to work with the first sheet ('Sheet1') in the workbook by assigning it to the `sheet` variable.

#### 4. Iterate through rows in the sheet:
   - It runs a loop to iterate through each row in the 'Sheet1' of the Excel workbook.
   - For each row, it extracts the value in the third column (column C) and multiplies it by 2 to calculate a corrected price.
   - It then writes this corrected price back to the fourth column (column D) of the same row.

#### 5. Create a reference to the data for the chart:
   - It creates a reference object called `values` that specifies the data range for the chart.
   - This range starts from the second row and includes all rows up to the maximum row in column D (the corrected price column).

#### 6. Create a bar chart:
   - It creates an instance of a bar chart using `BarChart()`.

#### 7. Add data to the chart:
   - It adds the data specified by the `values` reference to the bar chart.

#### 8. Add the chart to the sheet:
   - It adds the bar chart to the sheet at cell 'E2' (this is where the top-left corner of the chart will be positioned).

#### 9. Save the modified workbook:
   - Finally, it saves the modified Excel workbook to a file named 'filename'. However, there is an issue in this line, as 'filename' should be replaced with an actual filename, for example, 'output.xlsx', to save the changes to a new file.

In summary, this code loads an Excel workbook, processes the data in the workbook by doubling the values in the third column, creates a bar chart based on the corrected data, and attempts to save the modified workbook with a new filename. Please note that the code has a mistake in the `wb.save('filename')` line, and you should replace 'filename' with an actual filename to save the changes correctly.

### How to run?

### 1.Install the Required Libraries:
Ensure that you have the openpyxl library installed. If you haven't installed it already, you can install it using pip:
```
pip install openpyxl
```
#### 2.Prepare the Excel File:

Make sure you have an Excel file named 'transactions.xlsx'(*you can also rename the file of your choice but for that you will need to update the filename in the python script as well for successful execution.*) in the same directory as your Python script. This Excel file should contain the data you want to process and create a chart from. You can create this file manually in Excel and populate it with your data.

#### 3.Copy and Paste the Code:
Copy the provided Python code and paste it into a Python script file (e.g., process_excel.py).

#### 4.Modify the Save Filename:
In the code, there is a line where it attempts to save the modified workbook with a filename. Replace 'filename' with your desired filename or path (e.g., 'output.xlsx'). Ensure that you keep the file extension as '.xlsx'.

#### 5.Run the Script:
Open a terminal or command prompt and navigate to the directory where your Python script is located. Run the script by typing:

```
./process_excel.py  
```
OR
```
python process_excel.py  
```
OR  
```
python3 process_excel.py (if python3 installed)
```
#### 6.Execution:
The script will execute, opening the 'transactions.xlsx' file, doubling the values in the third column, creating a bar chart based on the corrected data, and attempting to save the modified workbook with the filename you specified.

#### 7.Check the Output:
After running the script, you should find the modified Excel workbook (e.g., 'output.xlsx') in the same directory where your script is located. You can open this file to view the changes and the newly created bar chart.

:fire: :star:Remember to ensure that you have the required permissions to read and write files in the directory where your script is located, and that you have the necessary data in 'transactions.xlsx' for the code to process and chart.





