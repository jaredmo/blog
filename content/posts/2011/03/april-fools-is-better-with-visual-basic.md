+++
date = "2011-03-31T19:49:08+00:00"
title = "April Fools Is Better With Visual Basic"
draft = true
tags = ["Fun", "Tech", "VBA"]
+++

Happy April Fools-Eve, my fellow practical joke connoisseurs. As someone who enjoys a good laugh at another's expense and also happens to be a bit of a geek, I thought would post the VBA code for two of my favorite Excel-based jokes. For those who aren't familiar with Visual Basic for Applications ("VBA"), it's a programming language that Microsoft created which allows users to customize the functionality of MS Office. For example, I have two macros setup which create hotkeys for Paste Values and Paste Formulas, since I use those on a daily basis. VBA is really useful when used correctly. 

Well, like any powerful tool, VBA can used for good and evil. That's one of the reasons Microsoft has warning messages and sirens every time a workbook is opened that contains code. On April 1 every year, I find VBA good for some laughs. Here's two of my favorites. (Full disclosure: The first one I created myself, but be careful as it can cause work to be lost if not used appropriately. The second I took from a VBA forum and modified to suit my needs.) 

#### Automatically Close Excel When A Workbook Is Opened 
This one is simple, but I don't tend to use it unless I plan to tell someone how to fix. Essentially, when a file is opened, the code below will automatically close all instances of Excel and tell the user "Excel is currently experiencing technical difficulties." Fair warning, you'll want to do this when you're sure all instances of Excel are already closed on their machine. If they have anything open, they could lose work and then nobody is laughing. Excel has a standard folder that, when the program is launched, every workbook in that folder is opened and all VBA code is auto-executed. Here's the path: `C:\Documents and Settings\[User name]\Application Data\Microsoft\Excel\XLSTART` 

If you create a new workbook in Excel 2007, go to Developer > Visual Basic, and copy the code below in the "ThisWorkbook" section. Save the file as .xls (2003) and then move it to the path above. Once you are ready for the madness to end, or your coworker is on the verge throwing their computer out the window, just remove your file from their XLSTART folder. 
```
Private Sub 
Workbook_Open() 
MsgBox "Excel is currently experiencing technical difficulties." 
Application.DisplayAlerts = False 
Application.Quit 
End Sub
``` 

#### Excel Is Very Polite Today 
I really like this one. It is harmless, and instantly recognizable as a joke. Only works at the workbook level, so you'll have to hand pick the file you place it in. Essentially, every time a user changes a cell in the target range (A1:AZ10000), a message box appears saying "Thank you. Cell XYZ has been changed." It's awesome. Just follow the same process as above to copy/paste the code into the VBA window of the desired sheet in your workbook. 
```
Private Sub 
Worksheet_Change(ByVal Target As Range) Dim KeyCells As Range 
' The variable KeyCells contains the cells that will 
' cause an alert when they are changed. 
Set KeyCells = Range("A1:AZ10000") 
If Not Application.Intersect(KeyCells, Range(Target.Address)) _ Is Nothing Then 
' Display a message when one of the designated cells has been 
' changed. 
' Place your code here. 
MsgBox "Thank you. Cell " & Target.Address & " has been changed." 
End If 
End Sub
``` 

Hope everyone has a safe and happy April 1. April Fools is always better when you're a geek.