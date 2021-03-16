+++
date = "2012-02-11T21:37:31+00:00"
title = "Excel Pull Function"
draft = true
tags = ["Excel", "Tech"]
+++

I was working on a project this week that required aggregation of data from many, many different Excel workbooks (and various tabs within those workbooks) into one unified dashboard. Now, this can be accomplished with traditional linking, but anyone who has tried to link more than a few workbooks together will tell you this is a pain to maintain and often not very reliable. 

One useful function for creating dynamic links within the same workbook is [INDIRECT](http://office.microsoft.com/en-us/excel-help/indirect-HP005209139.aspx). Instead of hardcoding cell and worksheet references into a formula, INDIRECT allows pieces of a formula to be dynamically updated using variables. Sounds pretty good right? Well, INDIRECT has an Achilles’ heel. It only works if you are linking within the same workbook. Need dynamic links to a closed workbook? Bill Gates says, "You're out of luck." 

Fortunately, after searching around, I found an [old Usenet post from 2004](https://groups.google.com/forum/?fromgroups#!msg/microsoft.public.excel.worksheet.functions/l9ObQ9ku6Bk/_a2jdMD2SeIJ) by a fellow named Harlan Grove. He wrote an Excel function in VB called PULL which essentially fixes INDIRECT's fatal flaw and allows the user to create links to closed workbooks. Just copy the code below into a Module within the workbook of your choice. Then you can use the PULL function just like any other. 

_Example:_   
`A1 = C:\work\`   
`B1 = data\`   
`C1 = \[ClosedBook.xls\]Sheet1'!$A`   
`E1 = PULL("'"&A1&B1&C1)`

```
'----- begin VBA -----
Function pull(xref As String) As Variant
'inspired by Bob Phillips and Laurent Longre
'but written by Harlan Grove
'-----------------------------------------------------------------
'Copyright (c) 2003 Harlan Grove.
'
'This code is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published
'by the Free Software Foundation; either version 2 of the License,
'or (at your option) any later version.
'-----------------------------------------------------------------
'2004-05-28
'fixed the previous fix - replaced all instances of 'expr' with 'xref'
'also now checking for initial single quote in xref, and if found
'advancing past it to get the full pathname [dumb, really dumb!]
'-----------------------------------------------------------------
'2004-03-25
'revised to check if filename in xref exists - if it does, proceed;
'otherwise, return a #REF! error immediately - this avoids Excel
'displaying dialogs when the referenced file doesn't exist
'-----------------------------------------------------------------

Dim xlapp As Object, xlwb As Workbook
Dim b As String, r As Range, C As Range, n As Long

'** begin 2004-05-28 changes **
'** begin 2004-03-25 changes **
n = InStrRev(Len(xref), xref, "\")

If n > 0 Then
If Mid(xref, n, 2) = "\[" Then
b = Left(xref, n)
n = InStr(n + 2, xref, "]") - n - 2
If n > 0 Then b = b & Mid(xref, Len(b) + 2, n)

Else
n = InStrRev(Len(xref), xref, "!")
If n > 0 Then b = Left(xref, n - 1)

End If

'** key 2004-05-28 addition **
If Left(b, 1) = "'" Then b = Mid(b, 2)

On Error Resume Next
If n > 0 Then If Dir(b) = "" Then n = 0
Err.Clear
On Error GoTo 0

End If

If n <= 0 Then
pull = CVErr(xlErrRef)
Exit Function
End If
'** end 2004-03-25 changes **
'** end 2004-05-28 changes **

pull = Evaluate(xref)

If CStr(pull) = CStr(CVErr(xlErrRef)) Then
On Error GoTo CleanUp   'immediate clean-up at this point

Set xlapp = CreateObject("Excel.Application")
Set xlwb = xlapp.Workbooks.Add  'needed by .ExecuteExcel4Macro

On Error Resume Next    'now clean-up can wait

n = InStr(InStr(1, xref, "]") + 1, xref, "!")
b = Mid(xref, 1, n)

Set r = xlwb.Sheets(1).Range(Mid(xref, n + 1))

If r Is Nothing Then
pull = xlapp.ExecuteExcel4Macro(xref)

Else
For Each C In r
C.Value = xlapp.ExecuteExcel4Macro(b & C.Address(1, 1, xlR1C1))
Next C

pull = r.Value

End If

CleanUp:
If Not xlwb Is Nothing Then xlwb.Close 0
If Not xlapp Is Nothing Then xlapp.Quit
Set xlapp = Nothing

End If

End Function
'-----  end VBA  -----
```

#### Update 1/22/20  
The original article had a lot of comments - one from the developer. I retained them below. The Wordpress to Hugo conversion isn't perfect, but I like having them in the post.  

---
#### Comments:
#### 
[Harlan Grove]( "hrlngrv@aol.com") - <time datetime="2012-02-20 17:59:04">Feb 1, 2012</time>

There's a newer version at [http://www.4shared.com/file/L\_eA8s4G/pull.html](http://www.4shared.com/file/L\_eA8s4G/pull.html) I'd make the same caveats I made in my newsgroup postings: the pull udf is SLOW, so you want to use it sparingly. I intentionally made it nonvolatile. If the file would be closed on your computer but someone else has it open and is saving changes, you may think you want it to be volatile, but Excel is a poor tool to use for this sort of data capture. Multiuser databases are much better. If you really, really must make it volatile, use it in formulas like =IF(NOW(),pull(...)) most general but limited =pull(...)+0\*NOW() if you're certain it'll return numbers =pull(...)&T(NOW()) if you're certain it'll return text
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-02-14 15:41:56">Feb 2, 2012</time>

Eric, The issue may lie in Excel not recognizing the leading comma in A14. You also had an additional comma after the filename. Try this and see if it works for you. PULL("'M:\\"&A$14) Cell A14 = SECURITY\\HQ Operations\\Security Operation Center Weekly Report 2012\\Weekly SOC Report xlsx\\\[SOC Weekly Report - 29-Jan-2012.xlsx\]SOC Weekly Report’!C3
<hr />

#### 
[Erik Gaull]( "egaull@worldbank.org") - <time datetime="2012-02-17 09:19:02">Feb 5, 2012</time>

Thanks, Jared!
<hr />

#### 
[Erik Gaull]( "egaull@worldbank.org") - <time datetime="2012-02-13 17:14:12">Feb 1, 2012</time>

Jared -- Thanks for your willingness to help! cell A14 = 'M:\\SECURITY\\HQ Operations\\Security Operation Center Weekly Report 2012\\Weekly SOC Report xlsx\\\[SOC Weekly Report - 29-Jan-2012.xlsx\]'SOC Weekly Report'!C3 cell F18 = =pull(A14) <== this returns a #REF! error Note that all the Excel files are in the same folder. Any help would be most appreciated. Thank you! -- Erik Gaull
<hr />

#### 
[Dave]( "david.breitenbach@gmail.com") - <time datetime="2012-03-27 14:15:51">Mar 2, 2012</time>

Can't get this to work. Does this work with Excel 2007?
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-03-27 16:47:15">Mar 2, 2012</time>

Dave, I'm using it currently with Excel 2007, so it should work for you.
<hr />

#### 
[Graham](http://gravatar.com/grbrewer "grbrewer@gmail.com") - <time datetime="2012-06-22 09:40:31">Jun 5, 2012</time>

This is great, but, in Excel 2010, it has a problem with textual data. All numeric-type data comes over fine, but anything with text, it returns a "#N/A" error. Other than that works beautifully!!!
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-06-22 14:17:42">Jun 5, 2012</time>

Glad it worked!
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-09-22 10:53:59">Sep 6, 2012</time>

Thanks?
<hr />

#### 
[anshuman]( "a2anshuman@gmail.com") - <time datetime="2012-07-28 07:23:47">Jul 6, 2012</time>

Harlan Thanks for creating the pull. However, i need to tweak it slightly. I am using it on an excel where the pull is happening from another file on my system. Now when I email this excel, the pull results in a REF, as the other person does not have the file to pull the data from. Now, vlookup like functions give you an option, to update links, and if you choose not to update links, it stores the values as it had. A solution to my problem can be if pull can have a similar functionality. Another solution can be I create a macro to paste special values of all cells using the pull. I will have to either manually trigger it or something. Any advice, or solutions you may have? Thanks
<hr />

#### 
[Card Transaction Process](http://www.internetmerchantaccountuk.co.uk "mittie.forsythe@gmail.com") - <time datetime="2012-09-22 10:47:45">Sep 6, 2012</time>

Pretty portion of content. I simply stumbled upon your weblog and in accession capital to claim that I get in fact loved account your blog posts. Any way I will be subscribing in your feeds and even I fulfillment you access persistently fast.
<hr />

#### 
[Jerry Delaney](http://gravatar.com/jdelayknee "jdelayknee@yahoo.com") - <time datetime="2012-09-27 08:12:31">Sep 4, 2012</time>

Hi Harlan I develop "one off" tools for a team of analysts to help save them from spreadsheet hell. I have been migrating to excel over the years but I started out in 123. I still have some 123 aps in use and I catch hell from the current MS generation. I have never found an equivalent for what I call "SQL in a cell" (retrieving data from multiple external database files using @dget or @dsum) I also use Lotus Approach RDBM. Do you have any new insights into doing this in Excel without getting into exotic VB programs. Example of a typical ap I have external data files (in db4) for: backlog of orders, finished goods, intransit products, forecast, etc , all of these keyed by Part number and dates User open a 123 spreadsheet template I developed 123 spreadsheet with columns for all the above variables and cell formulas with the @ functions and col to col foemulas. All they have to do: Part numbers are pasted in by user in col A They click one macro button and in 5-10 minutes all the data is retrieved. Another macro copies/paste as values to a 2nd worksheet and they are done. I have goggled this on and off and searched several forums to no avail. When I do post I get the usual dsum, sumif suggestions. I have tried using VB to open and walk through the tables row by row but it takes forever. Any way. If you've spent the time to read all this I really appreciate it. Any help you can give me would be appreciated. Jerry Delaney Apex, NC 9/27/12 jdelayknee@yahoo.com
<hr />

#### 
[Balasubramanian A]( "balasubramanian.eee@gmail.com") - <time datetime="2012-10-07 07:16:41">Oct 0, 2012</time>

I keep getting a #name error...
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-10-07 08:56:15">Oct 0, 2012</time>

Can you put your formula in the comments? I can take a look at it.
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2012-11-02 07:02:03">Nov 5, 2012</time>

I'll have to check it out. Thanks!
<hr />

#### 
[The Horse]( "strangehorse@eircom.net") - <time datetime="2012-11-02 04:06:51">Nov 5, 2012</time>

Guys - none of you has, obviously, heard of Laurent Longre's Morefunc add-in, which includes the INDIRECT.EXT function. This modified function enables you to create volatile links to closed workbooks. I've used it extensively.
<hr />

#### 
[paradroid]( "paradroid001@gmail.com") - <time datetime="2012-11-12 10:14:47">Nov 1, 2012</time>

Rock'n'roll! It works perfectly. In case somebody is as dumb as I am: it will not work when you copy =PULL(“‘”&A1&B1&C1) from the website right into your spreadsheet. Because Excel does not know what these funny typographic quotation marks are, you would need to use =PULL("'"&A1&B1&C1). Took me 10 minutes to find out what the propblem was :-(, - now it works perfectly :-) thanks!!
<hr />

#### 
[paradroid]( "paradroid001@gmail.com") - <time datetime="2012-11-12 10:17:04">Nov 1, 2012</time>

Ha! Seems like this website automatically assigns the typographic quotation marks. What I wanted to say: do not copy this part “‘” but rather type it in Excel!
<hr />

#### 
[Jason]( "jsnschober@gmail.com") - <time datetime="2013-01-25 10:01:26">Jan 5, 2013</time>

Thanks for the function, saved me some headaches! One question...is there a way to get this to pull data from a .csv as well/instead?
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2013-01-25 12:29:32">Jan 5, 2013</time>

Thanks! I don't think so...You may want to ask Harlan. He commented on this post.
<hr />

#### 
[Lars]( "brunner.lars@gmail.com") - <time datetime="2013-02-04 05:37:35">Feb 1, 2013</time>

Hi there. I'm currently trying to get the pull function to work as well. First thanks for the great work already accomplished! I experienced the first chunk of code to work as it is posted on this blog when simply copying it as a new module in the excel - macro enabled- file. The second version of it, the .xla file will not work for me though, the "pull" function calls are replaced by #NAME? instead. The macro is enabled though when I look at File -> Options -> Complements, "HG's own add-in". I have the feeling, Excel does not recognize the function correctly. Autocompletion works though. Any suggestion ?
<hr />

#### 
[Cristina]( "cristinajimgut@gmail.com") - <time datetime="2013-02-06 18:35:20">Feb 3, 2013</time>

Hi, Could someone please help me figure out what I'm doing wrong? Here's what I have: Cell B1= 'C:\\Users\\Cristina\\Documents\\My Dropbox\\Cristina\\Evaluacion Automatica de Pruebas\\Evaluacion 6 Necesidades\\Participantes\\\[Encuesta Necesidades S1P01.xlsx\]Sheet1'!$D9 =PULL("'"&B$1) I already tried the suggestion given above about doing the following and it is not working (except I cannot make excel not put a coma at the beginning of Cell B1 - I erase it and it keeps putting it back). So it looks like this: Cell B1= 'Users\\Cristina\\Documents\\My Dropbox\\Cristina\\Evaluacion Automatica de Pruebas\\Evaluacion 6 Necesidades\\Participantes\\\[Encuesta Necesidades S1P01.xlsx\]Sheet1'!$D9 =PULL("'C:\\"&B$1) Still not working.. Please help. Maybe the formula is not working well? how do I make sure it is? Thank you very very much! Cristina
<hr />

#### 
[Read data from file without opening them?](http://www.mrexcel.com/forum/excel-questions/686324-read-data-file-without-opening-them.html#post3396278 "") - <time datetime="2013-02-18 09:31:49">Feb 1, 2013</time>

\[...\] You might want to check out Harlan Grove's PULL function Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[...\]
<hr />

#### 
[Chris Burr]( "chrisburr@brpuk.com") - <time datetime="2013-03-18 17:38:53">Mar 1, 2013</time>

DOES IT WORK ON MAC EXCEL 2011?
<hr />

#### 
[Pappu Pager](http://none "pappu.pager2020@gmail.com") - <time datetime="2013-05-22 03:18:35">May 3, 2013</time>

Try this buddy. 'Requires filename, sheetname as first argument and cell reference as second argument 'Usage: type in an excel cell -> =getvalue(A1,B1) 'Example of A1 -> C:\\TEMP\\\[FILE1.XLS\]SHEET1' 'Example of B1 -> B3 'This will fetch contents of cell (B3) located in (sheet1) of (c:\\temp\\file1.xls) 'Create a module and paste the code into the module (e.g. Module1, Module2) Public xlapp As Object Private Function getvalue(ByVal filename As String, ref As String) As Variant ' Retrieves a value from a closed workbook Dim arg As String Dim path As String Dim file As String Dim count As Integer count = 0 While (Left(Right(filename, count), 1) "\\") count = count + 1 Wend count = count - 1 path = Left(filename, Len(filename) - count) count = count - 1 file = Right(filename, count) count = 0 While (Left(Right(file, count), 1) "\]") count = count + 1 Wend file = Left(file, Len(file) - count) If Dir(path & file) = "" Then getvalue = "File Not Found" Exit Function End If 'Set xlapp = CreateObject("Excel.application") - Object must be created only once and not at each function call. Do not enable ' Create the argument arg = "'" & filename & "!" & Range(ref).Range("A1").Address(, , xlR1C1) ' Execute an XLM macro getvalue = xlapp.ExecuteExcel4Macro(arg) End Function 'Module code ends here. Following code goes into Workbook\_Open event 'This code goes into Workbook open event within THISWORKBOOK Private Sub Workbook\_Open() Set xlapp = CreateObject("Excel.application") End Sub
<hr />

#### 
[Brad]( "brainman00x@yahoo.com") - <time datetime="2013-06-25 09:17:49">Jun 2, 2013</time>

I to am having trouble with this function in the same lines that Phil mentioned: While (Left(Right(filename, count), 1) “\\”) While (Left(Right(file, count), 1) “\]”) Any help would be appreciated
<hr />

#### 
[Viren]( "vtellis@appnexus.com") - <time datetime="2013-06-25 12:07:13">Jun 2, 2013</time>

While (Left(Right(filename, count), 1) “\\”) While (Left(Right(file, count), 1) “\]”) There needs to be equal signs here: While (Left(Right(filename, count), 1) = “\\”) While (Left(Right(file, count), 1) = “\]”) I'm having trouble with the part in ThisWorkBook Private Sub Workbook\_Open() Set xlapp = CreateObject(“Excel.application”) End Sub I'm getting an error when I run the code. Any help?
<hr />

#### 
[Viren]( "vtellis@appnexus.com") - <time datetime="2013-06-25 12:22:43">Jun 2, 2013</time>

specifically, I get an error saying that an object is required each time I open the workbook. I'm referring to pappu's code
<hr />

#### 
[Viren]( "vtellis@appnexus.com") - <time datetime="2013-06-25 18:54:44">Jun 2, 2013</time>

one last comment. the error is a runtime error 424
<hr />

#### 
[VBA Code to look at a range in a different workbook](http://www.mrexcel.com/forum/excel-questions/729689-visual-basic-applications-code-look-range-different-workbook.html#post3589189 "") - <time datetime="2013-09-29 06:10:49">Sep 0, 2013</time>

\[…\] I think that you willneed something like Harlan Grove's Pull function Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[…\]
<hr />

#### 
[Brad]( "brainman00x@yahoo.com") - <time datetime="2013-06-26 09:20:23">Jun 3, 2013</time>

Thanks Viren for pointing out the missing equal signs. Thant has resolved the compile errors I was getting. However, I am running into the same runtime error (424) that you are running into with pappu's code. Error occurs upon opeing the workbook.
<hr />

#### 
[](http://www.mrexcel.com/forum/excel-questions/699000-using-variable-filename-%93if-then-else%94-formula-excel-2003-a.html#post3454304 "") - <time datetime="2013-04-24 09:19:57">Apr 3, 2013</time>

\[...\] take a look at this: Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[...\]
<hr />

#### 
[Steve]( "stephenewski@aol.com") - <time datetime="2013-07-17 13:37:27">Jul 3, 2013</time>

I was able to make the pull function work, but is there any way to make the cell reference variable? Like if I put it in cell A1, can I drag the formula down to pull B1, C1, etc?
<hr />

#### 
[Steve]( "stephenewski@aol.com") - <time datetime="2013-07-17 14:19:21">Jul 3, 2013</time>

ha, nevermind. I got it.
<hr />

#### 
[Alex Frazier]( "alex.frazier@charter.net") - <time datetime="2014-01-22 09:29:32">Jan 3, 2014</time>

I found a much easier way to do this. The goal is to access data from closed workbooks using dynamic links, whereas INDIRECT won't work on a closed workbook. But you can still use dynamic links. You just have to go through the Name Manager and create a universal reference name. In the workbook I created, I use the drop down box cell link value to reference an array and piece together the cell reference I want. From there, what people are typically wont to do is use that constructed cell reference with INDIRECT in their formulas (× who knows how many formulas). Instead, I used a vba sub routine to rewrite the value of a name in the Name Manager, and I used the name as a variable in my formulas. So I might have in one of my cells the formula: =nV(VLOOKUP($E9,SN,3)) E9 is the date code reference, which I use as a key on my data table, SN is the variable named range, and 3 is the column in that range. (nV is a custom function I built that returns either a blank cell or the value you choose if the criteria equals "", 0, or an Error). The name SN in the Name Manager has the value: ='https://d.docs.live.net/ed2bad18b7a35798/Public/\[DT6438.xlsm\]Data Table'!$A$5:$DH$35 (I altered some numbers for security). When I change the drop down box to another selection, it reconstructs the reference with the correct file name according to the array, and the correct range of cells according to the specified month on the sheet. Using the array, the code looks like this: =nV(CONCATENATE("'https://d.docs.live.net/ed2bad18b7a35798/Public/\[",VLOOKUP(CJ1,CJ2:CL24,3),".xlsm\]Data Table'!","$A$",(DATE(VLOOKUP(D38,C1:D37,2),B13,1)-41635),":$DH$",EOMONTH(DATE(VLOOKUP(D38,C1:D37,2),B13,1),0)-41635)) Since the source files are all in the same drive, they all share https://d.docs.live.net/ed2bad18b7a35798/Public/ in common. This could be individually specified if necessary, if the files were in different drives. The VLOOKUP is checking the array. CJ1 is the drop down box cell reference. CJ2:CL24 is the array range. 3 is the column in the array where the file name is found. The specific cell reference range is built based on the date entered on the main sheet drop down box. It gives a range from the first day of the month to the last day of the month, using date codes minus the correct amount to equal the actual row number for the cell reference. The date code for January 1st, 2014, for example, is 41640. That date is row 5 in the data table. Ergo, DATE(2014,1,1)-41635, or 41640 minus 41635. The A and DH column parameters are just the full width of my data range on the data table. Once the reference is built, a vba Sub generates a variable out of the constructed reference and overwrites the value of SN in the Name Manager with this code: Sub Change\_Reference() New\_Ref = Sheets("Data Sheet").Range("CJ26") ActiveWorkbook.Names("SN").RefersTo = "=" & New\_Ref End Sub I have 992 cells using the SN reference, pulling over 100 pieces of data from 20 separate stores on a remote drive. The file size is currently 201 KB. It works. It works beautifully. It works quickly. And it's memory efficient. Relative to the requests for help I've seen across the web, not merely on the INDIRECT function, but INDIRECT in this specific application, I hope this is helpful to someone.
<hr />

#### 
[Calling a cell that contains part of a filename](http://www.mrexcel.com/forum/excel-questions/752930-calling-cell-contains-part-filename.html#post3696784 "") - <time datetime="2014-01-24 21:41:01">Jan 5, 2014</time>

\[…\] the file is closed, use a custom function called PULL Excel “Pull” Function: Creating dynamic links to closed workbooks If you do a web search for something like Excel Pull function, you'll find several threads on this \[…\]
<hr />

#### 
[EF]( "eiji.fujii@zurich.com") - <time datetime="2013-08-22 05:21:31">Aug 4, 2013</time>

Pappu, can you send me your files. The calculations are taking forever in my file. Thanks! EF djfobster@hotmail.com
<hr />

#### 
[Roman]( "romanstankiewicz@hotmail.com") - <time datetime="2013-09-12 04:23:23">Sep 4, 2013</time>

Im trying to use the pull function with a range- =pull($AM$46&AL13) AL13 is B10:Z10 AM46 is ''P:\\Coater\\RAS\\\[B-Shift holiday planner 2013.xls\]September'! However it does not recognise that AL13 is a range and only returns the first value of this ie. B10. Can anyone help me understand why this is? Is there a way to make it pull the whole range like with indirect?
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2014-01-22 11:25:20">Jan 3, 2014</time>

Thanks for the comment! I'll have to give that a try.
<hr />

#### 
[Alex Frazier]( "alex.frazier@charter.net") - <time datetime="2014-01-22 11:34:04">Jan 3, 2014</time>

With all the people that have tried to figure this out, it would have been wrong of me not to share such a simple solution now that I've figured out how to make it work. Let me know how it works out for you.
<hr />

#### 
[hsn]( "cemil0225@gmail.com") - <time datetime="2013-04-12 01:14:18">Apr 5, 2013</time>

Thanks for your function, it is really good..
<hr />

#### 
[Clayton]( "stoffelupegus@gmail.com") - <time datetime="2014-09-05 14:30:25">Sep 5, 2014</time>

Thanks Alex, with that I was able to figure out some of it. The thing I still can't figure out is how you handle multiple workbooks. Do you have a different "Name" for each cell\*workbook value? For example, if I have 3 workbooks each with 20 values to be imported, do I have to have 60 "Names" (workbook1cell1, workbook2cell1, workbook3cell1, workbook1cell2, etc.?) The thing I am trying to accomplish is essentially pull in 100 cells worth of data each from a series of workbooks that are all set up exactly the same just by changing the cell that contains the reference to the workbook.
<hr />

#### 
[Jhon](http://Doe "wnez.jhon@9ox.net") - <time datetime="2013-10-11 17:51:22">Oct 5, 2013</time>

The code didn't work for me, but forced me to look for more options. Using the code I found in http://stackoverflow.com/questions/9259862/executeexcel4macro-to-get-value-from-closed-workbook written by Siddharth Rout I created the following version of that function, tested in Excel 2007, removing the MsgBox: It is almost identical to the code above, just going to the point: getting the data. Function indirext(wbPath As String, wbName As String, wsName As String, cellRef As String) Dim Ret As String Ret = "'" & wbPath & "\[" & wbName & "\]" & \_ wsName & "'!" & Range(cellRef).Address(True, True, -4150) indirext = ThisWorkbook.xlapp.ExecuteExcel4Macro(Ret) End Function This requires two definitions in ThisWorkBook: Public xlapp3 As Object Private Sub Workbook\_Open() Set xlapp = CreateObject("Excel.application") End Sub I did that following the idea indicated by Papu Pagger Now you can test the function. example: indirext("C:\\e\\","test.xlsx","sheet1","a1") Make sure all parameters are strings My hat off for Harlan Grove. I thought this was not possible. Enjoy!
<hr />

#### 
[Jhon Doe]( "wnez.jhon@9ox.net") - <time datetime="2013-10-11 17:53:26">Oct 5, 2013</time>

I just noticed a typo. Where it says Public xlapp3 As Object it should say Public xlapp As Object
<hr />

#### 
[Kenny]( "kenny_tm_leong@yahoo.com") - <time datetime="2013-11-18 15:02:49">Nov 1, 2013</time>

It's great that there is a function like this user defined pull function. Although, while user defined functions are working in my excel 2003 on a winxp machine, I haven't been able to get the pull function UDF to work. I opened a new excel spreadsheet (a .xls file) and opened up the VBE, chose 'insert module', then copied the pull function UDF (from the original post of this thread) to 'module 1'. Then exited VBE. Then I did a test, by putting the number '4' in cell E1, and the number '3' in cell E4. I then firstly put =INDIRECT("E"&E1) into cell A1, which resulted in a '3' in cell A1. Then I put =pull("E"&E1) into cell A2, which resulted in #REF! in cell A2. I actually want to use the pull function for external worksheets, but I first needed to see if the pull function works within its own local spreadsheet to start with. Other user defined functions are working on my system, and I'm trying to get this pull function to work (and it has been correctly copied to 'module 1'). I would love to get this one to work because this would be a very nice feature to have. Thanks in advance!
<hr />

#### 
[Variables within File Name References](http://www.mrexcel.com/forum/excel-questions/733110-variables-within-file-name-references.html#post3606623 "") - <time datetime="2013-10-17 08:01:29">Oct 4, 2013</time>

\[…\] Maybe Harlan Grove's PULL function Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[…\]
<hr />

#### 
[Referencing another workbook, where the name changes...INDIRECT?](http://www.mrexcel.com/forum/excel-questions/724976-referencing-another-workbook-where-name-changes-indirect.html#post3567682 "") - <time datetime="2013-09-06 04:22:18">Sep 5, 2013</time>

\[…\] INDIRECT will not work with closed workbooks. You could try Harlan Grove's PULL function Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[…\]
<hr />

#### 
[Pappu Pager](http://127.0.0.1 "pappu.pager2020@gmail.com") - <time datetime="2013-05-01 08:37:16">May 3, 2013</time>

Hi harlan grove, This function is slow for only one reason. The Set xlapp = CreateObject("Excel.Application") has been defined in the function, which creates a new object every time a cell is changed. We need to initialize it only once. So you can just declare - - public xlapp as object - - in the module just above the line - - Function pull(xref As String) As Variant and PUT THE - - Set xlapp = CreateObject("Excel.Application") within workbook\_open event under THISWORKBOOK. Believe me it works awesome. I have tried it and pulled 100 rows each from 20 files in 15 seconds. I can send you my excel files which are working with awesome speed. Please give me your email id. Pappu Pager. (From India that is Bharat)
<hr />

#### 
[File name references](http://www.mrexcel.com/forum/excel-questions/827054-file-name-references.html#post4035483 "") - <time datetime="2015-01-05 09:37:09">Jan 1, 2015</time>

\[…\] document.write(''); Excel doesn't create an external link from a concatenated string. There is a custom User Defined function (UDF) called PULL that will link to a close file from a concatenated string. Excel PULL Function \[…\]
<hr />

#### 
[Clayton]( "stoffelupegus@gmail.com") - <time datetime="2014-08-28 09:59:35">Aug 4, 2014</time>

Alex, Would you be able to share a copy of an excel workbook with this functioning in it, it sounds exactly like what i'm looking for but I am having time visualizing what exactly you did.
<hr />

#### 
[Alexander Frazier](https://www.facebook.com/alexander.frazier.90 "alex.frazier@charter.net") - <time datetime="2014-09-01 14:55:45">Sep 1, 2014</time>

Clayton, I apologize, but the only examples I have are company spreadsheets. I'm not at liberty to share them with the public. But I'll do the best I can to be less technical and explain this a little better. If you were to piece together a cell reference for use in an INDIRECT command, you might have "A" in A1 and "27" in B1, so that you can build the cell reference "A27" in C1 with a formula like =A1 & B1. Then in D1 you could use the formula =INDIRECT(C1), which would return, not the value A27, but the value OF A27. The values of "A" in A1 and "27" in B1 may change depending on other formulas you have, so that INDIRECT(C1) might return the value of A27, or D32, or whatever. That's the power of the INDIRECT command. What most people are trying to do is use it to access closed workbooks or workbooks on remote servers, which INDIRECT simply will not do. What you CAN do is use a file or web path, with workbook and cell reference directly in your formula to retrieve the information you want. But this is not variable. If you have something set up to manage multiple stores, like I do, then you might want the same value from the same cell, but across ten or twelve workbooks. That means you have to either change your formulas, write a lengthy, complicated script that rewrites your formulas, or you have to find a way to make INDIRECT work on remote workbooks ... which, again, just can't be done. So what I did was use a Name in my formulas. In the Name Manager, you can create a Name that stands in place of a value, formula, or whatever you need it to stand in place of. The value of the Name Manager is that if you have a repeated, lengthy formula in numerous cells, you can clear up the congestion by creating a Name that IS the formula. And as fate would have it, a Name that stands in place of a file or web path, complete with workbook and cell reference, works just fine in a formula. As fate would also have it, a simple vb script can rewrite that single Name reference. So what you have to do is create a list of file or web paths and build your cell reference, just as in the example of building a cell reference for INDIRECT. Then have vb take the value of that cell (vb doesn't distinguish between a constructed reference and a bonafide reference) and have it rewrite what the Name you've created refers to. Like so: Sub Change\_Reference() New\_Ref = Sheets(“Sheet”).Range(“Cell″) ActiveWorkbook.Names(“Name”).RefersTo = “=” & New\_Ref End Sub New\_Ref takes its value from the sheet and cell you specify. Then the Name you specify in the active workbook you're using is made to refer to the new value you've specified for New\_Ref. It's really quite simple once you see it work.
<hr />

#### 
[Pull vba from Harlan Grove - Need Code](http://www.mrexcel.com/forum/excel-questions/808749-pull-visual-basic-applications-harlan-grove-need-code.html#post3953235 "") - <time datetime="2014-09-30 10:07:03">Sep 2, 2014</time>

\[…\] document.write(''); The vba will pull information from a closed workbook instead of requiring that it be open. I am using the function to link more than 30 workbooks for a weekly management report. A link I found was: Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[…\]
<hr />

#### 
[Jan Pavuk](http://www.janpavuk.info "jan.j.pavuk@gmail.com") - <time datetime="2013-07-23 03:04:23">Jul 2, 2013</time>

Hi Graham. I think this is the core issue and need at the same time. You can pull numeric data also with a simple =SUM formula (e.g. =SUM(\[1\_Table\_Mapping\_Overview.xlsx\]Table\_Overview!F182)) Have you managed to find a solution for text data?
<hr />

#### 
[reference to a particular cell in a closed workbook (indirect)](http://www.mrexcel.com/forum/excel-questions/700501-reference-particular-cell-closed-workbook-indirect.html#post3463092 "") - <time datetime="2013-05-04 07:57:30">May 6, 2013</time>

\[...\] INDIRECT does not work with closed workbooks. Try Harlan Groves' PULL function Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[...\]
<hr />

#### 
[Rob]( "robert_foppiani@timeinc.com") - <time datetime="2013-05-20 17:56:20">May 1, 2013</time>

Pappu - can you post the changes in code please? Thanks
<hr />

#### 
[aleksandra]( "a.wysoczanska@gmail.com") - <time datetime="2013-07-09 09:20:23">Jul 2, 2013</time>

I am trying to use that function to get values from csv, should it work as well? (It doesn't, but that may be as I made some other mistake)
<hr />

#### 
[Ashley McCash]( "ashley.mccash@gmail.com") - <time datetime="2013-07-28 15:32:13">Jul 0, 2013</time>

Pappu - I use this function but due to the data volume and amount of files it takes forever! I would love a copy of your excel file. Tried your code but it didn't work - probably user error!
<hr />

#### 
[Pappu Pager](http://none "pappu.pager2020@gmail.com") - <time datetime="2013-05-22 03:13:35">May 3, 2013</time>

Hello Guys, Sorry for the late reply. Pasted below the code I am using. You will get a fair idea what I want to say. I just initialized the object once at workbook\_open event. And for doing so I declared it as global variable. 'Requires filename, sheetname as first argument and cell reference as second argument 'Usage: type in an excel cell -> =getvalue(A1,B1) 'Example of A1 -> C:\\TEMP\\\[FILE1.XLS\]SHEET1' 'Example of B1 -> B3 'This will fetch contents of cell (B3) located in (sheet1) of (c:\\temp\\file1.xls) 'Create a module and paste the code into the module (e.g. Module1, Module2) Public xlapp As Object Private Function getvalue(ByVal filename As String, ref As String) As Variant ' Retrieves a value from a closed workbook Dim arg As String Dim path As String Dim file As String Dim count As Integer count = 0 While (Left(Right(filename, count), 1) "\\") count = count + 1 Wend count = count - 1 path = Left(filename, Len(filename) - count) count = count - 1 file = Right(filename, count) count = 0 While (Left(Right(file, count), 1) "\]") count = count + 1 Wend file = Left(file, Len(file) - count) If Dir(path & file) = "" Then getvalue = "File Not Found" Exit Function End If 'Set xlapp = CreateObject("Excel.application") - Object must be created only once and not at each function call. Do not enable ' Create the argument arg = "'" & filename & "!" & Range(ref).Range("A1").Address(, , xlR1C1) ' Execute an XLM macro getvalue = xlapp.ExecuteExcel4Macro(arg) End Function 'Module code ends here. Following code goes into Workbook\_Open event 'This code goes into Workbook open event within THISWORKBOOK Private Sub Workbook\_Open() Set xlapp = CreateObject("Excel.application") End Sub
<hr />

#### 
[Alexander Frazier](https://www.facebook.com/alexander.frazier.90 "alex.frazier@charter.net") - <time datetime="2014-12-24 06:49:00">Dec 3, 2014</time>

If you want the 100 cells worth of values all on the same page at the same time, then each workbook you are referencing would need its own name. The purpose of the name is to allow variables. If you change the value the name represents, it changes the data you retrieve. If you have 100 specific cells you want to pull information from in a specific sheet, you can access that data directly. If the 100 cells you pull data from change within that sheet due to variables, you need to use a name so you can dynamically alter the cells you are retrieving data from, since INDIRECT won't work on closed workbooks stored on a remote drive. If you are pulling up data from one workbook at a time, then you only need a single name. When you rewrite the name's reference value, you can change both the workbook it refers to and the cells it retrieves at the same time. Just think of the indirect command and how it would be done if the command worked the way we want it to. You would piece together the drive, file name, etc., and you would piece together the cell reference. Then you would put them together using the INDIRECT command to access the data you want. For this method, build the reference the very same way. Then use vb to rewrite the reference value of the NAME to the fully constructed reference you've created.
<hr />

#### 
[Sum Data from Closed Work Sheets](http://www.mrexcel.com/forum/excel-questions/695665-sum-data-closed-work-sheets.html#post3439158 "") - <time datetime="2013-04-06 15:10:00">Apr 6, 2013</time>

\[...\] Checkout Harlan Grove's custom user defined Pull function \[...\]
<hr />

#### 
[Danesh]( "daneshshamsi18@gmail.com") - <time datetime="2015-02-09 08:16:54">Feb 1, 2015</time>

Pappu, I know that this post is years old, but I would really appreciate it if you could further explain how to get your alterations working in this code. I simply don't understand how it works in conjunction with the pull code. Thanks!
<hr />

#### 
[Phil](http://U "psu@btinternet.com") - <time datetime="2013-05-30 09:00:50">May 4, 2013</time>

Hi I cant get the function to run. There are 2 lines that are giving compile errors. While (Left(Right(filename, count), 1) “\\”) While (Left(Right(file, count), 1) “\]”) Would someone be able to help?
<hr />

#### 
[J]( "james.reyenga@gmail.com") - <time datetime="2013-05-15 17:07:23">May 3, 2013</time>

Pappu, are there any other details? I can't seem to get the pull function to work with your proposed changes, and at the moment it is might slow... Thanks
<hr />

#### 
[Trygve Wastvedt](http://www.trygvewastvedt.com "trygve.wastvedt@gmail.com") - <time datetime="2015-09-30 10:19:09">Sep 3, 2015</time>

I had the same issue. When I copy/pasted I got different characters for double quotes, single quotes, and minus signs than what I get when I type those on my keyboard. I ran a search and replace for those three characters and now it works for me.
<hr />

#### 
[Trygve Wastvedt](http://twastvedt.wordpress.com "trygve.wastvedt@gmail.com") - <time datetime="2015-09-30 10:22:46">Sep 3, 2015</time>

I had the same issue. I had to run a search and replace to convert the double quotes, single quotes, and minus signs to the characters I get from my keyboard. When I copy/pasted from the post above I got different symbols than the standard for those three.
<hr />

#### 
[Anderson Santos]( "andsantos@trombini.com.br") - <time datetime="2014-12-05 11:48:21">Dec 5, 2014</time>

Hello Guy, I tried to use de pull function but I'm had some problems. When the another workbook was into "C:" the function worked, but I have workbooks in a server, (H:, W:, F:) and when I used these workbooks that were into W: for exemple, the function didn't work. The result was an error #REF. What can I do? Thanks.
<hr />

#### 
[peter Bedson]( "peterzbedson@gmail.com") - <time datetime="2015-03-12 04:25:44">Mar 4, 2015</time>

This seems to be a very elegant solution - you can use an automatic procedure on Worksheet Change to update the name reference if the control cells which build the reference change - that would let you have multiple different names driven off changes in the same WS. I knew that names don't have to refer to a cell or to a range but can refer to anything including text or entire file paths (which I have used in the past) but had not made the jump to have them dynamically update like this. You learn something new every day...
<hr />

#### 
[Keith]( "kgwarburton@yahoo.co.uk") - <time datetime="2017-06-16 03:29:23">Jun 5, 2017</time>

I receive a #VALUE! error when I run this code. Any ideas why that would be?
<hr />

#### 
[Rex Barker]( "mysterious_mr_bill@hotmail.com") - <time datetime="2014-12-17 08:03:44">Dec 3, 2014</time>

Simplified, yet again. Here, there is no need to add the ThisWorkbook\_Open() method; the function checks if the object exists and creates one if it doesn't. Paste this into a new VBA module in your Excel workbook. Watch for character translations that may have occurred due to pasting in the website: 'Requires filename, sheetname as first argument and cell reference as second argument 'Usage: type in an excel cell -> =getvalue(A1,B1) 'Example of A1 -> C:\\TEMP\\\[FILE1.XLS\]SHEET1' 'Example of B1 -> B3 'This will fetch contents of cell (B3) located in (sheet1) of (c:\\temp\\file1.xls) 'Create a module and paste the code into the module (e.g. Module1, Module2) Public xlapp As Object Private Function getvalue(ByVal filename As String, ref As String) As Variant ' Retrieves a value from a closed workbook Dim arg As String Dim path As String Dim file As String Dim count As Integer filename = Trim(filename) path = Mid(filename, 1, InStrRev(filename, "\\")) file = Mid(filename, InStr(1, filename, "\[") + 1, InStr(1, filename, "\]") - InStr(1, filename, "\[") - 1) If Dir(path & file) = "" Then getvalue = "File Not Found" Exit Function End If If xlapp Is Nothing Then 'Object must be created only once and not at each function call. Do not enable Set xlapp = CreateObject("Excel.application") End If ' Create the argument arg = "'" & filename & "'!" & Range(ref).Range("A1").Address(, , xlR1C1) 'Execute an XLM macro getvalue = xlapp.ExecuteExcel4Macro(arg) End Function
<hr />

#### 
[Jerm]( "pinkpeace7@hotmail.com") - <time datetime="2015-07-02 20:42:41">Jul 4, 2015</time>

Hi! Will anyone know if the PULL function would works for dynamic name range reference to a closed workbook? The name range is for example '\[Data.xlsx\]Leads'!$B$5:$B$23742 and the row number 23742 would change according to the last row of data in the closed workbook hence I need it to be dynamic and still able to use it within Sumproduct formula. Any help would be great!!
<hr />

#### 
[Can a spreadsheet update a spreadsheet that is not open](http://www.mrexcel.com/forum/excel-questions/855045-can-spreadsheet-update-spreadsheet-not-open.html#post4157742 "") - <time datetime="2015-05-14 07:59:59">May 4, 2015</time>

\[…\] document.write(''); have read here Excel “Pull” Function: Creating dynamic links to closed workbooks | Numbermonger \[…\]
<hr />

#### 
[Jay]( "ja_soccer@hotmail.com") - <time datetime="2015-06-18 12:59:31">Jun 4, 2015</time>

Rex, your final function is awesome. Easy to use (other than the many copy-paste issues) and absurdly fast compared to Harlan's pull function. Just wanted to say thanks, it saved me on a project that would have been impractically slow using the other method.
<hr />

#### 
[Norman]( "norman_conquest@hotmail.com") - <time datetime="2015-09-17 08:30:41">Sep 4, 2015</time>

This line is giving me an error: file = Mid(filename, InStr(1, filename, "\[") + 1, InStr(1, filename, "\]") – InStr(1, filename, "\[") – 1) Any ideas why?
<hr />

#### 
[Numbermonger]( "jared@monger.cc") - <time datetime="2015-11-24 16:09:02">Nov 2, 2015</time>

Rex, Would you be able to post and link to a .txt file with this function? Having difficulty and I think it likely has to do with copying/pasting from the wordpress comment.
<hr />
