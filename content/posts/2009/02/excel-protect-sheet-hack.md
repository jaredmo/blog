+++
date = "2009-02-07T15:45:17+00:00"
title = "Excel Protect Sheet Hack"
draft = false
tags = ["Excel", "Tech", "VBA"]
+++

I had a situation yesterday where I was trying to get into a password protected Excel sheet and did not know the password. After a quick Google search I realized that the "Protect Sheet" function in Excel has a backdoor which can be exploited with a few lines of VBA. I didn't write the code below, but I am not ashamed to use it. 

The algorithm relies on a backdoor password that can be 1 to 9 characters long where each character is either an "A" or "B" except the last which can be any character from ASCII code 32 to 255. Enjoy and please use responsibly. Of course, if someone is foolish enough to believe Excel is a secure medium for storing confidential information anyway, they deserve what's coming to them. 

```
Sub PasswordBreaker() 
	Dim i As Integer, j As Integer, k As Integer 
	Dim l As Integer, m As Integer, n As Integer 
	Dim i1 As Integer, i2 As Integer, i3 As Integer 
	Dim i4 As Integer, i5 As Integer, i6 As Integer 
	On Error Resume Next 
	For i = 65 To 66: For j = 65 To 66: For k = 65 To 66 
	For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66 
	For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66 
	For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126 
	ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _ 
		Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _ 
		Chr(i4) & Chr(i5) & Chr(i6) & Chr(n) 
	If ActiveSheet.ProtectContents = False Then 
		MsgBox " One usable password is " & Chr(i) & Chr(j) & _ 
			Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _ 
			Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n) 
		Exit Sub 
	End If 
	Next: Next: Next: Next: Next: Next 
	Next: Next: Next: Next: Next: Next 
End Sub
```