+++
date = "2010-03-18T13:47:53+00:00"
title = "Access VBA: Starting from Ground Zero"
draft = false
tags = ["Access", "Microsoft", "Tech", "VBA"]
+++

Recently, I have been getting out my comfort zone with SAS/SQL and expanding into the world of Visual Basic. For those unfamiliar, Visual Basic for Applications (or VBA) is Microsoft’s language which allows users to customize Microsoft Office. It is especially useful in Access and Excel. A simple example of VBA would be an Excel macro. When a user records a macro, the steps are recorded as VBA code and mapped to a keyboard shortcut. For instance, I have a macro setup so that when I press Ctrl + M, Visual Basic is invoked to perform the Paste Values function in Excel. This turns three clicks of the mouse into one easy keystroke. 

The primary reason I started dabbling in VBA was because of Access. There were some processes in my current work that needed a more robust solution than spreadsheets and manual record keeping. The most cost effective solution I could come up with was to create an Access database with custom forms for data entry. I was pleasantly surprised that Access VBA is not as complicated as I had originally thought. Here are some examples of useful functions. 

#### Requerying data in forms 
A good use of forms in Access is to query records based on combo or list box selections and return only the relevant results. In the example code below, the records returned are based on three key fields: 1) Team 2) Fiscal Year and 3) Product. This VBA subroutine is invoked when the first combo box `(cmb\_TeamDrop)` is changed. The code then requeries the values in the Fiscal Year combo box `(cmb\_YearDrop)` and Product list box `(lst\_Product)`. Then finally it requeries the results of the master form itself `(Me.Requery)`. This causes the form to return only the records that match the three selections instead of the entire database. 
```
Private Sub 
cmb_TeamDrop_Change() 
Me.cmb_YearDrop.Requery 
Me.cmb_YearDrop.Value = "" 
Me.lst_Product.Requery 
Me.Requery 
End Sub 
```

#### Locking subform data 
One of the needs I had in this particular project was to lock data for fiscal years that had already passed. I’m a big fan of locking down data after a reporting period has ended. You never want an overly ambitious staff “correcting” last year’s information without your knowledge. Reports don’t match and much work always ensues. Fortunately, there is a VBA parameter called “Locked” which can be toggled on and off based on a field’s value using an IF statement. In this particular example, the code is invoked by making a selection in a list box on the master form. The code then examines a Yes/No field on the Fiscal Year table called “Locked”. If that value is yes, then all subforms are subsequently locked. As long as the raw tables themselves are locked down, staff cannot alter any prior year data. 
```
Private Sub
lst_Product _Click() 
Me.Locked.Requery 
Me.Requery 
If Me.Locked = -1 Then Me.[subform].Locked = True Else Me.[subform].Locked = False End If 
End Sub 
```

That said. If you, like me, are comfortable in querying languages such as SAS or SQL, but wary of true object oriented programming, don’t let it scare you. The same skills that allowed you to learn one language give you an edge in learning another.