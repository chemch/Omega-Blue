VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form3 
   Caption         =   "Confirmation"
   ClientHeight    =   5028
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4908
   OleObjectBlob   =   "Form3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
'When to form first opens the following two subroutines will be called
delCode
'delCode will go into the users active worksheet module and remove the code that was added before.

zirnCore
'zirnCore is the heart of this program. It functions by using "white space technology," which is a logical design that uses the spacial
    'organization of white space to understand relative locations of cells and their values. It functions by navigating and understanding
    'the layout of a worksheet not by the values in the sheet but by their absence.

End Sub

Private Sub CAcq_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T10.Enabled = True
F3T10.Value = ""
F3T10.SetFocus
End Sub

Private Sub CAcqT_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T9.Enabled = True
F3T9.Value = ""
F3T9.SetFocus
End Sub

Private Sub CCase_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T0.Enabled = True
F3T0.Value = ""
F3T0.SetFocus
End Sub

Private Sub CChan_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T6.Enabled = True
F3T6.Value = ""
F3T6.SetFocus
End Sub

Private Sub CCoverage_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T5.Enabled = True
F3T5.Value = ""
F3T5.SetFocus
End Sub

Private Sub CCustSet_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T7.Enabled = True
F3T7.Value = ""
F3T7.SetFocus
End Sub

Private Sub CDiv_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T3.Enabled = True
F3T3.Value = ""
F3T3.SetFocus
End Sub

Private Sub CGeo_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T1.Enabled = True
F3T1.Value = ""
F3T1.SetFocus
End Sub

Private Sub CLI_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T2.Enabled = True
F3T2.Value = ""
F3T2.SetFocus
End Sub

Private Sub CPlat_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T8.Enabled = True
F3T8.Value = ""
F3T8.SetFocus
End Sub

Private Sub CTP_Click()
'This code corresponds to the little icon next to a textbox. If zirnCore can't figure out a particular dimension's value or it makes
    'a mistake the user can overide the automatic value with a manually entered one.
F3T4.Enabled = True
F3T4.Value = ""
F3T4.SetFocus
End Sub

Private Sub F3CCancel_Click()
Form3.Hide
'If a user clicks cancel form3 should close
deleteCode
End Sub

Private Sub F3CSubmit_Click()
'This is the final command button. Its purpose is to open the internet and pass the values on to the userform on the webpage.
    'This webpage will pull these dimension parameters and return client detail to the user.

deleteCode
End Sub

Sub zirnCore()
'This program will systematically go through a pull, starting at a target cell, and locate where each of the dimensions are located
    'as well as the selected values for each of those dimensions. Each dimension needs to have parameters and this program is meant
    'to figure out what selections a user made regarding said parameters. This is the main subroutine. In essence it functions
    'like a giant clock, where different cogs move thereby moving other cogs.

'First declare variables
Dim sName As String
'Above variable is for the sheet name
Dim scellx(0 To 3)
Dim scelly(0 To 3)
'The above two variables are for storing the x,y coordinates of three cells. The first cell is the target cell which has its coordinates
    'stored in cells A1 and B1. The second cell is called the key cell. It is the most upper left empty cell below the gray objects in a
    'TM1 mapped worksheet (in a pull). All white space navigation functions using the key cell and its relative distance from the target
    'cell.
    
Dim zArray(0 To 11, 0 To 11, 0 To 11)
Dim vArray(0 To 11)
Dim varx As Integer
Dim vary As Integer
Dim i As Integer
Dim k As Integer
Dim m As Integer
Dim h As Integer
Dim z As Integer
Dim wString As String
Dim r As String
Dim b As Integer
Dim dstring As String
Dim dArray(0 To 10)
Dim fArray(0 To 10)
Dim textBoxVal(0 To 10)

b = 11
'b is used a safe spot in the multidimensional arrays below. The llth spot in each array will
sName = ActiveSheet.Name
'Assign the value of the active sheet's name to the variable SName

scellx(0) = Worksheets(sName).Range("A1").Value
scelly(0) = Worksheets(sName).Range("B1").Value
'Assign the coordinates of the target cell to the first variable space in the scell array

On Error GoTo ErrHandler:
'If there is an error form3 should close and the user should restart the program

Worksheets(sName).Cells(scelly(0), scellx(0)).Activate
'Select the target cell
Selection.End(xlToLeft).Select
Selection.End(xlUp).Select
Selection.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
'Move to the key cell. These steps will always move the cursor from the target cell to the key cell.

scellx(1) = ActiveCell.Column
scelly(1) = ActiveCell.Row
'Assign the coordinates of the key cell to the second variable space in the scell array.

Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
'Select the key cell
Selection.End(xlToRight).Select
'Find how many white cells (empty cells) are to the right of the key cell until a non-white cell.
varx = ActiveCell.Column
scellx(2) = (varx - scellx(1))
'Store the number of white cells to the right of the key cell in the third variable space in the scell array. This is for the x coordinate.

Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
'Select the key cell
Selection.End(xlDown).Select
'Find how many white cells (empty cells) are below the key cell until a non-white cell.
vary = ActiveCell.Row
scelly(2) = (vary - scelly(1))
'Store the number of white cells below the key cell in the third variable space in the scell array. This is for the y coordinate.

Worksheets(sName).Cells(scelly(1), scellx(0)).Activate
'Now the subroutine will systematically go through all the dimensions that have been dragged out to the right of the pull and store their coordinates.
    'The above statement selects the cell that is on the same row as the key cell and on the same column as the target cell.

'-------------------------------------------------------------------------------------------------------------------------------
i = 0
Do While i < scelly(2)
'Scelly(2) is the number of dimensions (or white spaces) to the right of the key cell. This is equal to the number of dimensions that are dragged out to the
    'right. This will loop through and store three values in the zArray: the x-coordinate for the dimensions cell, the y-coordinate for the dimensions cell
    'and the value that is stored in that cell. Each value will be one of the 11 dimensions (eg case, customer set, channel etc)
    
Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
ActiveCell.Offset(0, (i + scellx(2))).Activate

zArray(i, i, b) = ActiveCell.Column
'Store the x-coordinate
zArray(i, b, i) = ActiveCell.Row
'Store the y-coordinate
zArray(i, b, b) = ActiveCell.Value
'Store the value (eg the dimension that is in that cell)

'The below msgboxes are good for testing what gets stored. I used them for trouble shooting and testing my logic. They can help a user follow the steps.
'MsgBox (zarray(i, i, b))
'MsgBox (zarray(i, b, i))
'MsgBox (zarray(i, b, b))

i = i + 1

Loop
'Once the subroutine exits the loop above that means it has recognized all the dimensions that are broken out horizontally. Next it will figure out the
    'dimensions that are positioned vertically below the key cell.

'-------------------------------------------------------------------------------------------------------------------------------
j = 0
Do While j < scellx(2)
'Scellx(2) is the number of dimensions (or white spaces) below the key cell. This is equal to the number of dimensions that are dragged out below
    'This will loop through and store three values in the zArray: the x-coordinate for the dimensions cell, the y-coordinate for the dimensions cell
    'and the value that is stored in that cell. Each value will be one of the 11 dimensions (eg case, customer set, channel etc)
    
Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
ActiveCell.Offset(i, j).Activate

zArray((i + j), (i + j), b) = ActiveCell.Column
'Store the x-coordinate
zArray((i + j), b, (i + j)) = ActiveCell.Row
'Store the y-coordinate
zArray((i + j), b, b) = ActiveCell.Value
'Store the value (eg the dimension that is in that cell)

'The below msgboxes are good for testing what gets stored. I used them for trouble shooting and testing my logic. They can help a user follow the steps.
'MsgBox (zarray((i + j), (i + j), b))
'MsgBox (zarray((i + j), b, (i + j)))
'MsgBox (zarray((i + j), b, b))

j = j + 1

Loop
'Once the subroutine exits the loop above that means it has recognized all the dimensions that are broken out vertically below. Next it will figure out the
    'dimensions that are positioned vertically above the key cell. Also recognize that the values stored in the above loop construct are added to the
    'pre-existing zArray array. This array keeps a running accumulation of the dimensions' locations.

'-------------------------------------------------------------------------------------------------------------------------------
k = 11 - (i + j)
'k represents the number of dimensions left to be identified above the key cell
m = 1
Do While (m - 1) < k
'k - (m - 1) represents the number of dimensions left to be found and stored

Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
ActiveCell.Offset((m * -1), 0).Activate

zArray((i + j + (m - 1)), (i + j + (m - 1)), b) = ActiveCell.Column
'Store the x-coordinate
zArray((i + j + (m - 1)), b, (i + j + (m - 1))) = ActiveCell.Row
'Store the y-coordinate
wString = ActiveCell.Formula
r = ((InStr(21, wString, ",")) - 22)
dstring = Mid(wString, 21, r)
zArray((i + j + (m - 1)), b, b) = dstring
'Store the value (eg the dimension that is in that cell). Note that storing the values for the dimensions that are vertically above the key cell is a little
    'more difficult due to the fact that they are not straight values. They have formulas and objects. By parsing the formula the subroutine can find which
    'dimension should be stored.

'The below msgboxes are good for testing what gets stored. I used them for trouble shooting and testing my logic. They can help a user follow the steps.
'MsgBox (zarray((i + j + (m - 1)), (i + j + (m - 1)), b))
'MsgBox (zarray((i + j + (m - 1)), b, (i + j + (m - 1))))
'MsgBox (zarray((i + j + (m - 1)), b, b))

m = m + 1

Loop
'Once the subroutine exits the loop above that means it has recognized all the dimensions that are broken out vertically above. Also recognize that the
    'values stored in the above loop construct are added to the pre-existing zArray array. This array keeps a running accumulation of the dimensions'
    'locations.

'-------------------------------------------------------------------------------------------------------------------------------
i = 0
Do While i < scelly(2)
Worksheets(sName).Cells(scelly(1), scellx(0)).Activate
ActiveCell.Offset((i + 1), 0).Activate
vArray(i) = ActiveCell.Value
i = i + 1
Loop
'The loop construct above will store the value that the parameter that the user select. For each dimension a different single parameter has been selected.
    'An example is Customer Set. In the first three loops above the subroutine finds out where the Customer Set dimension is on the workbook. It stores the
    'address and the word Customer Set. Now the above loop will match that array with the proper parameter (eg it will now assign the corresponding array
    'variable spaces with Financial Services or Public, depending on what the user requested to see).

j = 0
Do While j < scellx(2)
Worksheets(sName).Cells(scelly(0), scellx(1)).Activate
ActiveCell.Offset(0, j).Activate
vArray(i + j) = ActiveCell.Value
j = j + 1
Loop
'The loop construct above will store the value that the parameter that the user select. For each dimension a different single parameter has been selected.
    'An example is Customer Set. In the first three loops above the subroutine finds out where the Customer Set dimension is on the workbook. It stores the
    'address and the word Customer Set. Now the above loop will match that array with the proper parameter (eg it will now assign the corresponding array
    'variable spaces with Financial Services or Public, depending on what the user requested to see).

k = 11 - (i + j)
m = 1
Do While (m - 1) < k
Worksheets(sName).Cells(scelly(1), scellx(1)).Activate
ActiveCell.Offset((m * -1), 0).Activate
wString = ActiveCell.Formula
r = (InStrRev(wString, ",") + 2)
dstring = Mid(wString, r, ((Len(wString) - 1) - r))
vArray((i + j + (m - 1))) = dstring
'Storing the values for the dimensions vertically above the key cell require some string formulas to parse the formulas and find the proper strings
m = m + 1
Loop
'The loop construct above will store the value that the parameter that the user select. For each dimension a different single parameter has been selected.
    'An example is Customer Set. In the first three loops above the subroutine finds out where the Customer Set dimension is on the workbook. It stores the
    'address and the word Customer Set. Now the above loop will match that array with the proper parameter (eg it will now assign the corresponding array
    'variable spaces with Financial Services or Public, depending on what the user requested to see).

'Now an artificial array is created. This array will be responsible for putting the other two cogs in motion. A loop will allow the above arrays match up
    'and then send the corresponding values to the proper textboxes on form3.
    
fArray(0) = "case"
fArray(1) = "geography"
fArray(2) = "line_item_codes"
fArray(3) = "division"
fArray(4) = "period"
fArray(5) = "coverage"
fArray(6) = "channel"
fArray(7) = "customer_set"
fArray(8) = "platform"
fArray(9) = "acquisition_type"
fArray(10) = "acquisition"
'These values all represent the 11 dimensions that must each have a parameter value. Each would have been selected by the user.


For h = 0 To 10
    For z = 0 To 10
    If fArray(h) = zArray(z, b, b) Then
        textBoxVal(h) = vArray(z)
    End If
    Next z
Next h
'The above array is single most important loop in this program. Here what is happening is the artificial array, fArray, will loop from 0 to 10. On each loop
    'one zArray will have the same value. The second loop will loop the zArray within the fArray and once they match the corresponding value in vArray will
    'be assigned to the proper textbox. The interesting thing about this loop is that exactly how it runs will be different every time it runs. Also it can
    'not fail unless an improper value has been stored in either zArray or vArray. Here the fArray cog is rotating while within it the zArray is looping much
    'faster. Once they match the final cog, textBoxVal will move one tick because it was assigned a single value from the vArray.

'The below loop helps a user trouble shoot or watch how the propram is storing during each cycle of the macro loop.
'For q = 0 To 10
'MsgBox (textBoxVal(q))
'Next q

F3T0 = textBoxVal(0)
F3T1 = textBoxVal(1)
F3T2 = textBoxVal(2)
F3T3 = textBoxVal(3)
F3T4 = textBoxVal(4)
F3T5 = textBoxVal(5)
F3T6 = textBoxVal(6)
F3T7 = textBoxVal(7)
F3T8 = textBoxVal(8)
F3T9 = textBoxVal(9)
F3T10 = textBoxVal(10)
'The textboxes on form3 are now given the values that have been saved for them in the textBoxVal(x) array.

ErrHandler: Form3.Hide
'If there is an error anywhere along the way Form3 should close and the user should restart the program
Rows("1:1").Select
Selection.Delete
Worksheets(sName).Cells((scelly(0) - 1), scellx(0)).Activate
'Delete the top row that was used to store values during the transition from form2 to form3

End Sub

Private Sub UserForm_Terminate()
deleteCode
End Sub

Sub delCode()
'delCode will remove the code from the users active worksheet code module. This code was added in the subroutine called eventListener
    'early in form1.

'First declare variable
Dim sName As String


sName = Application.ActiveSheet.Name
'Assign variable the name of the users active sheet

With ThisWorkbook.VBProject.VBComponents(sName).CodeModule
.DeleteLines 1, .CountOfLines
'Deletes all lines of code in the active sheet module

End With

End Sub


Sub deleteCode()
     
On Error Resume Next
    With ActiveWorkbook.VBProject
        For x = .VBComponents.Count To 1 Step -1
            .VBComponents.Remove .VBComponents(x)
        Next x
        For x = .VBComponents.Count To 1 Step -1
            .VBComponents(x).CodeModule.DeleteLines _
            1, .VBComponents(x).CodeModule.CountOfLines
        Next x
    End With
On Error GoTo 0
     
End Sub

         