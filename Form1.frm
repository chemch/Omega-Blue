VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 
   Caption         =   "Connection"
   ClientHeight    =   2436
   ClientLeft      =   24
   ClientTop       =   288
   ClientWidth     =   2880
   OleObjectBlob   =   "Form1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Naming Conventions
'------------------------------------------------------------------------------------------------
'Form abbreviation first: F1, F2 or F3
'Followed by C for command button, I for image, F for frame, T for textbox, L for label
'Followed by the name or number that describes that object. Example: F1T1 is form1 textbox1
'This program uses only forms, no module sheets
'Pieces of this program will add code to a users workbook and later delete that code
'Subroutines appear in the code sheets in the order that they will normally run
'Any subroutine that is called by another is located at the bottom of the code sheets
'------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
'This is the first code to execute. It will check to see if a user is signed into TM1. It calls two subroutines.


Application.Wait (Now + TimeValue("0:00:02"))
'Wait a few seconds. This adds a cinematic feel to the program

showImgPersp
'Call a subroutine that will check if a user is signed into TM1 Perspectives. Green check means yes. Red x means no.

Application.Wait (Now + TimeValue("0:00:02"))
'Wait a few seconds. This adds a cinematic feel to the program

showImgIcer
'Call a subroutine that will check if a user is signed into TM1 Icer. That is this program so they are always signed in. Always green check.
End Sub

Private Sub F1CSubmit_Click()
'Once a user knows his connection status the program will prepare for him/her to select the cell he/she wants to analyze


Range("A1").Select
Selection.EntireRow.Insert
Selection.EntireRow.Hidden = True
Cells(1, 1).Select
'Hide a row where the program can store relevant information that needs to be passed on to the next form

Form1.Hide
MsgBox ("Double click to select a cell.")
'Notify the user that he/she will need to next select a cell by double clicking on the value within that cell to proceed

addEventListener
'This subroutine functions a lot like a virus whereby it plants its own code in a users active sheet so that when they perform a certain event
    'it will trigger (eg when they double click a cell). Later it will remove itself and the user will never know it was there in the first
    'place.
    
End Sub

Private Sub F1CCancel_Click()
Form1.Hide
deleteCode
End Sub

Sub showImgPersp()
'This is the top image that will display. It will be either a red x or a green check mark based on their connection status.
    'It checks to see if a user is signed into TM1 Perspectives

'Variables are first declared
Dim perspStat As Integer
Dim var As Variant
Dim varlen As Variant

Range("A1").Select
Selection.EntireRow.Insert
Cells(1, 1).Select
ActiveCell.FormulaR1C1 = "=TM1USER("")"
var = ActiveCell.Value
'A formula tells is added to cell(1,1) and the returned value is assigned to a variable

If IsError(var) Then
    F1I1Red.Visible = True
    Range("A1").Select
    Selection.EntireRow.Delete
     Exit Sub
End If
'An error will mean that no user is signed into TM1 Perspectives

varlen = Len(var)
If varlen > 0 Then
    F1I1Green.Visible = True
End If
'If the length of the variable var is greater than 0 characters then a user is signed in

If varlen = 0 Then
    F1I1Red.Visible = True
End If
'If the length of the variable var is not greater than 0 characters then a user is not signed in

Range("A1").Select
Selection.EntireRow.Delete
'Delete this row as it was just used to add the formula and check the connection status

End Sub

Sub showImgIcer()
'This will check to see if a user is signed into TM1 Icer. They will always be signed in so it will always be a green check.

F1I2Green.Visible = True

End Sub

Sub addEventListener()
'This subroutine will insert 14 lines of code into the users active sheet's code module. Later it will be deleted. What it does is it
    'adds an event listener. The event that excel will be listening for is the event double click. Technically the program ends here.
    'When a user double clicks a value in a cell the program will be restarted at a different stage and will proceed from there.

'Declare variables first
Dim sName As String
Dim code1 As String
Dim code2 As String
Dim code3 As String
Dim code4 As String
Dim code5 As String
Dim code6 As String
Dim code7 As String
Dim code8 As String
Dim code9 As String
Dim code10 As String
Dim code11 As String
Dim code12 As String
Dim code13 As String
Dim code14 As String

'Each line of code is stored in a variable. Each line of code must be inserted on its own line in order for VBA to recognize the lines as
    'distinct instructions.
    
'The code that will be inserted into the user's active worksheet starts here
'-------------------------------------------------------------------------------------------------------------
code1 = "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)"
'This code will be inserted into the users code module for the users active worksheet. It will be listening for the double click event.


'Declare variables first
code2 = "Dim Selcelx"
code3 = "Dim selcely"
code4 = "If IsEmpty(ActiveCell.value) Then"
code5 = "Else"
'Check to see if a user double clicked an empty cell. If so the program should do nothing. A user must double click a valid cell.

code6 = "Selcelx = Target.Column"
code7 = "selcely = Target.Row"
code8 = "Cells(1, 1).Value = Selcelx"
code9 = "Cells(1, 2).Value = selcely"
'Store the row and column of the cell that the user double clicked. This cell is the "target cell". It is the cell the user wants to
    'analyze.

code10 = "Form2.F2T2.Value = Selcelx"
code11 = "Form2.F2T1.Value = selcely"
'Fill the text boxes on form2 with the address of the target cell so a user can verify his/her choice

code12 = "Form2.Show"

code13 = "End If"

code14 = "End Sub"
'This is the end of the subroutine that will be inserted into the user's worksheet.
'------------------------------------------------------------------------------------------------------------

sName = Application.ActiveSheet.Name

With ThisWorkbook.VBProject.VBComponents(sName).CodeModule
.InsertLines .CountOfLines + 1, code1
.InsertLines .CountOfLines + 1, code2
.InsertLines .CountOfLines + 1, code3
.InsertLines .CountOfLines + 1, code4
.InsertLines .CountOfLines + 1, code5
.InsertLines .CountOfLines + 1, code6
.InsertLines .CountOfLines + 1, code7
.InsertLines .CountOfLines + 1, code8
.InsertLines .CountOfLines + 1, code9
.InsertLines .CountOfLines + 1, code10
.InsertLines .CountOfLines + 1, code11
.InsertLines .CountOfLines + 1, code12
.InsertLines .CountOfLines + 1, code13
.InsertLines .CountOfLines + 1, code14
'Insert all the lines of code that are stored in the variables above into the module sheet for the users active worksheet.
End With

End Sub

Private Sub UserForm_Terminate()
deleteCode
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



g0??]@?F? U`A?k?x??if????d?v????^,c?'R?
j?*???m|7?Mu??S???????????T??.???Z??c??>D~