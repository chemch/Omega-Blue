VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form2 
   Caption         =   "Selection"
   ClientHeight    =   2412
   ClientLeft      =   24
   ClientTop       =   276
   ClientWidth     =   3456
   OleObjectBlob   =   "Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
'When the form first opens the text boxes will be filled with the address coordinates of the target cell
F2T1 = Range("A1").Value
F2T2 = Range("B1").Value
End Sub

Private Sub F2C1_Click()
'If the click to proceed form2 will close and form3 will open
Form2.Hide
Form3.Show
End Sub

Private Sub F2C2_Click()
Form2.Hide
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

Private Sub UserForm_Terminate()
deleteCode
End Sub
nto D a t a . C l p D o c H a s S e s s i o n M e t a d a t a " :   t r u e ,   " D a t a . C l p D o c S e s s M e t a d a t a D i r t y " :   f a l s e ,   " D a t a . C l p D o c H a s S p o M e t a d a t a " :   f a l s e ,   " D a t a . C l p D o c H a s S p o P k g " :   f a l s e ,   " D a t a . C l p D o c C n t F a i l S e t L b l s " :   0 ,   " D a t a . C l p D o c H a s I d e n t i t y " :   f a l s e ,   " D a t a . r e s u l t " :   0 ,   " D a t a . d w T a g " :   5 8 9 8 4 9 3 5 2 }       