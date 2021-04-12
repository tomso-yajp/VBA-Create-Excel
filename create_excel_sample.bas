Attribute VB_Name = "create_excel_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_object: create excel application
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_object()
Dim app As Variant
Set app = CreateObject("excel.application")
With app
  .Visible = True
  
  .DisplayAlerts = False
  With .Workbooks.Add(xlWBATWorksheet)
    .SaveAs Filename:=ThisWorkbook.Path & "\sample", FileFormat:=52
    'AppActivate (.Name)
    Call add_mojule(app)
    app.Run .Name & "!msg"
    .Close SaveChanges:=False
  End With
  .DisplayAlerts = True
  .Quit
End With
Set app = Nothing
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  add_mojule: add code to the module
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub add_mojule(Optional app As Variant, _
    Optional wname As String = "sample.xlsm")
Dim mname As String: mname = "sample"
Dim code As Variant
Dim i As Integer: i = 1
code = "Sub msg()" & vbLf & _
       "MsgBox ""シート付のブックを作成しました"",vbYes + vbInformation,""ブック作成：確認""" & vbLf & _
       "End Sub"
With app.Workbooks(wname)
  With .VBProject.VBComponents
    With .Add(1)
      .Name = mname
      With .CodeModule
        .InsertLines 1, code
      End With
    End With
  End With
End With
End Sub

