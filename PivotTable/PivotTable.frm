VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PivotTable 
   Caption         =   "PivotTable"
   ClientHeight    =   2880
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4728
   OleObjectBlob   =   "PivotTable.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "PivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim ptcache1 As PivotCache
    Dim pt1 As PivotTable
    Dim ptcache2 As PivotCache
    Dim pt2 As PivotTable

    '设置区域
    Set ptcache1 = ActiveWorkbook.PivotCaches.Add(xlDatabase, Sheet1.Range("A1").CurrentRegion.Address)
    '增加透视表到新的工作表
    Set pt1 = ptcache1.CreatePivotTable("", "PT1")

    '设定行字段，页字段
    pt1.AddFields RowFields:=Array("检测人"), PageFields:=Array("作业周")

    '设定值字段
    With pt1.PivotFields("图片点数")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    With pt1.PivotFields("核心项合计")
       .Orientation = xlDataField
       .Function = xlSum
    End With

    Set ptcache2 = ActiveWorkbook.PivotCaches.Add(xlDatabase, Sheet1.Range("A1").CurrentRegion.Address)
    '加入已有工作表
    Set pt2 = ptcache2.CreatePivotTable("", "PT1")

    pt2.AddFields RowFields:=Array("检测人"), PageFields:=Array("质检周")
    With pt2.PivotFields("实际总数")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    With pt2.PivotFields("错误总数")
        .Orientation = xlDataField
        .Function = xlSum
    End With

End Sub

Private Sub CommandButton2_Click()
    Dim d As Object
    Dim total As String
    
    Set d = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    On Error Resume Next
    
    Sheets("sheet1").Select
    ir = Range("b65536").End(xlUp).Row
    '遍历表的每行
    For i = 2 To ir
        '找到对应列数据（实际总数）
        total = Cells(i, 17)
        '在字典存入下标为任务ID的数据
        d(Cells(i, 1).Value & "") = str
        str = ""
    Next
    Sheets("sheet2").Select
    ir = Range("b65536").End(xlUp).Row
    For i = 3 To ir
        Cells(i, 19) = d(Cells(i, 4).Value & "")
    Next
    
    
    ar = ActiveWorkbook.Sheets(1).Range("A1").CurrentRegion
    ActiveWorkbook.Close False
    For i = 2 To UBound(ar)
        d(ar(i, 1)) = Array(ar(i, 2), ar(i, 3))
    Next
    With Sheet1
        For i = 2 To .[a1048576].End(3).Row
            If d.exists(.Cells(i, 1).Value) Then
                .Cells(i, 2) = Date
                .Cells(i, 3).Resize(1, 2) = d(.Cells(i, 1).Value)
            End If
        Next
    End With
    Application.ScreenUpdating = True
End Sub



Dim d As Object
Dim ir, i, j
Dim str$
Set d = CreateObject("scripting.dictionary")
Application.ScreenUpdating = False
 On Error Resume Next
Sheets("sheet1").Select
ir = Range("b65536").End(xlUp).Row
For i = 2 To ir
  For j = 3 To 8
  str = str & "|" & Cells(i, j)
  Next
  str = Mid(str, 2)
  d(Cells(i, 2).Value & "") = str
  str = ""
Next
 
Sheets("sheet2").Select
ir = Range("b65536").End(xlUp).Row
For i = 3 To ir
Range(Cells(i, 3), Cells(i, 8)) = Split(d(Cells(i, 2).Value & ""), "|")
 
Next
Application.ScreenUpdating = True
