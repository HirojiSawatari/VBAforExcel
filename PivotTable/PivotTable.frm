VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PivotTable 
   Caption         =   "PivotTable"
   ClientHeight    =   2880
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4728
   OleObjectBlob   =   "PivotTable.frx":0000
   StartUpPosition =   1  '����������
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

    '��������
    Set ptcache1 = ActiveWorkbook.PivotCaches.Add(xlDatabase, Sheet1.Range("A1").CurrentRegion.Address)
    '����͸�ӱ��µĹ�����
    Set pt1 = ptcache1.CreatePivotTable("", "PT1")

    '�趨���ֶΣ�ҳ�ֶ�
    pt1.AddFields RowFields:=Array("�����"), PageFields:=Array("��ҵ��")

    '�趨ֵ�ֶ�
    With pt1.PivotFields("ͼƬ����")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    With pt1.PivotFields("������ϼ�")
       .Orientation = xlDataField
       .Function = xlSum
    End With

    Set ptcache2 = ActiveWorkbook.PivotCaches.Add(xlDatabase, Sheet1.Range("A1").CurrentRegion.Address)
    '�������й�����
    Set pt2 = ptcache2.CreatePivotTable("", "PT1")

    pt2.AddFields RowFields:=Array("�����"), PageFields:=Array("�ʼ���")
    With pt2.PivotFields("ʵ������")
        .Orientation = xlDataField
        .Function = xlSum
    End With
    With pt2.PivotFields("��������")
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
    '�������ÿ��
    For i = 2 To ir
        '�ҵ���Ӧ�����ݣ�ʵ��������
        total = Cells(i, 17)
        '���ֵ�����±�Ϊ����ID������
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
