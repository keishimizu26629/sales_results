Attribute VB_Name = "main"
Option Explicit

Sub classify_amount()
    Dim Lists() As Contractor
    Dim data()
    Dim count As Integer: Dim i As Long

    Dim data_sheet As Worksheet
    Dim table_sheet As Worksheet
    Dim data_start_row As Integer: Dim table_start_row As Integer: Dim data_end_row As Long: Dim table_end_row As Long
    Dim contract_name As String: Dim kind As String: Dim amount As String

    Dim index_dic As Object: Set index_dic = CreateObject("Scripting.Dictionary")

    data_start_row = 2
    table_start_row = 4
    count = 1
    Set data_sheet = ActiveWorkbook.Sheets("【貼り付け用】e-Getsデータ")
    Set table_sheet = ActiveWorkbook.Sheets("ビルダー別実績")
    data_end_row = data_sheet.Cells(Rows.count, 5).End(xlUp).row
    table_end_row = table_sheet.Cells(Rows.count, 5).End(xlUp).row
    If table_start_row > table_end_row Then: table_end_row = table_start_row

    table_sheet.Range(table_sheet.Cells(table_start_row, 1), table_sheet.Cells(table_end_row, 10)) = ""

    Call empty_init(Lists, index_dic)

    For i = data_start_row To data_end_row
        contract_name = data_sheet.Cells(i, 69).Value
        kind = data_sheet.Cells(i, 59).Value
        amount = data_sheet.Cells(i, 14).Value
        If contract_name = "" Then
            contract_name = "その他"
        Else
            If Not index_dic.Exists(contract_name) Then
                Call create_instance(Lists, index_dic, count, contract_name)
                count = count + 1
            End If
        End If
        Call Lists(index_dic.Item(contract_name)).add_amount(kind, amount)
    Next

    Call conversion_data(Lists, data)

    If Not i = 0 Then: table_sheet.Range(table_sheet.Cells(table_start_row, 1), table_sheet.Cells(table_start_row + UBound(Lists), 10)) = WorksheetFunction.Transpose(data)

    msgbox "OK"

End Sub

Function empty_init(ByRef Lists, ByRef index_dic)
    Call create_instance(Lists, index_dic, 0, "その他")
End Function

Function create_instance(ByRef Lists, ByRef index_dic, ByVal count, ByRef contract_name)
    ReDim Preserve Lists(count)
    Set Lists(count) = New Contractor
    Lists(count).init (contract_name)
    index_dic.Add contract_name, count
End Function

Function conversion_data(ByRef Lists, ByRef data)
    Dim count As Long
    Dim i As Integer
    count = 0
    For i = 1 To UBound(Lists)
        Call assign_data(Lists, data, i, count)
        count = count + 1
    Next
    Call assign_data(Lists, data, 0, count)
End Function

Function assign_data(ByRef Lists, ByRef data, ByVal i, ByVal count)
    ReDim Preserve data(9, count)
    data(0, count) = Lists(i).name
    data(1, count) = "=SUM(RC[1]:RC[9])"
    data(2, count) = round(Lists(i).sash_amount / 1000, 0)
    data(3, count) = round(Lists(i).exterior_amount / 1000, 0)
    data(4, count) = round(Lists(i).sanitary_amount / 1000, 0)
    data(5, count) = round(Lists(i).kitchen_amount / 1000, 0)
    data(6, count) = round(Lists(i).ribiken_amount / 1000, 0)
    data(7, count) = round(Lists(i).panel_amount / 1000, 0)
    data(8, count) = round(Lists(i).electric_amount / 1000, 0)
    data(9, count) = round(Lists(i).others_amount / 1000, 0)
End Function

Sub test1()



End Sub
