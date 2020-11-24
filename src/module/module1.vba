Option Explicit

Public Hashtable As Object

Type ActiveCellData
    row As Integer
    column As Integer
    name As String
End Type


Enum ItemNames
    apple
    banana
    candy
    unknown
End Enum


Sub Macro1()
'
' Macro1 Macro
' ハッシュテーブル
'
    Dim 辞書 As Object
    Set 辞書 = newHashTable()
    
    Dim 現在地情報 As ActiveCellData
    現在地情報 = newActiveCellData()
    
    Dim name As ItemNames
    name = fromStringToItemName(現在地情報.name)
    If name = ItemNames.unknown Then
        ' pass
    Else
        Cells(現在地情報.row + 2, 現在地情報.column).value = 辞書.item(name).leadtime
    End If
End Sub


Function newActiveCellData() As ActiveCellData
    Dim x As ActiveCellData
    x.row = ActiveCell.row
    x.column = ActiveCell.column
    x.name = Cells(x.row, x.column).value
    newActiveCellData = x
End Function


Function fromItemNameToLeadTime(ItemName As String) As String
    fromItemNameToLeadTime = "2week"
End Function


Function newHashTable() As Object
    Dim h As Object
    Set h = CreateObject("Scripting.Dictionary")
    setHashtable h, ItemNames.apple, "1week"
    setHashtable h, ItemNames.banana, "2week"
    setHashtable h, ItemNames.candy, "3week"
    setHashtable h, ItemNames.unknown, "ERROR"
        
    Set newHashTable = h
End Function

Function fromStringToItemName(aName As String) As ItemNames
    Select Case aName
        Case "リンゴ"
            fromStringToItemName = ItemNames.apple
        Case "バナナ"
            fromStringToItemName = ItemNames.banana
        Case "キャンディ"
            fromStringToItemName = ItemNames.candy
        Case Else
            MsgBox "ERROR Name: " & aName
            fromStringToItemName = ItemNames.unknown
    End Select
End Function

Sub setHashtable(h As Object, name As ItemNames, leadtime As String)
    Dim item As ItemData
    Set item = New ItemData
    item.initialize name, leadtime
    h.Add name, item
End Sub
