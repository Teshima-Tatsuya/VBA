Attribute VB_Name = "Strings"
Option Explicit

' 選択した範囲の文字列を結合する（デフォルトはカンマ区切り）
' rng 結合したい範囲
' delim 結合文字
Function JOIN(rng As Range, Optional delim As String = ",") As String
    Dim result As String
    Dim r As Range

    result = ""
    For Each r In rng
        result = result & r & delim
    Next

    JOIN = Left(result, Len(result) - 1)
End Function

' 文字列を右から検索して、最初にヒットした位置を返す。
' target 検索対象の文字列
' needle 検索する文字列
Function FINDR(target As String, needle As String) As Integer
    FINDR = InStrRev(target, needle)
End Function

Public Function VLOOKUPEX(needle As String, search_range As Range, return_array, Optional if_not_find = "") As String
    Dim cell As Range
    Dim i As Long
    
    i = 1
    For Each cell In search_range
        If cell.Value = needle Then
            VLOOKUPEX = return_array(i).Value
            Exit Function
        End If
        i = i + 1
    Next
    
    If TypeName(if_not_find) = "" Then
    End If
    VLOOKUPEX = if_not_find
End Function
