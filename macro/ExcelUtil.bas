Attribute VB_Name = "ExcelUtil"
Option Explicit
Const Separator As String = "|"
Const HeaderSufix As String = "h"
Const backlogNewLine = "&br;"

'選択している箇所をバックログ記法の表形式へ変換します
'選択範囲の1行目はヘッダになります
Sub selectToBacklogTable()
    Dim result As String
    Dim sel As Range
    Set sel = Application.selection
    If sel Is Nothing Then
        MsgBox "セルを選択してください"
    Else
        result = rangeToBacklogTable(sel)
        InputBox Prompt:="コピペしてください(please copy and paste)", Default:=result
    End If
End Sub

Function rangeToBacklogTable(sel As Range) As String
    Dim y As Long
    Dim x As Long
    Dim result As String
    Dim row As String
    Dim startY As Long
    
    startY = sel(1).row
    
    result = ""
    
    For y = startY To sel(sel.Count).row
        row = Separator
        For x = sel(1).Column To sel(sel.Count).Column
            row = row + CStr(Cells(y, x).Value) + Separator
        Next x
        If y = startY Then
            ' 1行目はヘッダにする
            row = row + HeaderSufix
        End If
        
        ' セル内の改行をバックログの改行に変換する
        ' 最後に一括で変換すると必要な改行までなくなってしまうので1行単位で置換する
        row = Replace(row, vbLf, backlogNewLine)
        
        result = result + row + vbCrLf
    Next y
    
    rangeToBacklogTable = result
    
End Function
