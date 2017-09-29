Attribute VB_Name = "selectToBacklogTable"
Const Separator As String = "|"
Const HeaderSufix As String = "h"

'選択している箇所をバックログ記法の表形式へ変換します
'選択範囲の1行目はヘッダになります
Sub selectToBacklogTable()
    Dim x As Long
    Dim y As Long
    Dim result As String
    Dim row As String
    
    result = ""
    For x = Selection(1).row To Selection(Selection.Count).row
        row = Separator
        For y = Selection(1).Column To Selection(Selection.Count).Column
            row = row + CStr(Cells(x, y).Value) + Separator
        Next y
        If Selection(x).row - Selection(1).row = 0 Then
            ' 1行目はヘッダにする
            row = row + HeaderSufix
        End If
        result = result + row + vbCrLf
    Next x
    
    InputBox Prompt:="コピペしてください(please copy and paste)", Default:=result
End Sub
