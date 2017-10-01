Const Separator As String = "|"
Const HeaderSufix As String = "h"

'選択している箇所をバックログ記法の表形式へ変換します
'選択範囲の1行目はヘッダになります
Sub selectToBacklogTable()
    Dim y As Long
    Dim x As Long
    Dim result As String
    Dim row As String
    Dim startY As Long
    
    startY = Selection(1).row
    
    result = ""
    
    For y = startY To Selection(Selection.Count).row
        row = Separator
        For x = Selection(1).Column To Selection(Selection.Count).Column
            row = row + CStr(Cells(y, x).Value) + Separator
        Next x
        If y = startY Then
            ' 1行目はヘッダにする
            row = row + HeaderSufix
        End If
        result = result + row + vbCrLf
    Next y
    
    InputBox Prompt:="コピペしてください(please copy and paste)", Default:=result
End Sub

