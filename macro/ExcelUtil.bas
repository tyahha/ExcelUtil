Const Separator As String = "|"
Const HeaderSufix As String = "h"

'�I�����Ă���ӏ����o�b�N���O�L�@�̕\�`���֕ϊ����܂�
'�I��͈͂�1�s�ڂ̓w�b�_�ɂȂ�܂�
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
            ' 1�s�ڂ̓w�b�_�ɂ���
            row = row + HeaderSufix
        End If
        result = result + row + vbCrLf
    Next y
    
    InputBox Prompt:="�R�s�y���Ă�������(please copy and paste)", Default:=result
End Sub

