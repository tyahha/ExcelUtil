Attribute VB_Name = "selectToBacklogTable"
Const Separator As String = "|"
Const HeaderSufix As String = "h"

'�I�����Ă���ӏ����o�b�N���O�L�@�̕\�`���֕ϊ����܂�
'�I��͈͂�1�s�ڂ̓w�b�_�ɂȂ�܂�
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
            ' 1�s�ڂ̓w�b�_�ɂ���
            row = row + HeaderSufix
        End If
        result = result + row + vbCrLf
    Next x
    
    InputBox Prompt:="�R�s�y���Ă�������(please copy and paste)", Default:=result
End Sub
