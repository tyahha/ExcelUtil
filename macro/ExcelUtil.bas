Attribute VB_Name = "ExcelUtil"
Option Explicit
Const Separator As String = "|"
Const HeaderSufix As String = "h"
Const backlogNewLine = "&br;"

'�I�����Ă���ӏ����o�b�N���O�L�@�̕\�`���֕ϊ����܂�
'�I��͈͂�1�s�ڂ̓w�b�_�ɂȂ�܂�
Sub selectToBacklogTable()
    Dim result As String
    Dim sel As Range
    Set sel = Application.selection
    If sel Is Nothing Then
        MsgBox "�Z����I�����Ă�������"
    Else
        result = rangeToBacklogTable(sel)
        InputBox Prompt:="�R�s�y���Ă�������(please copy and paste)", Default:=result
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
            ' 1�s�ڂ̓w�b�_�ɂ���
            row = row + HeaderSufix
        End If
        
        ' �Z�����̉��s���o�b�N���O�̉��s�ɕϊ�����
        ' �Ō�Ɉꊇ�ŕϊ�����ƕK�v�ȉ��s�܂łȂ��Ȃ��Ă��܂��̂�1�s�P�ʂŒu������
        row = Replace(row, vbLf, backlogNewLine)
        
        result = result + row + vbCrLf
    Next y
    
    rangeToBacklogTable = result
    
End Function
