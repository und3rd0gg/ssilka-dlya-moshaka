Attribute VB_Name = "NewMacros"
Sub ���������������()
Attribute ���������������.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.���������������"
'
' ��������������� ������
'
'
    Application.Keyboard (1033)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="["
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Superscript = wdToggle
    Selection.Font.Superscript = wdToggle
    Selection.Font.Bold = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub
