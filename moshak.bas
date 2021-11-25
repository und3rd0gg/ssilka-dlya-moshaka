Attribute VB_Name = "NewMacros"
Sub LinkForMoshak()
Attribute LinkForMoshak.VB_Description = "Правильная ссылка для Мошака"
Attribute LinkForMoshak.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.LinkForMoshak"
'
' LinkForMoshak Макрос
' Правильная ссылка для Мошака
'
    With ActiveDocument.Range(Start:=ActiveDocument.Content.Start, End:= _
        ActiveDocument.Content.End).EndnoteOptions
        .Location = wdEndOfDocument
        .NumberingRule = wdRestartContinuous
        .StartingNumber = 1
        .NumberStyle = wdNoteNumberStyleArabic
    End With
    With Selection
        With .EndnoteOptions
            .Location = wdEndOfDocument
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
        End With
        .Endnotes.Add Range:=Selection.Range, Reference:=""
    End With
    If ActiveWindow.ActivePane.View.Type = wdPrintView Or ActiveWindow. _
        ActivePane.View.Type = wdWebView Or ActiveWindow.ActivePane.View.Type = _
        wdPrintPreview Then
        ActiveWindow.View.SeekView = wdSeekMainDocument
    Else
        ActiveWindow.Panes(2).Close
    End If
    Application.Keyboard (1033)
    Selection.TypeText Text:="["
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="]"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.Font.Superscript = wdToggle
    Selection.Font.Superscript = wdToggle
    Selection.Font.Bold = wdToggle
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub
