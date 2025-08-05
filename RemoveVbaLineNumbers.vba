
Private Sub RemoveLineNumbers()
On Error GoTo HandleErrors

    Dim vbProj As Object  'VBIDE.VBProject
    Dim vbComp As Object  'VBIDE.VBComponent
    Dim modCurrent As Object  'VBIDE.CodeModule

    Dim iLine As Long
    Dim sLineText As String
    Dim sFirstWord As String
    Dim sReplacement As String
    Dim sNewText As String
    
    Set vbProj = Application.VBE.VBProjects(1)
    For Each vbComp In vbProj.VBComponents
        
        Set modCurrent = vbComp.CodeModule
        For iLine = 1 To modCurrent.CountOfLines
            sLineText = Trim$(modCurrent.Lines(iLine, 1))
            
            If Len(sLineText) > 0 Then
                
                If InStr(1, sLineText, " ") > 0 Then
                    sFirstWord = Trim$(Left$(sLineText, InStr(1, sLineText, " ") - 1))
                    If IsNumeric(sFirstWord) Then
                        sReplacement = String(Len(sFirstWord), " ")
                        sLineText = modCurrent.Lines(iLine, 1)
                        sNewText = Right$(sLineText, Len(sLineText) - (InStr(sLineText, sFirstWord) + Len(sFirstWord) - 1))
                        sNewText = Left$(sLineText, InStr(sLineText, sFirstWord) - 1) & sReplacement & sNewText
                        modCurrent.ReplaceLine iLine, sNewText
                    End If
                End If
            
            End If
        Next iLine
        
    Next
    
    MsgBox "The line numbers were removed from your VBA project successfully!", vbInformation, "Line Numbers Removed"
    
ExitMethod:
    Exit Sub
HandleErrors:
    MsgBox Err.Description, vbCritical, "Error " & Nz(Err.Number, "")
    Resume ExitMethod
End Sub

