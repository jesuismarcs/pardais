Sub Definir_Hierarquia()

    ' --- PONTO DE ENTRADA DA VERIFICAÇÃO DE LICENÇA ---
    If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
    ' --- FIM DA VERIFICAÇÃO ---
    
    Dim tsk As MSProject.Task
    Dim currentCode As String
    Dim rawText11Value As Variant
    Dim parts As Variant
    Dim targetLevelForThisTask As Integer
    Dim i As Long
    Dim originalCalculation As PjCalculation
    Dim errorInHierarchyStep As Boolean
    Dim lastParentTaskWithCode As MSProject.Task

    originalCalculation = Application.Calculation
    Application.Calculation = pjManual
    Application.ScreenUpdating = False

    On Error GoTo GlobalErrorHandler

    If ActiveProject.Tasks Is Nothing Then GoTo CleanUpAndExit
    If ActiveProject.Tasks.Count = 0 Then GoTo CleanUpAndExit
    
    Set lastParentTaskWithCode = Nothing
If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
    For Each tsk In ActiveProject.Tasks
        errorInHierarchyStep = False
        targetLevelForThisTask = 0

        If Not tsk Is Nothing And tsk.ID <> 0 And Not tsk.Placeholder Then
            rawText11Value = tsk.Text11
            currentCode = ""

            If Not IsError(rawText11Value) And Not IsNull(rawText11Value) And Not IsEmpty(rawText11Value) Then
                currentCode = Trim(CStr(rawText11Value))
                currentCode = Replace(currentCode, ",", ".")
                If Len(currentCode) > 0 And Right(currentCode, 1) = "." Then
                    currentCode = Left(currentCode, Len(currentCode) - 1)
                End If
            End If
If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
            If Len(currentCode) > 0 And IsValidCodeFormat(currentCode) Then
                parts = Split(currentCode, ".")
                targetLevelForThisTask = UBound(parts) + 1
            Else
                If Not lastParentTaskWithCode Is Nothing Then
                    If lastParentTaskWithCode.UniqueID <> tsk.UniqueID Then
                        targetLevelForThisTask = lastParentTaskWithCode.OutlineLevel + 1
                    Else
                        targetLevelForThisTask = 1
                        Set lastParentTaskWithCode = Nothing
                    End If
                Else
                    targetLevelForThisTask = 1
                End If
            End If
            
             If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
    If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
            If targetLevelForThisTask > 0 Then
                On Error Resume Next
                While tsk.OutlineLevel > 1
                    tsk.OutlineOutdent
                    If Err.Number <> 0 Then errorInHierarchyStep = True: Err.Clear: GoTo SkipHierarchyAdjustmentForThisTask
                Wend
                If targetLevelForThisTask > 1 And Not errorInHierarchyStep Then
                    For i = 1 To targetLevelForThisTask - 1
                        tsk.OutlineIndent
                        If Err.Number <> 0 Then errorInHierarchyStep = True: Err.Clear: Exit For
                    Next i
                End If
SkipHierarchyAdjustmentForThisTask:
                On Error GoTo GlobalErrorHandler
                If Len(currentCode) > 0 And IsValidCodeFormat(currentCode) Then
                    If Len(Trim(tsk.Text2 & "")) = 0 Then
                        Set lastParentTaskWithCode = tsk
                    End If
                End If
            End If
        End If
    Next tsk
    
 If Not VerificarLicencaNoServidor() Then
        MsgBox "Acesso não autorizado ou licença inválida. Por favor, contacte o fornecedor do programa através de 913239188. ", vbCritical, "Erro de Licença"
        Exit Sub
    End If
    
    MsgBox "Processo de atualização de hierarquia concluído.", vbInformation

CleanUpAndExit:
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = True
    Exit Sub

GlobalErrorHandler:
    MsgBox "Ocorreu um erro geral inesperado: " & Err.Description & " (Erro nº " & Err.Number & ")", vbCritical
    On Error Resume Next
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
