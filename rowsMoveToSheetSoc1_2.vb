
'Module 1
Sub MoveTodouble_MVM_PT()
    Dim xRg As Range
    Dim xCell As Range
    Dim A As Long
    Dim B As Long
    Dim I As Long
    A = Worksheets("master").UsedRange.Rows.Count
    B = Worksheets("double_MVM_PT").UsedRange.Rows.Count
    If B = 1 Then  
        If Application.WorksheetFunction.CountA(Worksheets("double_MVM_PT").UsedRange) = 0 Then B = 0
    End If
    Set xRg =  Worksheets("master").Range("I1:I" & A)
    On Error Resume Next
    Application.ScreenUpdating = False  
    For I = 1 To xRg.Count
        If CStr(xRg(I).Value) = "double_MVM_PT" Then
            xRg(I).EntireRow.Copy Destination:= Worksheets("double_MVM_PT").Range("A" & B + 1)
            xRg(I).EntireRow.Delete  
            If CStr(xRg(I).Value) = "double_MVM_PT" Then
                I = I - 1
            End If
            B = B + 1 
        End If
    Next 
    Application.ScreenUpdating = True  
End Sub

'Module 2
Sub MoveTodouble_MVM_NPT()
    Dim xRg As Range
    Dim xCell As Range
    Dim A As Long
    Dim B As Long
    Dim I As Long
    A = Worksheets("master").UsedRange.Rows.Count
    B = Worksheets("double_MVM_NPT").UsedRange.Rows.Count
    If B = 1 Then  
        If Application.WorksheetFunction.CountA(Worksheets("double_MVM_NPT").UsedRange) = 0 Then B = 0
    End If
    Set xRg =  Worksheets("master").Range("I1:I" & A)
    On Error Resume Next
    Application.ScreenUpdating = False  
    For I = 1 To xRg.Count
        If CStr(xRg(I).Value) = "double_MVM_NPT" Then
            xRg(I).EntireRow.Copy Destination:= Worksheets("double_MVM_NPT").Range("A" & B + 1)
            xRg(I).EntireRow.Delete  
            If CStr(xRg(I).Value) = "double_MVM_NPT" Then
                I = I - 1
            End If
            B = B + 1 
        End If 
    Next 
    Application.ScreenUpdating = True 
End Sub


'Module 3
Sub MoveTotheta_T_mul_PT()
    Dim xRg As Range
    Dim xCell As Range
    Dim A As Long
    Dim B As Long
    Dim I As Long
    A = Worksheets("master").UsedRange.Rows.Count
    B = Worksheets("theta_T_mul_PT").UsedRange.Rows.Count
    If B = 1 Then  
        If Application.WorksheetFunction.CountA(Worksheets("theta_T_mul_PT").UsedRange) = 0 Then B = 0
    End If
    Set xRg =  Worksheets("master").Range("I1:I" & A)
    On Error Resume Next
    Application.ScreenUpdating = False  
    For I = 1 To xRg.Count
        If CStr(xRg(I).Value) = "theta_T_mul_PT" Then
            xRg(I).EntireRow.Copy Destination:= Worksheets("theta_T_mul_PT").Range("A" & B + 1)
            xRg(I).EntireRow.Delete  
            If CStr(xRg(I).Value) = "theta_T_mul_PT" Then
                I = I - 1
            End If
            B = B + 1 
        End If 
    Next 
    Application.ScreenUpdating = True 
End Sub

'Module 4
Sub MoveTotheta_T_mul_NPT()
    Dim xRg As Range
    Dim xCell As Range
    Dim A As Long
    Dim B As Long
    Dim I As Long
    A = Worksheets("master").UsedRange.Rows.Count
    B = Worksheets("theta_T_mul_NPT").UsedRange.Rows.Count
    If B = 1 Then  
        If Application.WorksheetFunction.CountA(Worksheets("theta_T_mul_NPT").UsedRange) = 0 Then B = 0
    End If
    Set xRg =  Worksheets("master").Range("I1:I" & A)
    On Error Resume Next
    Application.ScreenUpdating = False  
    For I = 1 To xRg.Count
        If CStr(xRg(I).Value) = "theta_T_mul_NPT" Then
            xRg(I).EntireRow.Copy Destination:= Worksheets("theta_T_mul_NPT").Range("A" & B + 1)
            xRg(I).EntireRow.Delete  
            If CStr(xRg(I).Value) = "theta_T_mul_NPT" Then
                I = I - 1
            End If
            B = B + 1 
        End If
    Next 
    Application.ScreenUpdating = True 
End Sub

'Function to move rows automatically based on the string specified on the I column.
Private Sub Worksheet_Change(ByVal Target As Range)
        Dim Z As Long
        Dim xVal As String 
        On Error Resume Next
        If Intersect (Target, Range("I:I")) Is Nothing Then Exit Sub
        Application.EnableEvents = False
        For Z = 1 To Target.Count
            If Target(Z).Value > 0 Then
                'Call MoveTodouble_MVM_PT
                'Call MoveTodouble_MVM_NPT
                'Call MoveTotheta_T_mul_PT
                'Call MoveTotheta_T_mul_NPT
            End If
        Next
        Application.EnableEvents = True
End Sub