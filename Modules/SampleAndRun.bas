Attribute VB_Name = "SampleAndRun"
Option Explicit

Const FIRSTROW As Integer = 4
Public iterations As Long, mean As Double, stddev As Double, min As Double, max As Double, prob As Double
Public varName As String, refcell As String, dist As String
Public inputhead As Range, dependent As String

Sub run_sampling()
    DistSelectionForm.Show
End Sub

Sub run_simulations()
    RunSimulationsForm.Show
End Sub

Private Sub sample()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Dim i As Long
    Dim lastcol As Long
    Dim outrng As Range
    Dim samp() As Double
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationSemiautomatic
    
    'Redefine samp as an array with iterations values
    ReDim samp(iterations)
    
    'Populate samp with specified number of rand outputs
    Select Case dist
    
    Case "Normal"
        For i = 0 To iterations - 1
            samp(i) = normaldist(mean, stddev)
        Next i
    Case "Lognormal"
        For i = 0 To iterations - 1
            samp(i) = Sgn(mean) * lognormdist(mean, stddev)
        Next i
    Case "Inverse Lognormal"
        For i = 0 To iterations - 1
            samp(i) = Sgn(mean) * invlognormdist(mean, stddev)
        Next i
    Case "Uniform"
        For i = 0 To iterations - 1
            samp(i) = uniformdist(min, max - 1)
        Next i
    Case "PERT"
        For i = 0 To iterations - 1
            samp(i) = pertdist(min, mean, max)
        Next i
    Case "Binomial"
        For i = 0 To iterations - 1
            samp(i) = binomialdist(prob)
        Next i
    
    End Select
    
    'Check if sims sheet exists, create if not
    If Not doessheetexist("sims") Then
        wb.Worksheets.Add.name = "sims"
        Set ws = wb.Worksheets("sims")
        ws.Activate
        ActiveWindow.Zoom = 85
        ActiveWindow.DisplayGridlines = False
    Else
        Set ws = wb.Worksheets("sims")
    End If
    
    'Find last column with data
    lastcol = findlastcol(FIRSTROW, ws)
    
    'Set last column to 0 if no data exists on sheet to avoid skipping over first column
    If lastcol = 1 And ws.Cells(FIRSTROW, 1).Value = 0 Then
        lastcol = 0
    End If
    
    'Set output range to start at cell A4
    Set outrng = Range(ws.Cells(FIRSTROW, lastcol + 1), ws.Cells(iterations + (FIRSTROW - 1), lastcol + 1))
    
    'Write input name and original cell reference to sheet to store relationship
    outrng(-1, 1) = varName
    'Need to input single quote when cell ref contains spaces
    If InStr(refcell, " ") > 0 Then
        outrng(0, 1) = "'" & refcell
    Else
        outrng(0, 1) = refcell
    End If
    
    'Write samp to output range defined above
    For i = 0 To iterations - 1
        outrng(i + 1, 1).Value = samp(i)
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub run_mcs()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim outws As Worksheet
    Dim i As Long, j As Long
    Dim inputVars() As String
    Dim varcnt As Integer
    Dim output() As Double, originalValue() As Double
    Dim lastrow As Long, lastcol As Long, iter As Long
    Dim outrng As Range, inrng As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationSemiautomatic
    
    'Initialize and fill inputVars array with cell refs of inputs
    varcnt = inputhead.Count
    ReDim inputVars(varcnt)
    
    For i = 1 To varcnt
        inputVars(i - 1) = inputhead(2, i)
    Next i
    
    'Redefine originalValue as array with items = number of variables/inputs
    ReDim originalValue(varcnt)
    
    'Set output target worksheet
    Set outws = wb.Worksheets("sims")
    
    'Find last column and last row with data in it from sims sheet
    'Set number of iterations to number of cells with data
    lastcol = findlastcol(FIRSTROW, outws)
    lastrow = findlastrow(lastcol, outws)
    iter = lastrow - (FIRSTROW - 1)      'numerical data doesn't begin until row 4
    
    'Redefine output as array with iterations values
    ReDim output(iter)
    
    'Initialize input and output ranges
    Set outrng = Range(outws.Cells(FIRSTROW, lastcol + 1), outws.Cells(lastrow, lastcol + 1))
    Set inrng = Range(outws.Cells(FIRSTROW, 1), outws.Cells(lastrow, 1))
    
    'Save original input value so that it can be restored after running simulations
    For i = 0 To varcnt - 1
        originalValue(i) = Range(inputVars(i)).Value
    Next i
    
    'Plug in samples for input and save output value to array
    For i = 0 To iter - 1
        For j = 0 To varcnt - 1
            Range(inputVars(j)).Value = inrng(i + 1, j + 1)
        Next j
        output(i) = Range(dependent).Value
    Next i
    
    'Write output values to output range
    For i = 0 To iter - 1
        outrng(i + 1, 1).Value = output(i)
    Next i
    
    'Restore original input value
    For i = 0 To varcnt - 1
        Range(inputVars(i)).Value = originalValue(i)
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Function doessheetexist(sh As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(sh)
    On Error GoTo 0
    
    If Not ws Is Nothing Then doessheetexist = True
End Function

Function findlastcol(row As Long, ws As Worksheet) As Long
    findlastcol = ws.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function findlastrow(col As Long, ws As Worksheet) As Long
    findlastrow = ws.Cells(Rows.Count, col).End(xlUp).row
End Function

Function normaldist(x As Double, s As Double)
    normaldist = Application.WorksheetFunction.NormInv(Rnd(), x, s)
End Function

Function lognormdist(x As Double, s As Double)
    Dim lnmean As Double, lnstddev As Double
    
    lnmean = Application.WorksheetFunction.Ln((x ^ 2) / ((x ^ 2 + s ^ 2) ^ 0.5))
    lnstddev = Application.WorksheetFunction.Ln(1 + ((s ^ 2) / (x ^ 2))) ^ 0.5
    
    lognormdist = Application.WorksheetFunction.LogNorm_Inv(Rnd(), lnmean, lnstddev)
End Function

Function invlognormdist(x As Double, s As Double)
    Dim lnmean As Double, lnstddev As Double, invmean As Double, convfactor As Double
    
    'Set conversion factor to reverse the data to allow for negative skew distribution
    convfactor = Abs(x) + s * 10              'Sets conversion factor at 10 standard deviations from the mean
    invmean = convfactor - Abs(x)

    lnmean = Application.WorksheetFunction.Ln((invmean ^ 2) / ((invmean ^ 2 + s ^ 2) ^ 0.5))
    lnstddev = Application.WorksheetFunction.Ln(1 + ((s ^ 2) / (invmean ^ 2))) ^ 0.5
    
    invlognormdist = convfactor - Application.WorksheetFunction.LogNorm_Inv(Rnd(), lnmean, lnstddev)
End Function

Function pertdist(a As Double, b As Double, c As Double)
    Dim alpha As Double, beta As Double
    
    alpha = (4 * b + c - 5 * a) / (c - a)
    beta = (5 * c - a - 4 * b) / (c - a)
    
    pertdist = Application.WorksheetFunction.Beta_Inv(Rnd(), alpha, beta, a, c)
End Function

Function uniformdist(x As Double, y As Double)
    uniformdist = (y - x) * Rnd() + x
End Function

Function binomialdist(p As Double)
    binomialdist = Application.WorksheetFunction.Binom_Inv(1, p, Rnd())
End Function

