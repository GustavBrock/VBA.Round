Attribute VB_Name = "RoundingMethodsTestExcel"
' RoundingMethodsTestExcel v1.0.0
' (c) 2021-04-13. Gustav Brock, Cactus Data ApS, CPH.
' https://github.com/GustavBrock/VBA.Round
'
' Test function to list rounding of example values.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Verify correct Round returns using Excel.WorksheetFunction.Round.
' Returns True if all tests are passed.
'
' 2021-04-13. Gustav Brock, Cactus Data, CPH.
'
' Original source:
' 2005-06-14, Donald Lessau, Cologne.
' http://www.xbeat.net/vbspeed/IsGoodRound.htm
'
Public Function IsGoodRoundExcel() As Boolean

    Dim Failed As Boolean
  
    ' Note the differences to VBA/VB6's native Round function!
    
    ' Check zero.
    If Excel.WorksheetFunction.Round(0, 0) <> 0 Then Stop: Failed = True
    
    ' Check half-rounding.
    If Excel.WorksheetFunction.Round(1.49999999999999, 0) <> 1 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(1.5, 0) <> 2 Then Stop: Failed = True
    ' ! VB6 Round returns 2 ("banker's rounding").
    If Excel.WorksheetFunction.Round(2.5, 0) <> 3 Then Stop: Failed = True
    ' ! VB6 Round returns 1.234.
    If Excel.WorksheetFunction.Round(1.2345, 3) <> 1.235 Then Stop: Failed = True
    ' ! VB6 Round returns -1.234.
    If Excel.WorksheetFunction.Round(-1.2345, 3) <> -1.235 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(1.2355, 3) <> 1.236 Then Stop: Failed = True
    
    ' 2010-06-01: Bug fixed by Jeroen De Maeijer.
    If Excel.WorksheetFunction.Round(-0.0714285714, 1) <> -0.1 Then Stop: Failed = True
    
    ' 2006-02-01: Bug by Hallman.
    If Excel.WorksheetFunction.Round(0.09, 1) <> 0.1 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(0.0099, 1) <> 0# Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(0.0099, 2) <> 0.01 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(0.0099, 3) <> 0.01 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(0.0099, 4) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(Excel.WorksheetFunction.Round(1.01234012340125, 14)) <> 1.01234012340125 Then Stop: Failed = True
    ' ! VB6 Round returns 1.0123401234012.
    If Excel.WorksheetFunction.Round(1.01234012340125, 13) <> 1.0123401234013 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(Excel.WorksheetFunction.Round(10 ^ 13 + 0.74, 1)) <> 10000000000000.7 Then Stop: Failed = True
    ' ! VB6 Round returns -9999999999999.2.
    If Excel.WorksheetFunction.Round(-10 ^ 13 + 0.75, 1) <> -9999999999999.3 Then Stop: Failed = True
    ' ! VB6 error 5.
    If Excel.WorksheetFunction.Round(1.11111111111111E+16, -15) <> 1.1E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If Excel.WorksheetFunction.Round(10 ^ 307, 0) <> 1E+307 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(-10 ^ 308, 0) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(Excel.WorksheetFunction.Round(10.5, 20)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5.
    If NiceDbl(Excel.WorksheetFunction.Round(10.5, -20)) <> 0 Then Stop: Failed = True
    
    ' Check negative numbers (should round, not truncate).
    If Excel.WorksheetFunction.Round(-226.6, 0) <> -227 Then Stop: Failed = True
    ' ! VB6 Round returns -226.
    If Excel.WorksheetFunction.Round(-226.5, 0) <> -227 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(-226.4, 0) <> -226 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these.
    If Excel.WorksheetFunction.Round(226.7, -1) <> 230 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(226.7, -2) <> 200 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(226.7, -3) <> 0 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(226.7, -4) <> 0 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    ' ! VB6 Round fails on all four ("banker's rounding")
    ' Some emulations fail on the first two.
    If Excel.WorksheetFunction.Round(2.445, 2) <> 2.45 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(-2.445, 2) <> -2.45 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(3.445, 2) <> 3.45 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(-3.445, 2) <> -3.45 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(100.05, 1) <> CDec(100.1) Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(-100.05, 1) <> CDec(-100.1) Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!).
    If Excel.WorksheetFunction.Round(30.675, 2) <> 30.68 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(31.675, 2) <> 31.68 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(32.675, 2) <> 32.68 Then Stop: Failed = True 'VB6 Round -> 32.67 !!
    If Excel.WorksheetFunction.Round(33.675, 2) <> 33.68 Then Stop: Failed = True 'VB6 Round -> 33.67 !!
    ' Added 2005-07-12.
    If Excel.WorksheetFunction.Round(128.015, 2) <> 128.02 Then Stop: Failed = True 'VB6 Round -> 128.01 !!
    If Excel.WorksheetFunction.Round(128.045, 2) <> 128.05 Then Stop: Failed = True 'VB6 Round -> 128.04 Bankers
    
    ' Twice the same value.
    If Excel.WorksheetFunction.Round(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    If Excel.WorksheetFunction.Round(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundExcel = Not Failed
  
End Function

' Helper for IsGoodRoundExcel
' 2002-04-04.
'
Private Function NiceDbl(Dbl As Double) As Double

    ' Some rounding algorithms return results that need a special
    ' treatment to cope with subtle floating point errors. For example:
    ' ? 10 ^ 13 + 0.74
    '   -> 10000000000000.7
    ' ? Round10(10 ^ 13 + 0.74, 1)
    '   -> 10000000000000.7
    ' BUT:
    ' ? Round10(10 ^ 13 + 0.74, 1) - 10000000000000.7
    '   -> 0.001953125  'and not 0 as you would expect!
    
    ' One way to handle this is to wrap the ReturnValue with Val.
    ' However, this does not work on systems where the
    ' decimal sign is not the period character (".").
    ' Val() is not such a smart function.
    '
    ' Another way is this, and it appears to work on all systems.
    NiceDbl = CDec(Dbl)
  
End Function

