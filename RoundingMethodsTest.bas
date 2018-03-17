Attribute VB_Name = "RoundingMethodsTest"
' RoundingMethodsTest v1.2.1
' (c) 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
' https://github.com/GustavBrock/VBA.Round
'
' Test function to list rounding of example values.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Verify correct Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
' Original source:
' 2005-06-14, Donald Lessau, Cologne.
' http://www.xbeat.net/vbspeed/IsGoodRound.htm
'
Public Function IsGoodRound() As Boolean

    Dim Failed As Boolean
  
    ' Replace "RoundMid" with the name of your function to test.
    '
    ' Note the differences to VBA/VB6's native Round function!
    
    ' Check zero.
    If RoundMid(0) <> 0 Then Stop: Failed = True
    
    ' Check half-rounding.
    If RoundMid(1.49999999999999) <> 1 Then Stop: Failed = True
    If RoundMid(1.5) <> 2 Then Stop: Failed = True
    ' ! VB6 Round returns 2 ("banker's rounding").
    If RoundMid(2.5) <> 3 Then Stop: Failed = True
    ' ! VB6 Round returns 1.234.
    If RoundMid(1.2345, 3) <> 1.235 Then Stop: Failed = True
    ' ! VB6 Round returns -1.234.
    If RoundMid(-1.2345, 3) <> -1.235 Then Stop: Failed = True
    If RoundMid(1.2355, 3) <> 1.236 Then Stop: Failed = True
    
    ' 2010-06-01: Bug fixed by Jeroen De Maeijer.
    If RoundMid(-0.0714285714, 1) <> -0.1 Then Stop: Failed = True
    
    ' 2006-02-01: Bug by Hallman.
    If RoundMid(0.09, 1) <> 0.1 Then Stop: Failed = True
    If RoundMid(0.0099, 1) <> 0# Then Stop: Failed = True
    If RoundMid(0.0099, 2) <> 0.01 Then Stop: Failed = True
    If RoundMid(0.0099, 3) <> 0.01 Then Stop: Failed = True
    If RoundMid(0.0099, 4) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundMid(1.01234012340125, 14)) <> 1.01234012340125 Then Stop: Failed = True
    ' ! VB6 Round returns 1.0123401234012.
    If RoundMid(1.01234012340125, 13) <> 1.0123401234013 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(RoundMid(10 ^ 13 + 0.74, 1)) <> 10000000000000.7 Then Stop: Failed = True
    ' ! VB6 Round returns -9999999999999.2.
    If RoundMid(-10 ^ 13 + 0.75, 1) <> -9999999999999.3 Then Stop: Failed = True
    ' ! VB6 error 5.
    If RoundMid(1.11111111111111E+16, -15) <> 1.1E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundMid(10 ^ 307) <> 1E+307 Then Stop: Failed = True
    If RoundMid(-10 ^ 308) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundMid(10.5, 20)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5.
    If NiceDbl(RoundMid(10.5, -20)) <> 0 Then Stop: Failed = True
    
    ' Check negative numbers (should round, not truncate).
    If RoundMid(-226.6) <> -227 Then Stop: Failed = True
    ' ! VB6 Round returns -226.
    If RoundMid(-226.5) <> -227 Then Stop: Failed = True
    If RoundMid(-226.4) <> -226 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these.
    If RoundMid(226.7, -1) <> 230 Then Stop: Failed = True
    If RoundMid(226.7, -2) <> 200 Then Stop: Failed = True
    If RoundMid(226.7, -3) <> 0 Then Stop: Failed = True
    If RoundMid(226.7, -4) <> 0 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    ' ! VB6 Round fails on all four ("banker's rounding")
    ' Some emulations fail on the first two.
    If RoundMid(2.445, 2) <> 2.45 Then Stop: Failed = True
    If RoundMid(-2.445, 2) <> -2.45 Then Stop: Failed = True
    If RoundMid(3.445, 2) <> 3.45 Then Stop: Failed = True
    If RoundMid(-3.445, 2) <> -3.45 Then Stop: Failed = True
    If RoundMid(100.05, 1) <> CDec(100.1) Then Stop: Failed = True
    If RoundMid(-100.05, 1) <> CDec(-100.1) Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!).
    If RoundMid(30.675, 2) <> 30.68 Then Stop: Failed = True
    If RoundMid(31.675, 2) <> 31.68 Then Stop: Failed = True
    If RoundMid(32.675, 2) <> 32.68 Then Stop: Failed = True 'VB6 Round -> 32.67 !!
    If RoundMid(33.675, 2) <> 33.68 Then Stop: Failed = True 'VB6 Round -> 33.67 !!
    ' Added 2005-07-12.
    If RoundMid(128.015, 2) <> 128.02 Then Stop: Failed = True 'VB6 Round -> 128.01 !!
    If RoundMid(128.045, 2) <> 128.05 Then Stop: Failed = True 'VB6 Round -> 128.04 Bankers
    
    ' Twice the same value.
    If RoundMid(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    If RoundMid(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRound = Not Failed
  
End Function

' Verify correct Banker's Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
Public Function IsGoodRoundBankers() As Boolean
  
    Dim Failed As Boolean
    
    ' Replace "RoundMid" with the name of your function to test.
    '
    ' Note the differences to VBA/VB6's native Round function!
    
    ' Check zero.
    If RoundMid(0, , True) <> 0 Then Stop: Failed = True
    
    ' Check half-rounding.
    If RoundMid(1.49999999999999, , True) <> 1 Then Stop: Failed = True
    If RoundMid(1.5, , True) <> 2 Then Stop: Failed = True
    ' VB6 Round returns 2 ("banker's rounding").
    If RoundMid(2.5, , True) <> 2 Then Stop: Failed = True
    ' VB6 Round returns 1.234.
    If RoundMid(1.2345, 3, True) <> 1.234 Then Stop: Failed = True
    ' VB6 Round returns -1.234.
    If RoundMid(-1.2345, 3, True) <> -1.234 Then Stop: Failed = True
    If RoundMid(1.2355, 3, True) <> 1.236 Then Stop: Failed = True
    
    ' 2010-06-01: Bug fixed by Jeroen De Maeijer.
    If RoundMid(-0.0714285714, 1, True) <> -0.1 Then Stop: Failed = True
    
    ' 2006-02-01: Bug by Hallman.
    If RoundMid(0.09, 1, True) <> 0.1 Then Stop: Failed = True
    If RoundMid(0.0099, 1, True) <> 0# Then Stop: Failed = True
    If RoundMid(0.0099, 2, True) <> 0.01 Then Stop: Failed = True
    If RoundMid(0.0099, 3, True) <> 0.01 Then Stop: Failed = True
    If RoundMid(0.0099, 4, True) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundMid(1.01234012340125, 14, True)) <> 1.01234012340125 Then Stop: Failed = True
    ' VB6 Round returns 1.0123401234012.
    If RoundMid(1.01234012340125, 13, True) <> 1.0123401234012 Then Stop: Failed = True
    If RoundMid(1.01234012340135, 13, True) <> 1.0123401234014 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(RoundMid(10 ^ 13 + 0.74, 1, True)) <> 10000000000000.7 Then Stop: Failed = True
    ' VB6 Round returns -9999999999999.2
    If RoundMid(-10 ^ 13 + 0.75, 1, True) <> -9999999999999.2 Then Stop: Failed = True
    ' ! VB6 error 5.
    If RoundMid(1.11111111111111E+16, -15, True) <> 1.1E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundMid(10 ^ 307, , True) <> 1E+307 Then Stop: Failed = True
    If RoundMid(-10 ^ 308, , True) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundMid(10.5, 20, True)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5.
    If NiceDbl(RoundMid(10.5, -20, True)) <> 0 Then Stop: Failed = True
    
    ' Check negative numbers (should round, not truncate).
    If RoundMid(-226.6, , True) <> -227 Then Stop: Failed = True
    ' VB6 Round returns -226.
    If RoundMid(-226.5, , True) <> -226 Then Stop: Failed = True
    If RoundMid(-226.4, , True) <> -226 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these:
    If RoundMid(226.7, -1, True) <> 230 Then Stop: Failed = True
    If RoundMid(226.7, -2, True) <> 200 Then Stop: Failed = True
    If RoundMid(226.7, -3, True) <> 0 Then Stop: Failed = True
    If RoundMid(226.7, -4, True) <> 0 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    ' Some emulations fail on the first two.
    If RoundMid(2.445, 2, True) <> 2.44 Then Stop: Failed = True
    If RoundMid(-2.445, 2, True) <> -2.44 Then Stop: Failed = True
    If RoundMid(3.445, 2, True) <> 3.44 Then Stop: Failed = True
    If RoundMid(-3.445, 2, True) <> -3.44 Then Stop: Failed = True
    If RoundMid(100.05, 1, True) <> 100 Then Stop: Failed = True
    If RoundMid(-100.05, 1, True) <> -100 Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!).
    If RoundMid(30.675, 2, True) <> 30.68 Then Stop: Failed = True
    If RoundMid(31.675, 2, True) <> 31.68 Then Stop: Failed = True
    If RoundMid(32.675, 2, True) <> 32.68 Then Stop: Failed = True 'VB6 Round -> 32.67 !!
    If RoundMid(33.675, 2, True) <> 33.68 Then Stop: Failed = True 'VB6 Round -> 33.67 !!
    ' Added 2005-07-12
    If RoundMid(128.015, 2, True) <> 128.02 Then Stop: Failed = True 'VB6 Round -> 128.01 !!
    If RoundMid(128.045, 2, True) <> 128.04 Then Stop: Failed = True 'VB6 Round -> 128.04 Bankers
    
    ' Twice the same value.
    If RoundMid(1.01010101010101, 2, True) <> 1.01 Then Stop: Failed = True
    If RoundMid(1.01010101010101, 2, True) <> 1.01 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundBankers = Not Failed
  
End Function

' Verify correct Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
Public Function IsGoodRoundDown() As Boolean
  
    Dim Failed As Boolean
    
    ' Replace "RoundDown" with the name of your function to test.
    
    ' Check half-rounding.
    If RoundDown(1.49999999999999) <> 1 Then Stop: Failed = True
    If RoundDown(1.5) <> 1 Then Stop: Failed = True
    If RoundDown(2.5) <> 2 Then Stop: Failed = True
    
    ' Check other example values.
    If RoundDown(1.2345, 3) <> 1.234 Then Stop: Failed = True
    If RoundDown(-1.2345, 3) <> -1.235 Then Stop: Failed = True
    If RoundDown(1.2355, 3) <> 1.235 Then Stop: Failed = True
    
    If RoundDown(-0.0714285714, 1) <> -0.1 Then Stop: Failed = True
    
    If RoundDown(0.09, 1) <> 0 Then Stop: Failed = True
    If RoundDown(0.0099, 1) <> 0# Then Stop: Failed = True
    If RoundDown(0.0099, 2) <> 0 Then Stop: Failed = True
    If RoundDown(0.0099, 3) <> 0.009 Then Stop: Failed = True
    If RoundDown(0.0099, 4) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundDown(1.01234012340125, 14)) <> 1.01234012340125 Then Stop: Failed = True
    If RoundDown(1.01234012340125, 13) <> 1.0123401234012 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(RoundDown(10 ^ 13 + 0.74, 1)) <> 10000000000000.7 Then Stop: Failed = True
    If RoundDown(-10 ^ 13 + 0.75, 1) <> -9999999999999.3 Then Stop: Failed = True
    ' ! VB6 error 5
    If RoundDown(1.11111111111111E+16, -15) <> 1.1E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundDown(10 ^ 307) <> 1E+307 Then Stop: Failed = True
    If RoundDown(-10 ^ 308) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundDown(10.5, 20)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5
    If NiceDbl(RoundDown(10.5, -20)) <> 0 Then Stop: Failed = True
    
    ' Check negative numbers (should truncate).
    If RoundDown(-226.6) <> -227 Then Stop: Failed = True
    If RoundDown(-226.5) <> -227 Then Stop: Failed = True
    If RoundDown(-226.4) <> -227 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these:
    If RoundDown(226.7, -1) <> 220 Then Stop: Failed = True
    If RoundDown(226.7, -2) <> 200 Then Stop: Failed = True
    If RoundDown(226.7, -3) <> 0 Then Stop: Failed = True
    If RoundDown(226.7, -4) <> 0 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    If RoundDown(2.445, 2) <> 2.44 Then Stop: Failed = True
    If RoundDown(-2.445, 2) <> -2.45 Then Stop: Failed = True
    If RoundDown(3.445, 2) <> 3.44 Then Stop: Failed = True
    If RoundDown(-3.445, 2) <> -3.45 Then Stop: Failed = True
    If RoundDown(100.05, 1) <> 100# Then Stop: Failed = True
    If RoundDown(-100.05, 1) <> -100.1 Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!)
    If RoundDown(30.675, 2) <> 30.67 Then Stop: Failed = True
    If RoundDown(31.675, 2) <> 31.67 Then Stop: Failed = True
    If RoundDown(32.675, 2) <> 32.67 Then Stop: Failed = True
    If RoundDown(33.675, 2) <> 33.67 Then Stop: Failed = True
    
    If RoundDown(128.015, 2) <> 128.01 Then Stop: Failed = True
    If RoundDown(128.045, 2) <> 128.04 Then Stop: Failed = True
    
    ' Twice the same value.
    If RoundDown(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    If RoundDown(1.01010101010101, 2) <> 1.01 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundDown = Not Failed
  
End Function

' Verify correct Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
Public Function IsGoodRoundDownZero() As Boolean
  
    Dim Failed As Boolean
    
    ' Replace "RoundDown" with the name of your function to test.
    
    ' Check half-rounding.
    If RoundDown(1.49999999999999, , True) <> 1 Then Stop: Failed = True
    If RoundDown(1.5, , True) <> 1 Then Stop: Failed = True
    If RoundDown(2.5, , True) <> 2 Then Stop: Failed = True
    
    ' Check other example values.
    If RoundDown(1.2345, 3, True) <> 1.234 Then Stop: Failed = True
    If RoundDown(-1.2345, 3, True) <> -1.234 Then Stop: Failed = True
    If RoundDown(1.2355, 3, True) <> 1.235 Then Stop: Failed = True
    
    If RoundDown(-0.0714285714, 1, True) <> -0# Then Stop: Failed = True
    
    If RoundDown(0.09, 1, True) <> 0 Then Stop: Failed = True
    If RoundDown(0.0099, 1, True) <> 0# Then Stop: Failed = True
    If RoundDown(0.0099, 2, True) <> 0 Then Stop: Failed = True
    If RoundDown(0.0099, 3, True) <> 0.009 Then Stop: Failed = True
    If RoundDown(0.0099, 4, True) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundDown(1.01234012340125, 14, True)) <> 1.01234012340125 Then Stop: Failed = True
    If RoundDown(1.01234012340125, 13, True) <> 1.0123401234012 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(RoundDown(10 ^ 13 + 0.74, 1, True)) <> 10000000000000.7 Then Stop: Failed = True
    If RoundDown(-10 ^ 13 + 0.75, 1, True) <> -9999999999999.2 Then Stop: Failed = True
    ' ! VB6 error 5
    If RoundDown(1.11111111111111E+16, -15, True) <> 1.1E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundDown(10 ^ 307, , True) <> 1E+307 Then Stop: Failed = True
    If RoundDown(-10 ^ 308, , True) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundDown(10.5, 20, True)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5
    If NiceDbl(RoundDown(10.5, -20, True)) <> 0 Then Stop: Failed = True
    
    ' Check negative numbers (should truncate)
    If RoundDown(-226.6, , True) <> -226 Then Stop: Failed = True
    If RoundDown(-226.5, , True) <> -226 Then Stop: Failed = True
    If RoundDown(-226.4, , True) <> -226 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these:
    If RoundDown(226.7, -1, True) <> 220 Then Stop: Failed = True
    If RoundDown(226.7, -2, True) <> 200 Then Stop: Failed = True
    If RoundDown(226.7, -3, True) <> 0 Then Stop: Failed = True
    If RoundDown(226.7, -4, True) <> 0 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    If RoundDown(2.445, 2, True) <> 2.44 Then Stop: Failed = True
    If RoundDown(-2.445, 2, True) <> -2.44 Then Stop: Failed = True
    If RoundDown(3.445, 2, True) <> 3.44 Then Stop: Failed = True
    If RoundDown(-3.445, 2, True) <> -3.44 Then Stop: Failed = True
    If RoundDown(100.05, 1, True) <> 100# Then Stop: Failed = True
    If RoundDown(-100.05, 1, True) <> -100# Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!)
    If RoundDown(30.675, 2, True) <> 30.67 Then Stop: Failed = True
    If RoundDown(31.675, 2, True) <> 31.67 Then Stop: Failed = True
    If RoundDown(32.675, 2, True) <> 32.67 Then Stop: Failed = True
    If RoundDown(33.675, 2, True) <> 33.67 Then Stop: Failed = True
    
    If RoundDown(128.015, 2, True) <> 128.01 Then Stop: Failed = True
    If RoundDown(128.045, 2, True) <> 128.04 Then Stop: Failed = True
    
    ' Twice the same value.
    If RoundDown(1.01010101010101, 2, True) <> 1.01 Then Stop: Failed = True
    If RoundDown(1.01010101010101, 2, True) <> 1.01 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundDownZero = Not Failed
  
End Function

' Verify correct Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
Public Function IsGoodRoundUp() As Boolean
  
    Dim Failed As Boolean
    
    ' Replace "RoundUp" with the name of your function to test.
    
    ' Check half-rounding.
    If RoundUp(1.49999999999999) <> 2 Then Stop: Failed = True
    If RoundUp(1.5) <> 2 Then Stop: Failed = True
    If RoundUp(2.5) <> 3 Then Stop: Failed = True
    
    ' Check other example values.
    If RoundUp(1.2345, 3) <> 1.235 Then Stop: Failed = True
    If RoundUp(-1.2345, 3) <> -1.234 Then Stop: Failed = True
    If RoundUp(1.2355, 3) <> 1.236 Then Stop: Failed = True
    
    If RoundUp(-0.0714285714, 1) <> 0# Then Stop: Failed = True
    
    If RoundUp(0.09, 1) <> 0.1 Then Stop: Failed = True
    If RoundUp(0.0099, 1) <> 0.1 Then Stop: Failed = True
    If RoundUp(0.0099, 2) <> 0.01 Then Stop: Failed = True
    If RoundUp(0.0099, 3) <> 0.01 Then Stop: Failed = True
    If RoundUp(0.0099, 4) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundUp(1.01234012340125, 14)) <> 1.01234012340125 Then Stop: Failed = True
    If RoundUp(1.01234012340125, 13) <> 1.0123401234013 Then Stop: Failed = True
    
    ' Check large numbers.
    If NiceDbl(RoundUp(10 ^ 13 + 0.74, 1)) <> 10000000000000.7 Then Stop: Failed = True
    If RoundUp(-10 ^ 13 + 0.75, 1) <> -9999999999999.2 Then Stop: Failed = True
    ' ! VB6 error 5
    If RoundUp(1.11111111111111E+16, -15) <> 1.2E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundUp(10 ^ 307) <> 1E+307 Then Stop: Failed = True
    If RoundUp(-10 ^ 308) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundUp(10.5, 20)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5
    If NiceDbl(RoundUp(10.5, -20)) <> 1E+20 Then Stop: Failed = True
    
    ' Check negative numbers (should truncate).
    If RoundUp(-226.6) <> -226 Then Stop: Failed = True
    If RoundUp(-226.5) <> -226 Then Stop: Failed = True
    If RoundUp(-226.4) <> -226 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these:
    If RoundUp(226.7, -1) <> 230 Then Stop: Failed = True
    If RoundUp(226.7, -2) <> 300 Then Stop: Failed = True
    If RoundUp(226.7, -3) <> 1000 Then Stop: Failed = True
    If RoundUp(226.7, -4) <> 10000 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    If RoundUp(2.445, 2) <> 2.45 Then Stop: Failed = True
    If RoundUp(-2.445, 2) <> -2.44 Then Stop: Failed = True
    If RoundUp(3.445, 2) <> 3.45 Then Stop: Failed = True
    If RoundUp(-3.445, 2) <> -3.44 Then Stop: Failed = True
    If RoundUp(100.05, 1) <> 100.1 Then Stop: Failed = True
    If RoundUp(-100.05, 1) <> -100# Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!)
    If RoundUp(30.675, 2) <> 30.68 Then Stop: Failed = True
    If RoundUp(31.675, 2) <> 31.68 Then Stop: Failed = True
    If RoundUp(32.675, 2) <> 32.68 Then Stop: Failed = True
    If RoundUp(33.675, 2) <> 33.68 Then Stop: Failed = True
    
    If RoundUp(128.015, 2) <> 128.02 Then Stop: Failed = True
    If RoundUp(128.045, 2) <> 128.05 Then Stop: Failed = True
    
    ' Twice the same value.
    If RoundUp(1.01010101010101, 2) <> 1.02 Then Stop: Failed = True
    If RoundUp(1.01010101010101, 2) <> 1.02 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundUp = Not Failed
  
End Function

' Verify correct Round returns.
' Returns True if all tests are passed.
'
' 2015-08-24. Gustav Brock, Cactus Data, CPH.
'
Public Function IsGoodRoundUpZero() As Boolean
  
    Dim Failed As Boolean
    
    ' Replace "RoundUp" with the name of your function to test.
    
    ' Check half-rounding.
    If RoundUp(1.49999999999999, , True) <> 2 Then Stop: Failed = True
    If RoundUp(1.5, , True) <> 2 Then Stop: Failed = True
    If RoundUp(2.5, , True) <> 3 Then Stop: Failed = True
    
    ' Check other example values.
    If RoundUp(1.2345, 3, True) <> 1.235 Then Stop: Failed = True
    If RoundUp(-1.2345, 3, True) <> -1.235 Then Stop: Failed = True
    If RoundUp(1.2355, 3, True) <> 1.236 Then Stop: Failed = True
    
    If RoundUp(-0.0714285714, 1, True) <> -0.1 Then Stop: Failed = True
    
    If RoundUp(0.09, 1, True) <> 0.1 Then Stop: Failed = True
    If RoundUp(0.0099, 1, True) <> 0.1 Then Stop: Failed = True
    If RoundUp(0.0099, 2, True) <> 0.01 Then Stop: Failed = True
    If RoundUp(0.0099, 3, True) <> 0.01 Then Stop: Failed = True
    If RoundUp(0.0099, 4, True) <> 0.0099 Then Stop: Failed = True
    
    ' Check resolution.
    If NiceDbl(RoundUp(1.01234012340125, 14, True)) <> 1.01234012340125 Then Stop: Failed = True
    If RoundUp(1.01234012340125, 13, True) <> 1.0123401234013 Then Stop: Failed = True
    
    ' Check large numbers.
    ' Conversion to Decimal is necessary.
    If NiceDbl(RoundUp(CDec(10 ^ 13) + CDec(0.74), 1, True)) <> 10000000000000.8 Then Stop: Failed = True
    If RoundUp(CDec(-10 ^ 13) + CDec(0.75), 1, True) <> -9999999999999.3 Then Stop: Failed = True
    ' ! VB6 error 5
    If RoundUp(1.11111111111111E+16, -15, True) <> 1.2E+16 Then Stop: Failed = True
    
    ' Check very large numbers.
    If RoundUp(10 ^ 307, , True) <> 1E+307 Then Stop: Failed = True
    If RoundUp(-10 ^ 308, , True) <> -1E+308 Then Stop: Failed = True
    ' Check very large decimal places (arbitrary limit set to 20/-20).
    If NiceDbl(RoundUp(10.5, 20, True)) <> 10.5 Then Stop: Failed = True
    ' ! VB6 error 5
    If NiceDbl(RoundUp(10.5, -20, True)) <> 1E+20 Then Stop: Failed = True
    
    ' Check negative numbers (should truncate)
    If RoundUp(-226.6, , True) <> -227 Then Stop: Failed = True
    If RoundUp(-226.5, , True) <> -227 Then Stop: Failed = True
    If RoundUp(-226.4, , True) <> -227 Then Stop: Failed = True
    
    ' Check negative rounding.
    ' ! VB6 Round raises error 5 on all of these:
    If RoundUp(226.7, -1, True) <> 230 Then Stop: Failed = True
    If RoundUp(226.7, -2, True) <> 300 Then Stop: Failed = True
    If RoundUp(226.7, -3, True) <> 1000 Then Stop: Failed = True
    If RoundUp(226.7, -4, True) <> 10000 Then Stop: Failed = True
    
    ' Check rounding of nasty reals (tnx Gustav Brock).
    If RoundUp(2.445, 2, True) <> 2.45 Then Stop: Failed = True
    If RoundUp(-2.445, 2, True) <> -2.45 Then Stop: Failed = True
    If RoundUp(3.445, 2, True) <> 3.45 Then Stop: Failed = True
    If RoundUp(-3.445, 2, True) <> -3.45 Then Stop: Failed = True
    If RoundUp(100.05, 1, True) <> 100.1 Then Stop: Failed = True
    If RoundUp(-100.05, 1, True) <> -100.1 Then Stop: Failed = True
    '
    ' More nasty reals.
    ' ! VB6 Round totally fails on some of those numbers (!!)
    If RoundUp(30.675, 2, True) <> 30.68 Then Stop: Failed = True
    If RoundUp(31.675, 2, True) <> 31.68 Then Stop: Failed = True
    If RoundUp(32.675, 2, True) <> 32.68 Then Stop: Failed = True
    If RoundUp(33.675, 2, True) <> 33.68 Then Stop: Failed = True
    
    If RoundUp(128.015, 2, True) <> 128.02 Then Stop: Failed = True
    If RoundUp(128.045, 2, True) <> 128.05 Then Stop: Failed = True
    
    ' Twice the same value.
    If RoundUp(1.01010101010101, 2, True) <> 1.02 Then Stop: Failed = True
    If RoundUp(1.01010101010101, 2, True) <> 1.02 Then Stop: Failed = True
    
    ' Well done.
    IsGoodRoundUpZero = Not Failed
  
End Function

' Helper for IsGoodRound
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

' Produces examples of values rounded to significant figures.
'
' 2015-08-25. Gustav Brock, Cactus Data, CPH.
'
Public Function RoundingSignificantDemo()

    Dim Value               As Variant
    Dim RoundToEven         As Variant
    Dim RoundAwayFromZero   As Variant
    Dim Digits              As Integer
    Dim I                   As Integer
    
    Value = CCur(30.675)
    Digits = 14
    
    Debug.Print "Cur 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantCur(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantCur(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CCur(30.675)
    Digits = 4
    
    Debug.Print "Cur 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantCur(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantCur(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(30.675)
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(30.675)
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(-30.675)
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(-30.675)
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDec(-30.675) * 10 ^ 9
    Digits = 28
    
    Debug.Print "Dec 28"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDec(-30.675) * 10 ^ 9
    Digits = 4
    
    Debug.Print "Dec 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDec(Value + I * 10 ^ 9 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    
    Value = CDbl(30.675)
    Digits = 14
    
    Debug.Print "Dbl 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675)
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

    Value = CDbl(30.675) * 10 ^ 300
    Digits = 14
    
    Debug.Print "Dbl 14"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675) * 10 ^ 300
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ 300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

    Value = CDbl(30.675) * 10 ^ -300
    Digits = 8
    
    Debug.Print "Dbl 8"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print
    
    Value = CDbl(30.675) * 10 ^ -300
    Digits = 4
    
    Debug.Print "Dbl 4"
    For I = 0 To 10
        RoundToEven = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits, , True)
        RoundAwayFromZero = RoundSignificantDbl(Value + I * 10 ^ -300 / 100, Digits)
        Debug.Print "Value:" & Str(Value) & " - RoundedToEven:" & Str(RoundToEven) & " - RoundedAwayFromZero:" & Str(RoundAwayFromZero)
    Next
    Debug.Print

End Function
