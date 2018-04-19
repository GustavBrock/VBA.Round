Attribute VB_Name = "Imperial"
' Imperial v1.0.4
' (c) 2018-04-19. Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Set of functions for converting between imperial and metric measures
' and formatting imperial measures.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)Option Explicit

Option Explicit

' Common constants.
'
' Meter/inch relation. 1 inch = 0.0254 m.
Public Const MetersPerInch      As Currency = 0.0254
' Inch/foot relation.
Public Const InchesPerFoot      As Currency = 12
'
' Ascii values for "Smart Quotes".
Public Const SmartSingleQuote   As Long = 146
Public Const SmartDoubleQuote   As Long = 148
'

' Rounds and formats a decimal value of inches to integer feet and inches and a fraction of inches
' applied either a default or a custom format.
'
' Parameter Exponent determines rounding. Rounds by default to integer inches.
' Parameter Format determines the format of the output. Default is: f' i-r"
' Parameter SmartQuotes will - if True, and if the output contains quotes - replace these quotes
' with "Smart Quotes" as used in Word.
'
' Parameters RoundingMethod determines the rounding method.
' Default is by 4/5, as it is for the native VBA.Format function.
' For rounding up or down, rounding of negative values can optionally be set to
' away-from-zero or towards-zero respectively by parameter RoundingAsAbsolute.

' Format placeholders:
'   f       foot value except zero.
'   F       foot value including zero.
'   i       inch value except zero.
'   I       inch value including zero.
'   r       fraction value except zero.
'   R       fraction value including zero.
'   '       foot unit.
'   "       inch unit.
'   ft      foot unit, short, spelled out.
'   in      inch unit, short, spelled out.
'   ft.     foot unit, short with dot, spelled out.
'   in.     inch unit, short with dot, spelled out.
'   foot    foot unit, long, spelled out.
'   inch    inch unit, long, spelled out.
'   /       fraction separator (divider)
'   <space> spacer.
'   -       dash.
'   \       escape character.
'
' Examples:
'   FormatFeetInches(17.222, 4)                         -> 1' 5-1/4"
'   FormatFeetInches(17.222, 4, , True)                 -> 1’ 5-1/4”    ' Smart Quotes.
'   FormatFeetInches(17.222, 4, "i-r")                  -> 17-1/4
'   FormatFeetInches(17.222, 4, "i-r""")                -> 17-1/4"
'   FormatFeetInches(17.222, 6, "i r""")                -> 17 7/32"
'   FormatFeetInches(7.222, 4, "f' i-r""")              -> 7-1/4"
'   FormatFeetInches(7.222, 4, "F' i-r""")              -> 0' 7-1/4"
'   FormatFeetInches(12.222, 4, "f' i-r""")             -> 1' 1/4"
'   FormatFeetInches(12.222, 4, "f' I-r""")             -> 1' 0-1/4"
'   FormatFeetInches(17.222, 0, "i-r""")                -> 17"
'   FormatFeetInches(17.222, 0, "i-R""")                -> 17-0/0"
'   FormatFeetInches(0.222, 2, "f' i-r""")              -> 1/4"
'   FormatFeetInches(0.222, 2, "F' i-r""")              -> 0' 1/4"
'   FormatFeetInches(12.222, 2, "f ft i r in")          -> 1 ft 1/4 in
'   FormatFeetInches(12.222, 0, "f ft i r in")          -> 1 ft
'   FormatFeetInches(12.222, 0, "f ft I r in")          -> 1 ft 0 in
'   FormatFeetInches(17.222, 2, "fft. I rin.")          -> 1 ft. 5 1/4 in.
'   FormatFeetInches(17.222, 2, "i r inches")           -> 17 1/4 inches
'   FormatFeetInches(1.222, 2, "i r inches")            -> 1 1/4 inches
'   FormatFeetInches(1.222, 0, "i r inches")            -> 1 inch
'   FormatFeetInches(17.222, 0, "i r inch")             -> 17 inches
'   FormatFeetInches(1.222, 0, "i r inch")              -> 1 inch
'   FormatFeetInches(27.222, 0, "f feet i r inches")    -> 2 feet 3 inches
'   FormatFeetInches(17.222, 0, "f feet i r inches")    -> 1 foot 5 inches
'   FormatFeetInches(7.222, 0, "F feet i r inches")     -> 0 feet 7 inches
'   FormatFeetInches(7.222, 2, "F feet i-r inches")     -> 0 feet 7-1/4 inches
'   FormatFeetInches(27.22, 6, "f foot and I r inch")   -> 2 feet and 3 7/32 inches
'
'   FormatFeetInches(17.222, 0, , , Up)                 -> 1' 6"
'   FormatFeetInches(17.222, 0, , , Down)               -> 1' 5"
'
' Also, se comments for the required functions:
'
'   ConvertDecimalFractions
'   RoundUpBase2
'   RoundMidBase2
'   RoundDownBase2
'
' 2018-04-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function FormatFeetInches( _
    ByVal Value As Variant, _
    Optional ByVal Exponent As Long, _
    Optional ByVal Format As String, _
    Optional ByVal SmartQuotes As Boolean, _
    Optional RoundingMethod As rmRoundingMethod = Midpoint, _
    Optional RoundingAsAbsolute As Boolean) _
    As String
    
    Const FootSymbol        As String = "f"
    Const InchSymbol        As String = "i"
    Const FractionSymbol    As String = "r"
    Const FootUnit          As String = "'"
    Const InchUnit          As String = """"
    Const FractionSeparator As String = "/"
    Const Spacer            As String = " "
    Const Dash              As String = "-"
    Const Escape            As String = "\"
    
    Const SingularFoot      As String = "foot"
    Const SingularInch      As String = "inch"
    Const PluralFoot        As String = "feet"
    Const PluralInch        As String = "inches"
    Const UniFoot           As String = "ft"
    Const UniInch           As String = "in"
    Const UniDotFoot        As String = UniFoot & "."
    Const UniDotInch        As String = UniInch & "."
        
    ' Default format: f' i-r"
    Const DefaultFormat     As String = FootSymbol & FootUnit & Spacer & InchSymbol & Dash & FractionSymbol & InchUnit
    
    Dim Numerator       As Long
    Dim Denominator     As Long
    Dim Feet            As Variant
    Dim AllInches       As Variant
    Dim Inches          As Variant
    
    Dim FootPart        As String
    Dim InchPart        As String
    Dim DashPart        As String
    Dim FractionPart    As String
    Dim FullPart        As String
    Dim Length          As Integer
    Dim Index           As Integer
    Dim Character       As String
    Dim LongFoot        As Boolean
    Dim LongInch        As Boolean
    Dim ShortFoot       As Boolean
    Dim ShortInch       As Boolean
    Dim ShortDotFoot    As Boolean
    Dim ShortDotInch    As Boolean

    If Not IsNumeric(Value) Then Exit Function
    
    If Format = "" Then
        Format = DefaultFormat
    End If
    ' Default spacer between integer inches and fraction of inches.
    DashPart = Spacer
    Length = Len(Format)
        
    ' Calculate the integer feet/inches and the fraction (remainder).
    AllInches = Fix(ConvertDecimalFractions(Value, Exponent, Numerator, Denominator, RoundingMethod, RoundingAsAbsolute))
    Feet = AllInches \ InchesPerFoot
    Inches = AllInches Mod InchesPerFoot
    
    ' Singularise spelled out long units.
    Format = Replace(Format, PluralFoot, SingularFoot)
    Format = Replace(Format, PluralInch, SingularInch)
    ' Temporarily replace all spelled out units with single character units.
    If InStr(1, Format, SingularFoot, vbTextCompare) > 1 Then
        LongFoot = True
        Format = Replace(Format, SingularFoot, FootUnit)
    ElseIf InStr(1, Format, UniDotFoot, vbTextCompare) > 1 Then
        ShortDotFoot = True
        Format = Replace(Format, UniDotFoot, FootUnit)
    ElseIf InStr(1, Format, UniFoot, vbTextCompare) > 1 Then
        ShortFoot = True
        Format = Replace(Format, UniFoot, FootUnit)
    End If
    If InStr(1, Format, SingularInch, vbTextCompare) > 1 Then
        LongInch = True
        Format = Replace(Format, SingularInch, InchUnit)
    ElseIf InStr(1, Format, UniDotInch, vbTextCompare) > 1 Then
        ShortDotInch = True
        Format = Replace(Format, UniDotInch, InchUnit)
    ElseIf InStr(1, Format, UniInch, vbTextCompare) > 1 Then
        ShortInch = True
        Format = Replace(Format, UniInch, InchUnit)
    End If
        
    ' Build parts.
    For Index = 1 To Length
        Character = Mid(Format, Index, 1)
        Select Case Character
            Case LCase(FootSymbol)
                If Feet > 0 Then
                    FootPart = CStr(Feet)
                Else
                    ' No display of feet.
                End If
            Case UCase(FootSymbol)
                ' Display any feet, even zero.
                FootPart = CStr(Feet)
            Case LCase(InchSymbol)
                If Inches > 0 Then
                    InchPart = CStr(Inches)
                Else
                    ' No display of inches.
                End If
            Case UCase(InchSymbol)
                ' Display any inches, even zero.
                InchPart = CStr(Inches)
            Case LCase(FractionSymbol)
                If Numerator > 0 Then
                    FractionPart = CStr(Numerator) & FractionSeparator & CStr(Denominator)
                Else
                    ' No display of fraction.
                End If
            Case UCase(FractionSymbol)
                ' Display any fraction, even when zero.
                FractionPart = CStr(Numerator) & FractionSeparator & CStr(Denominator)
            Case Dash
                ' Use a dash as spacer between integer inches and fraction of inches.
                DashPart = Dash
            Case Escape
                ' Skip the next character.
                Index = Index + 1
        End Select
    Next
    
    ' Adjust parts.
    If FootPart = "" Then
        If InchPart <> "" Then
            InchPart = CStr(AllInches)
        End If
    End If
    If InchPart = "" Or FractionPart = "" Then
        ' Not both integer inches and fraction of inches,
        ' thus no spacer between these.
        DashPart = ""
    End If
    
    ' Assemble parts.
    For Index = 1 To Length
        Character = Mid(Format, Index, 1)
        Select Case Character
            Case LCase(FootSymbol), UCase(FootSymbol)
                ' Append foot part.
                FullPart = FullPart & FootPart
            Case FootUnit
                ' Append foot unit if feet to display.
                If FootPart <> "" Then
                    ' Right-trim FullPart to remove space between value and unit.
                    FullPart = RTrim(FullPart) & FootUnit
                Else
                    ' No feet to display.
                End If
            Case LCase(InchSymbol), UCase(InchSymbol)
                ' Append inch part.
                FullPart = FullPart & InchPart
            Case InchUnit
                ' Append inch unit if inches to display.
                ' Right-trim FullPart to remove space between value and unit.
                If InchPart & FractionPart <> "" Then
                    FullPart = RTrim(FullPart) & InchUnit
                Else
                    ' No inches to display.
                End If
            Case LCase(FractionSymbol), UCase(FractionSymbol)
                ' Append fraction part.
                FullPart = FullPart & FractionPart
            Case Dash
                ' DashPart has been set in first loop.
                FullPart = FullPart & DashPart
            Case Spacer
                ' Right-trim FullPart to prevent double-spaces.
                FullPart = RTrim(FullPart) & Character
            Case Escape
                ' Skip this character and read the next literally.
                Index = Index + 1
                If Index <= Length Then
                    FullPart = FullPart & Mid(Format, Index, 1)
                End If
            Case Else
                ' Append any other character as is.
                FullPart = FullPart & Character
        End Select
    Next
    
    ' Restore spelled-out units.
    If LongFoot = True Then
        If Feet = 1 Then
            FullPart = Replace(FullPart, FootUnit, Spacer & SingularFoot)
        Else
            FullPart = Replace(FullPart, FootUnit, Spacer & PluralFoot)
        End If
    ElseIf ShortDotFoot = True Then
        FullPart = Replace(FullPart, FootUnit, Spacer & UniDotFoot)
    ElseIf ShortFoot = True Then
        FullPart = Replace(FullPart, FootUnit, Spacer & UniFoot)
    End If
    If LongInch = True Then
        If InchPart = "1" And Numerator = 0 Then
            FullPart = Replace(FullPart, InchUnit, Spacer & SingularInch)
        Else
            FullPart = Replace(FullPart, InchUnit, Spacer & PluralInch)
        End If
    ElseIf ShortDotInch = True Then
        FullPart = Replace(FullPart, InchUnit, Spacer & UniDotInch)
    ElseIf ShortInch = True Then
        FullPart = Replace(FullPart, InchUnit, Spacer & UniInch)
    End If
    
    If SmartQuotes = True Then
        ' Return output with "Smart Quotes".
        FullPart = Replace(FullPart, FootUnit, Chr(SmartSingleQuote))
        FullPart = Replace(FullPart, InchUnit, Chr(SmartDoubleQuote))
    End If
    
    FormatFeetInches = LTrim(FullPart)
    
End Function

' Parse a string for a value of feet and/or inches.
' The inch part can contain a fraction or be decimal.
' Returns the parsed values as decimal inches.
' For unparsable expressions, zero is returned.
'
' Maximum returned value is +/- 7922816299999618530273437599.
' Negative values will only be read as such, if the first
' non-space character is a minus sign.
'
' Smallest reliably parsed value is the fraction 1/2097152
' or the decimal value 0.000000476837158203125.
'
' Requires when not used in Access, for example Excel,
' either:
'   Module Access
' or a reference to Access, for example for Access 2016:
'   Microsoft Access 16.0 Object Library
'
' 2018-04-19. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ParseFeetInches( _
    ByVal Expression As String) _
    As Variant
    
    Dim ReplaceSets(20, 1)  As String
    Dim ExpressionParts     As Variant
    Dim ExpressionOneParts  As Variant

    Dim Sign                As Variant
    Dim DecimalInteger      As Variant
    Dim DecimalFraction     As Variant
    Dim DecimalInches       As Variant
    Dim Index               As Integer
    Dim Character           As String
    Dim FeetInches          As String
    Dim ExpressionOne       As String
    Dim ExpressionOneOne    As String
    Dim ExpressionOneTwo    As String
    Dim ExpressionTwo       As String
    Dim Numerator           As Long
    Dim Denominator         As Long

    ' Read sign.
    Sign = Sgn(Val(Expression))
    ' Trim double spacing.
    While InStr(Expression, "  ") > 0
        Expression = Replace(Expression, "  ", " ")
    Wend
    ' Replace foot units.
    ReplaceSets(0, 0) = "feet"
    ReplaceSets(0, 1) = "'"
    ReplaceSets(1, 0) = "foot"
    ReplaceSets(1, 1) = "'"
    ReplaceSets(2, 0) = "ft."
    ReplaceSets(2, 1) = "'"
    ReplaceSets(3, 0) = "ft"
    ReplaceSets(3, 1) = "'"
    ReplaceSets(4, 0) = Chr(SmartSingleQuote)   ' Smart Quote: "’"
    ReplaceSets(4, 1) = "'"
    ReplaceSets(5, 0) = " '"
    ReplaceSets(5, 1) = "'"
    ' Replace inch units.
    ReplaceSets(6, 0) = "inches"
    ReplaceSets(6, 1) = """"
    ReplaceSets(7, 0) = "inch."
    ReplaceSets(7, 1) = """"
    ReplaceSets(8, 0) = "inch"
    ReplaceSets(8, 1) = """"
    ReplaceSets(9, 0) = "in."
    ReplaceSets(9, 1) = """"
    ReplaceSets(10, 0) = "in"
    ReplaceSets(10, 1) = """"
    ReplaceSets(11, 0) = Chr(SmartDoubleQuote)  ' Smart Quote: "”"
    ReplaceSets(11, 1) = """"
    ReplaceSets(12, 0) = "''"
    ReplaceSets(12, 1) = """"
    ' Replace decimal separator.
    ReplaceSets(13, 0) = ","
    ReplaceSets(13, 1) = "."
    ' Replace units with operators.
    ReplaceSets(14, 0) = """"
    ReplaceSets(14, 1) = ""
    ReplaceSets(15, 0) = "'"
    ReplaceSets(15, 1) = "*" & CStr(InchesPerFoot) & " "
    ' Remove divider spaces.
    ReplaceSets(16, 0) = " /"
    ReplaceSets(16, 1) = "/"
    ReplaceSets(17, 0) = "/ "
    ReplaceSets(17, 1) = "/"
    ' Replace disturbing characters with neutral operator.
    ReplaceSets(18, 0) = " "
    ReplaceSets(18, 1) = "+"
    ReplaceSets(19, 0) = "-"
    ReplaceSets(19, 1) = "+"
    ReplaceSets(20, 0) = "+"
    ReplaceSets(20, 1) = "+0"
    
    ' Add leading neutral operator.
    Expression = "+0" & Expression
    ' Apply all replace sets.
    For Index = LBound(ReplaceSets, 1) To UBound(ReplaceSets, 1)
        Expression = Replace(Expression, ReplaceSets(Index, 0), ReplaceSets(Index, 1))
    Next
    ' Remove any useless or disturbing character.
    For Index = 1 To Len(Expression)
        Character = Mid(Expression, Index, 1)
        Select Case Character
            Case "0" To "9", "/", "+", "*", "."
                FeetInches = FeetInches & Character
        End Select
    Next
        
    ' For unparsable expressions, return 0.
    On Error GoTo Err_ParseFeetInches
    
    ExpressionParts = Split(FeetInches, "/")
    If UBound(ExpressionParts) = 0 Then
        ' FeetInches holds an integer part only, for example, "+00+038*12+0+05".
        ' Evaluate the cleaned expression as is.
        DecimalInches = Sign * CDec(Eval(FeetInches))
    Else
        ' FeetInches holds, for example, "+00+038*12+0+05+03/2048+0".
        ' For a maximum of decimals, split it into two parts:
        '   ExpressionOne = "+00+038*12+0+05+03"
        '   ExpressionTwo = "2048+0"
        ' or Eval would perform the calculation using Double only.
        ExpressionOne = ExpressionParts(0)
        ExpressionTwo = ExpressionParts(1)
        ' Split ExpressionOne into the integer part and the numerator part.
        ExpressionOneParts = Split(StrReverse(ExpressionOne), "+", 2)
        ' Retrieve the integer part and the numerator part.
        '   ExpressionOneOne = "+00+038*12+0+05"
        '   ExpressionOneTwo = "03"
        ExpressionOneOne = StrReverse(ExpressionOneParts(1))
        ExpressionOneTwo = StrReverse(ExpressionOneParts(0))
        
        ' Extract numerator and denominator.
        If Trim(ExpressionOneOne) = "" Then
            ' No integer expression is present.
            ' Use zero.
            ExpressionOneOne = "0"
        End If
        Numerator = Val(ExpressionOneTwo)
        Denominator = Val(ExpressionTwo)
        
        ' Evaluate the cleaned expression for the integer part.
        DecimalInteger = CDec(Eval(ExpressionOneOne))
        ' Calculate the fraction using CDec to obtain a maximum of decimals.
        If Denominator = 0 Then
            ' Cannot divide by zero.
            ' Return zero.
            DecimalFraction = CDec(0)
        Else
            DecimalFraction = CDec(Numerator) / CDec(Denominator)
        End If
        ' Sum and sign the integer part and the fraction part.
        DecimalInches = Sign * (DecimalInteger + DecimalFraction)
    End If
    
Exit_ParseFeetInches:
    ParseFeetInches = DecimalInches
    Exit Function
    
Err_ParseFeetInches:
    ' Ignore error and return zero.
    DecimalInches = CDec(0)
    Resume Exit_ParseFeetInches
    
End Function

' Converts a value for a measure in meters to inches.
' Returns 0 (zero) for invalid inputs.
'
' Will convert any value within the range of Decimal
' with the precision of Decimal.
' Converts values exceeding the range of Decimal as
' Double.
' Largest value with full 28-digit precision is 1E+27
' Smallest value with full 28-digit precision is 1E-26
'
' Examples:
'   Meter = 4.0
'   Inch = InchMeter(Meter)
'   Inch -> 157.48031496062992125984251969
'
'   Meter = 2.54
'   Inch = InchMeter(Meter)
'   Inch -> 100.0
'
' 2018-04-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function InchMeter( _
    ByVal Value As Variant) _
    As Variant

    Dim Result  As Variant
    
    If IsNumeric(Value) Then
        On Error Resume Next
        Result = CDec(Value) / MetersPerInch
        If Err.Number <> 0 Then
            ' Decimal overflow.
            ' Calculate without conversion to Decimal.
            Result = CDbl(Value) / MetersPerInch
        End If
    Else
        Result = 0
    End If
    
    InchMeter = Result
    
End Function

' Converts a value for a measure in inches to meters.
' Returns 0 (zero) for invalid inputs.
'
' Will convert any value within the range of Decimal
' with the precision of Decimal.
' Converts values exceeding the range of Decimal as
' Double.
'
' Largest value with full 28-digit precision is 1E+26
' Smallest value with full 28-digit precision is 1E-24
'
' Examples:
'   Inch = 40.0
'   Meter = MeterInch(Inch)
'   Meter -> 1.016
'
'   Inch = 1 / MetersPerInch            ' Double.
'   Inch -> 39.3700787401575
'   Meter = MeterInch(Inch)
'   Meter -> 1.0000000000000005
'
'   Inch = CDec(1) / MetersPerInch      ' Decimal.
'   Inch -> 39.370078740157480314960629921
'   Meter = MeterInch(Inch)
'   Meter -> 1.0
'
' 2018-04-09. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function MeterInch( _
    ByVal Value As Variant) _
    As Variant

    Dim Result  As Variant
    
    If IsNumeric(Value) Then
        On Error Resume Next
        Result = CDec(Value) * MetersPerInch
        If Err.Number <> 0 Then
            ' Decimal overflow.
            ' Calculate without conversion to Decimal.
            Result = CDbl(Value) * MetersPerInch
        End If
    Else
        Result = 0
    End If
    
    MeterInch = Result
    
End Function


