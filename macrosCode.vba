'Variable declaration. _
 The following array variable is required for the "AssignGridLetters" function.

Dim GridLetters(6, 12) As String


Function GridSquare(Eastings, Northings, SquareSize)
'Input easting and northings (in meters), and a square size (in kilometers). _
 Return the appropriate National Grid reference including 100km grid letters.
 
'REQUIRES THE "AssignGridLetters" FUNCTION

'REQUIRES the GridLetters(6, 12) array declared as a String

'Test that input coordinates are within the National Grid
    If Eastings < 0 Or Eastings >= 700000 Then
        GridSquare = "Easting incorrect <0 or >=700000"
        GoTo EndOfFunction
    End If
    
    If Northings < 0 Or Northings >= 1300000 Then
        GridSquare = "Northings incorrect <0 or >=1300000"
        GoTo EndOfFunction
    End If

'Convert Easting and Northing to text values with format: _
 "000000" (6 digits) for Eastings and _
 "0000000" (7 digits) for Northings.
    EastText = Format(Int(Eastings), "000000")
    NorthText = Format(Int(Northings), "0000000")
    
'Assign 100km square letters to GridLetters array variable
    AssignGridLetters
    
'Compute GridLetters array reference
    X = Val(Left(EastText, 1))
    Y = Val(Left(NorthText, 2))
    
'Format output according to value of SquareSize
    Select Case SquareSize
        Case 100 '100km square.
            'Output letters only.
            GridSquare = GridLetters(X, Y)
        Case 10 '10km square.
            'Output letters, 2nd EastText digit and 3rd NorthText digit
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 1) + Mid(NorthText, 3, 1)
        Case 5 '5km square
            'Calculate quadrant within 10km square
            If Val(Mid(EastText, 3, 1)) < 5 And Val(Mid(NorthText, 4, 1)) < 5 Then
                Quadrant = "SW"
            ElseIf Val(Mid(EastText, 3, 1)) < 5 And Val(Mid(NorthText, 4, 1)) >= 5 Then
                Quadrant = "NW"
            ElseIf Val(Mid(EastText, 3, 1)) >= 5 And Val(Mid(NorthText, 4, 1)) < 5 Then
                Quadrant = "SE"
            ElseIf Val(Mid(EastText, 3, 1)) >= 5 And Val(Mid(NorthText, 4, 1)) >= 5 Then
                Quadrant = "NE"
            End If
            'Output letters, 2nd EastText digit, 3rd NorthText digit and Quadrant
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 1) + Mid(NorthText, 3, 1) + Quadrant
        Case 1 '1km square
            'Output letters, 2nd-3rd EastText digits and 3rd-4th NorthText digits
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 2) + Mid(NorthText, 3, 2)
        Case 0.5 '500m square
            'Calculate quadrant within 1km square
            If Val(Mid(EastText, 4, 1)) < 5 And Val(Mid(NorthText, 5, 1)) < 5 Then
                Quadrant = "SW"
            ElseIf Val(Mid(EastText, 4, 1)) < 5 And Val(Mid(NorthText, 5, 1)) >= 5 Then
                Quadrant = "NW"
            ElseIf Val(Mid(EastText, 4, 1)) >= 5 And Val(Mid(NorthText, 5, 1)) < 5 Then
                Quadrant = "SE"
            ElseIf Val(Mid(EastText, 4, 1)) >= 5 And Val(Mid(NorthText, 5, 1)) >= 5 Then
                Quadrant = "NE"
            End If
            'Output letters, 2nd-3rd EastText digits, 3rd-4th NorthText digits and Quadrant
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 2) + Mid(NorthText, 3, 2) + Quadrant
        Case 0.1 '100m square
            'Output letters, 2nd-4th EastText digits and 3rd-5th NorthText digits
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 3) + Mid(NorthText, 3, 3)
        Case 0.01 '10m square
            'Output letters, 2nd-5th EastText digits and 3rd-5th NorthText digits
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 4) + Mid(NorthText, 3, 4)
        Case 0.001 '1m square
            'Output letters, 2nd-6th EastText digits and 3rd-7th NorthText digits
            GridSquare = GridLetters(X, Y) + Mid(EastText, 2, 5) + Mid(NorthText, 3, 5)
    
    End Select

EndOfFunction:
End Function

Function GridRef_to_Eastings(GridRef As String)
'Convert a National Grid Reference, with 100km grid letters to full Eastings. _
 E.g. SE2720 (a 1km square) returns 427000.
 
'REQUIRES THE "AssignGridLetters" FUNCTION

'REQUIRES the GridLetters(6, 12) array declared as a String

'Convert GridRef to upper case just in case it isn't already
GridRef = UCase(GridRef)
 
'Check that 1st 2 characters of GridRef are valid
FirstCHR = Left(GridRef, 1)

If FirstCHR <> "H" And FirstCHR <> "J" And FirstCHR <> "N" _
And FirstCHR <> "O" And FirstCHR <> "S" And FirstCHR <> "T" Then
    GridRef_to_Eastings = "Invalid Reference"
    GoTo EndOfFunction
End If

SecondCHR = Mid(GridRef, 2, 1)

If Asc(SecondCHR) < 65 Or Asc(SecondCHR) > 90 Or SecondCHR = "I" Then
    GridRef_to_Eastings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Check that GridRef is a valid length and has an even number of characters
If Len(GridRef) < 2 Or Len(GridRef) > 12 Or Len(GridRef) Mod 2 <> 0 Then
    GridRef_to_Eastings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Check to see if GridRef has a quadrant
QuadrantExists = False
If Len(GridRef) >= 6 Then
    Quadrant = Right(GridRef, 2)
    If Quadrant = "SW" Or Quadrant = "NW" Or Quadrant = "NE" Or Quadrant = "SE" Then
        QuadrantExists = True
    End If
End If

'Extract numbers from GridRef
If QuadrantExists = False And Len(GridRef) > 2 Then
    NumbersAsText = Mid(GridRef, 3, Len(GridRef) - 2)
End If
If QuadrantExists = True Then
    NumbersAsText = Mid(GridRef, 3, Len(GridRef) - 4)
End If

'Check that extracted numbers are actually numbers
If IsNumeric(Format(NumbersAsText, "0")) = False And Len(GridRef) > 2 Then
    GridRef_to_Eastings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Build 100km square ref from first two letters
HundredKmSquare = FirstCHR & SecondCHR

'Assign 100km square letters to GridLetters array variable
AssignGridLetters

'Look up Letters in GridLetters array to get first Eastings digit
LettersMatch = False
For X = 0 To 6
    For Y = 0 To 12
        If GridLetters(X, Y) = HundredKmSquare Then
            EastingsFirstDigit = X
            LettersMatch = True
            Exit For
        End If
    Next Y
Next X

'Check for valid return from Letters search
If LettersMatch = False Then
    GridRef_to_Eastings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Format output according to square size being referred to (Len(GridRef))
Select Case Len(GridRef)
    Case 2 '100 km square
        GridRef_to_Eastings = Val(EastingsFirstDigit) * 100000
    Case 4 '10km square
        'Extract first half of NumbersAsText
        Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings) * 10000
    Case 6 '5km or 1km square
        If QuadrantExists = True Then '5km square
            'Calculate quadrant digit
            If Quadrant = "SW" Or Quadrant = "NW" Then
                QuadrantDigit = "0"
            ElseIf Quadrant = "SE" Or Quadrant = "NE" Then
                QuadrantDigit = "5"
            End If
            'Extract first half of NumbersAsText
            Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings & QuadrantDigit) * 1000
        ElseIf QuadrantExists = False Then '1km square
            'Extract first half of NumbersAsText
            Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings) * 1000
        End If
    Case 8 '500m or 100m square
        If QuadrantExists = True Then '500m square
            'Calculate quadrant digit
            If Quadrant = "SW" Or Quadrant = "NW" Then
                QuadrantDigit = "0"
            ElseIf Quadrant = "SE" Or Quadrant = "NE" Then
                QuadrantDigit = "5"
            End If
            'Extract first half of NumbersAsText
            Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings & QuadrantDigit) * 100
        ElseIf QuadrantExists = False Then '100m square
            'Extract first half of NumbersAsText
            Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings) * 100
        End If
    Case 10 '10m square
        'Extract first half of NumbersAsText
        Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings) * 10
    Case 12 '1m square
        'Extract first half of NumbersAsText
        Eastings = Left(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Eastings = Val(EastingsFirstDigit & Eastings)
End Select

EndOfFunction:
End Function

Function GridRef_to_Northings(GridRef As String)
'Convert a National Grid Reference, with 100km grid letters to full Northings. _
 E.g. SE2720 (a 1km square) returns 420000.
  
'REQUIRES THE "AssignGridLetters" FUNCTION

'REQUIRES the GridLetters(6, 12) array declared as a String

'Convert GridRef to upper case just in case it isn't already
GridRef = UCase(GridRef)
 
'Check that 1st 2 characters of GridRef are valid
FirstCHR = Left(GridRef, 1)

If FirstCHR <> "H" And FirstCHR <> "J" And FirstCHR <> "N" _
And FirstCHR <> "O" And FirstCHR <> "S" And FirstCHR <> "T" Then
    GridRef_to_Northings = "Invalid Reference"
    GoTo EndOfFunction
End If

SecondCHR = Mid(GridRef, 2, 1)

If Asc(SecondCHR) < 65 Or Asc(SecondCHR) > 90 Or SecondCHR = "I" Then
    GridRef_to_Northings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Check that GridRef is a valid length and has an even number of characters
If Len(GridRef) < 2 Or Len(GridRef) > 12 Or Len(GridRef) Mod 2 <> 0 Then
    GridRef_to_Northings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Check to see if GridRef has a quadrant
QuadrantExists = False
If Len(GridRef) >= 6 Then
    Quadrant = Right(GridRef, 2)
    If Quadrant = "SW" Or Quadrant = "NW" Or Quadrant = "NE" Or Quadrant = "SE" Then
        QuadrantExists = True
    End If
End If

'Extract numbers from GridRef
If QuadrantExists = False And Len(GridRef) > 2 Then
    NumbersAsText = Mid(GridRef, 3, Len(GridRef) - 2)
End If
If QuadrantExists = True Then
    NumbersAsText = Mid(GridRef, 3, Len(GridRef) - 4)
End If

'Check that extracted numbers are actually numbers
If IsNumeric(Format(NumbersAsText, "0")) = False And Len(GridRef) > 2 Then
    GridRef_to_Northings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Build 100km square ref from first two letters
HundredKmSquare = FirstCHR & SecondCHR

'Assign 100km square letters to GridLetters array variable
AssignGridLetters

'Look up Letters in GridLetters array to get first Northings digit
LettersMatch = False
For X = 0 To 6
    For Y = 0 To 12
        If GridLetters(X, Y) = HundredKmSquare Then
            NorthingsFirstDigit = Y
            LettersMatch = True
            Exit For
        End If
    Next Y
Next X

'Check for valid return from Letters search
If LettersMatch = False Then
    GridRef_to_Northings = "Invalid Reference"
    GoTo EndOfFunction
End If

'Format output according to square size being referred to (Len(GridRef))
Select Case Len(GridRef)
    Case 2 '100 km square
        GridRef_to_Northings = Val(NorthingsFirstDigit) * 100000
    Case 4 '10km square
        'Extract second half of NumbersAsText
        Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Northings = Val(NorthingsFirstDigit & Northings) * 10000
    Case 6 '5km or 1km square
        If QuadrantExists = True Then '5km square
            'Calculate quadrant digit
            If Quadrant = "SW" Or Quadrant = "SE" Then
                QuadrantDigit = "0"
            ElseIf Quadrant = "NW" Or Quadrant = "NE" Then
                QuadrantDigit = "5"
            End If
            'Extract first half of NumbersAsText
            Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Northings = Val(NorthingsFirstDigit & Northings & QuadrantDigit) * 1000
        ElseIf QuadrantExists = False Then '1km square
            'Extract first half of NumbersAsText
            Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Northings = Val(NorthingsFirstDigit & Northings) * 1000
        End If
    Case 8 '500m or 100m square
        If QuadrantExists = True Then '500m square
            'Calculate quadrant digit
            If Quadrant = "SW" Or Quadrant = "SE" Then
                QuadrantDigit = "0"
            ElseIf Quadrant = "NW" Or Quadrant = "NE" Then
                QuadrantDigit = "5"
            End If
            'Extract first half of NumbersAsText
            Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Northings = Val(NorthingsFirstDigit & Northings & QuadrantDigit) * 100
        ElseIf QuadrantExists = False Then '100m square
            'Extract first half of NumbersAsText
            Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
            GridRef_to_Northings = Val(NorthingsFirstDigit & Northings) * 100
        End If
    Case 10 '10m square
        'Extract first half of NumbersAsText
        Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Northings = Val(NorthingsFirstDigit & Northings) * 10
    Case 12 '1m square
        'Extract first half of NumbersAsText
        Northings = Right(NumbersAsText, Len(NumbersAsText) / 2)
        GridRef_to_Northings = Val(NorthingsFirstDigit & Northings)
End Select

EndOfFunction:
End Function

Function AssignGridLetters()
'Assign 100km square letters to GridLetters array variable

'REQUIRES the GridLetters(6, 12) array declared as a String
    
    GridLetters(0, 0) = "SV"
    GridLetters(0, 1) = "SQ"
    GridLetters(0, 2) = "SL"
    GridLetters(0, 3) = "SF"
    GridLetters(0, 4) = "SA"
    GridLetters(0, 5) = "NV"
    GridLetters(0, 6) = "NQ"
    GridLetters(0, 7) = "NL"
    GridLetters(0, 8) = "NF"
    GridLetters(0, 9) = "NA"
    GridLetters(0, 10) = "HV"
    GridLetters(0, 11) = "HQ"
    GridLetters(0, 12) = "HL"
    GridLetters(1, 0) = "SW"
    GridLetters(1, 1) = "SR"
    GridLetters(1, 2) = "SM"
    GridLetters(1, 3) = "SG"
    GridLetters(1, 4) = "SB"
    GridLetters(1, 5) = "NW"
    GridLetters(1, 6) = "NR"
    GridLetters(1, 7) = "NM"
    GridLetters(1, 8) = "NG"
    GridLetters(1, 9) = "NB"
    GridLetters(1, 10) = "HW"
    GridLetters(1, 11) = "HR"
    GridLetters(1, 12) = "HM"
    GridLetters(2, 0) = "SX"
    GridLetters(2, 1) = "SS"
    GridLetters(2, 2) = "SN"
    GridLetters(2, 3) = "SH"
    GridLetters(2, 4) = "SC"
    GridLetters(2, 5) = "NX"
    GridLetters(2, 6) = "NS"
    GridLetters(2, 7) = "NN"
    GridLetters(2, 8) = "NH"
    GridLetters(2, 9) = "NC"
    GridLetters(2, 10) = "HX"
    GridLetters(2, 11) = "HS"
    GridLetters(2, 12) = "HN"
    GridLetters(3, 0) = "SY"
    GridLetters(3, 1) = "ST"
    GridLetters(3, 2) = "SO"
    GridLetters(3, 3) = "SJ"
    GridLetters(3, 4) = "SD"
    GridLetters(3, 5) = "NY"
    GridLetters(3, 6) = "NT"
    GridLetters(3, 7) = "NO"
    GridLetters(3, 8) = "NJ"
    GridLetters(3, 9) = "ND"
    GridLetters(3, 10) = "HY"
    GridLetters(3, 11) = "HT"
    GridLetters(3, 12) = "HO"
    GridLetters(4, 0) = "SZ"
    GridLetters(4, 1) = "SU"
    GridLetters(4, 2) = "SP"
    GridLetters(4, 3) = "SK"
    GridLetters(4, 4) = "SE"
    GridLetters(4, 5) = "NZ"
    GridLetters(4, 6) = "NU"
    GridLetters(4, 7) = "NP"
    GridLetters(4, 8) = "NK"
    GridLetters(4, 9) = "NE"
    GridLetters(4, 10) = "HZ"
    GridLetters(4, 11) = "HU"
    GridLetters(4, 12) = "HP"
    GridLetters(5, 0) = "TV"
    GridLetters(5, 1) = "TQ"
    GridLetters(5, 2) = "TL"
    GridLetters(5, 3) = "TF"
    GridLetters(5, 4) = "TA"
    GridLetters(5, 5) = "OV"
    GridLetters(5, 6) = "OQ"
    GridLetters(5, 7) = "OL"
    GridLetters(5, 8) = "OF"
    GridLetters(5, 9) = "OA"
    GridLetters(5, 10) = "JV"
    GridLetters(5, 11) = "JQ"
    GridLetters(5, 12) = "JL"
    GridLetters(6, 0) = "TW"
    GridLetters(6, 1) = "TR"
    GridLetters(6, 2) = "TM"
    GridLetters(6, 3) = "TG"
    GridLetters(6, 4) = "TB"
    GridLetters(6, 5) = "OW"
    GridLetters(6, 6) = "OR"
    GridLetters(6, 7) = "OM"
    GridLetters(6, 8) = "OG"
    GridLetters(6, 9) = "OB"
    GridLetters(6, 10) = "JW"
    GridLetters(6, 11) = "JR"
    GridLetters(6, 12) = "JM"

End Function
