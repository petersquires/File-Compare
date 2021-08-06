Imports System.IO

Public Module Code
    Public BYTE0, BYTE1, BYTE2, BYTE3, BYTE4, BYTE5, BYTE6, BYTE7, BYTE8, BYTE9, BYTEA, BYTEB, BYTEC, BYTED, BYTEE, BYTEF, BYTE10 As String  'TWO HEX CHARS.
    Public PARENTPATH, SOURCEPATH1, SOURCEPATH2, FILE1, FILE2 As String
    Public CURRENT_LINE, SYNCHBRECORD As String
    Public NUM1, NUM2, NUM3, NUM4 As Integer   'for debug use

    Public DISPLAYMAXCHARS, DISPLAYLINELENGTH, LISTSIZE, CLICKLINE As Integer
    Public SCORECOLOURS() As Brush
    Public DISPLAYLINES(), DISPLAYLINES1(), DISPLAYLINES2(), DEBUGPRINTLINES(), LINEtoDISPLAY, WRITEFILE() As String
    Public GREENLINES, BROWNLINES, GREENLINES1, BROWNLINES1, GREENLINES2, BROWNLINES2, SHADE As String
    Public REDLINES, REDLINES1, REDLINES2, BLUELINES, BLUELINES1, BLUELINES2 As String
    Public DISPLAYOBJECT1, DISPLAYOBJECT2 As Display
    Public PRINTDEBUG, STORERESULTS As Boolean
    Public REACH, SIZEOFBLOCK As Integer

    Public FIRSTSELECTED, SECONDSELECTED, SELECTEDFILE1, SELECTEDFILE2 As String
    Public CONTENTS1(), CONTENTS2(), ORIGINAL1(), ORIGINAL2(), SHADOW1(), SHADOW2(), LINENUM1(), LINENUM2() As String
    Public a1, INDEX1, INDEX2, INDEX1END, INDEX2END As Integer
    Public DISPLAYDATAFILE, FILE1DATE, FILE2DATE As String
    Public IGNORESPACES, SHOWTABS, SKIPBLANKLINES, IGNORETABS, SYNCHB_USED As Boolean
    Private SynchWorked As Boolean
    Private UsedReach As Integer

    'Public DISPLAYLINEINDEX As Integer



    Public Sub StripToTab()  'strips Current_Line up to a Tab or end of line.
        'Tab is stripped, and Byte0 holds the text before the tab. Remaining leading tab in Current_Line is stripped.
        Dim localPos As Integer
        BYTE0 = ""  'default
        If Len(CURRENT_LINE) = 0 Then Exit Sub
        localPos = InStr(1, CURRENT_LINE, Chr(9), vbBinaryCompare)
        If localPos = 0 Then
            BYTE0 = CURRENT_LINE
            CURRENT_LINE = ""
            Exit Sub
        Else
            BYTE0 = Strings.Left(CURRENT_LINE, localPos - 1)
            CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - localPos)
        End If

    End Sub
    Public Sub StripToComma()   'strips Current_Line up to a comma or end of line.
        'Comma is stripped, and Byte0 holds the text before the comma.
        Dim localPos As Integer
        BYTE0 = ""  'default
        If Len(CURRENT_LINE) = 0 Then Exit Sub
        localPos = InStr(1, CURRENT_LINE, ",", vbBinaryCompare)
        If localPos = 0 Then
            BYTE0 = CURRENT_LINE
            CURRENT_LINE = ""
            Exit Sub
        Else
            BYTE0 = Strings.Left(CURRENT_LINE, localPos - 1)
            CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - localPos)
        End If
    End Sub
    Public Sub DoDummyPrint()
        Dim localLoop, fileNum As Integer

        fileNum = FreeFile()
        FileOpen(fileNum, PARENTPATH & "\" & "DEBUGPrint.txt", OpenMode.Output)
        For localLoop = 1 To UBound(DEBUGPRINTLINES)
            Print(fileNum, DEBUGPRINTLINES(localLoop))
        Next localLoop
        FileClose(fileNum)
    End Sub
    Public Function SPEAK(SoundName As String) As Boolean 'if sound off, returns true but makes no sound.
        'if sound is beep, returns false but makes no sound.
        'if sound is on, attempts to speak.
        SPEAK = True

        Select Case UCase(SoundName)
            Case "DING"
                My.Computer.Audio.Play(My.Resources.DING, AudioPlayMode.Background)
            Case "CHORD"
                My.Computer.Audio.Play(My.Resources.CHORD, AudioPlayMode.Background)
            Case Else
                Beep()
                SPEAK = False
        End Select
endOfSpeak:
    End Function
    Public Sub ADDBLUELINE1()
        SHADE = "BLUE"
        Call ADD_DISPLAYLINES1()
        SHADE = ""
    End Sub
    Public Sub ADDBLUELINE2()
        SHADE = "BLUE"
        Call ADD_DISPLAYLINES2()
        SHADE = ""
    End Sub
    Public Sub ADDREDLINE1()
        SHADE = "RED"
        Call ADD_DISPLAYLINES2()
        SHADE = ""
    End Sub
    Public Sub ADDREDLINE2()
        SHADE = "RED"
        Call ADD_DISPLAYLINES2()
        SHADE = ""
    End Sub
    Public Sub ADDGREENLINE1()
        SHADE = "GREEN"
        Call ADD_DISPLAYLINES1()
        SHADE = ""
    End Sub
    Public Sub ADDBROWNLINE1()
        SHADE = "BROWN"
        Call ADD_DISPLAYLINES1()
        SHADE = ""
    End Sub
    Public Sub ADDGREENLINE2()
        SHADE = "GREEN"
        Call ADD_DISPLAYLINES2()
        SHADE = ""
    End Sub
    Public Sub ADDBROWNLINE2()
        SHADE = "BROWN"
        Call ADD_DISPLAYLINES2()
        SHADE = ""
    End Sub
    Public Sub ADD_DISPLAYLINES1()   'adds current_line to display stack, after adding a filler and line end. 
        Dim fiLLer, tabFreeLine, localLine, localChar As String
        Dim localLength, localNum As Integer
        Dim localFlag As Boolean

        Try
            If UBound(DISPLAYLINES1) = Nothing Then
                ReDim DISPLAYLINES1(0)
                BLUELINES1 = ","
                REDLINES1 = ","
                GREENLINES1 = ","
                BROWNLINES1 = ","
            End If
        Catch ex As Exception
            ReDim DISPLAYLINES1(0)
        End Try
        fiLLer = "   "
        'replace tabs in CURRENT_LINE with a single space.
        'If Len(CURRENT_LINE) > 0 Then
        localLine = CURRENT_LINE    'hold it.
        If LINEtoDISPLAY <> "" Then CURRENT_LINE = LINEtoDISPLAY

        localLength = 0

        GoTo SkipTabRemove      'debug
        tabFreeLine = ""
        Do While Len(CURRENT_LINE) > 0  'Replace tabs with a single space.
            localChar = Strings.Left(CURRENT_LINE, 1)
            If localChar = Chr(9) Then localChar = " "
            tabFreeLine = tabFreeLine & localChar
            localLength = localLength + 1
            CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - 1)
        Loop
        CURRENT_LINE = tabFreeLine
        If localLength > DISPLAYLINELENGTH Then DISPLAYLINELENGTH = localLength

SkipTabRemove:

        localFlag = False
        If Strings.Right(CURRENT_LINE, 1) = Chr(10) Then
            CURRENT_LINE = Strings.Left(CURRENT_LINE, Len(CURRENT_LINE) - 2)    'remove the extra crlf
            localFlag = True
        End If
        ReDim Preserve DISPLAYLINES1(UBound(DISPLAYLINES1) + 1)
        DISPLAYLINES1(UBound(DISPLAYLINES1)) = CURRENT_LINE & fiLLer & vbCrLf
        If SHADE = "BLUE" Then
            localNum = UBound(DISPLAYLINES1)
            BLUELINES1 = BLUELINES1 & localNum.ToString & ","
        ElseIf SHADE = "RED" Then
            localNum = UBound(DISPLAYLINES1)
            REDLINES1 = REDLINES1 & localNum.ToString & ","
        ElseIf SHADE = "GREEN" Then
            localNum = UBound(DISPLAYLINES1)
            GREENLINES1 = GREENLINES1 & localNum.ToString & ","
        ElseIf SHADE = "BROWN" Then
            localNum = UBound(DISPLAYLINES1)
            BROWNLINES1 = BROWNLINES1 & localNum.ToString & ","
        End If
        SHADE = ""
        If localFlag Then
            ReDim Preserve DISPLAYLINES1(UBound(DISPLAYLINES1) + 1)
            DISPLAYLINES1(UBound(DISPLAYLINES1)) = vbCrLf
            CURRENT_LINE = CURRENT_LINE & vbCrLf 'restore, in case.
        End If
        If LINEtoDISPLAY <> "" Then
            CURRENT_LINE = localLine    'restore it.
            LINEtoDISPLAY = ""
        End If
    End Sub
    Public Sub ADD_DISPLAYLINES2()   'adds current_line to display stack, after adding a filler and line end. 
        Dim fiLLer, tabFreeLine, localLine, localChar As String
        Dim localLength, localNum As Integer
        Dim localFlag As Boolean

        Try
            If UBound(DISPLAYLINES2) = Nothing Then
                ReDim DISPLAYLINES2(0)
                BLUELINES2 = ","
                REDLINES2 = ","
                GREENLINES2 = ","
                BROWNLINES2 = ","
            End If
        Catch ex As Exception
            ReDim DISPLAYLINES2(0)
        End Try
        fiLLer = "   "
        'replace tabs in CURRENT_LINE with a single space.
        'If Len(CURRENT_LINE) > 0 Then
        localLine = CURRENT_LINE    'hold it.
        If LINEtoDISPLAY <> "" Then CURRENT_LINE = LINEtoDISPLAY
        localLength = 0

        GoTo SkipTabRemove  'debug
        tabFreeLine = ""
        Do While Len(CURRENT_LINE) > 0  'Replace tabs with a single space.
            localChar = Strings.Left(CURRENT_LINE, 1)
            If localChar = Chr(9) Then localChar = " "
            tabFreeLine = tabFreeLine & localChar
            localLength = localLength + 1
            CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - 1)
        Loop
        CURRENT_LINE = tabFreeLine
        If localLength > DISPLAYLINELENGTH Then DISPLAYLINELENGTH = localLength

SkipTabRemove:

        localFlag = False
        If Strings.Right(CURRENT_LINE, 1) = Chr(10) Then
            CURRENT_LINE = Strings.Left(CURRENT_LINE, Len(CURRENT_LINE) - 2)    'remove the extra crlf
            localFlag = True
        End If
        ReDim Preserve DISPLAYLINES2(UBound(DISPLAYLINES2) + 1)
        DISPLAYLINES2(UBound(DISPLAYLINES2)) = CURRENT_LINE & fiLLer & vbCrLf
        If SHADE = "BLUE" Then
            localNum = UBound(DISPLAYLINES2)
            BLUELINES2 = BLUELINES2 & localNum.ToString & ","
        ElseIf SHADE = "RED" Then
            localNum = UBound(DISPLAYLINES2)
            REDLINES2 = REDLINES2 & localNum.ToString & ","
        ElseIf SHADE = "GREEN" Then
            localNum = UBound(DISPLAYLINES2)
            GREENLINES2 = GREENLINES2 & localNum.ToString & ","
        ElseIf SHADE = "BROWN" Then
            localNum = UBound(DISPLAYLINES2)
            BROWNLINES2 = BROWNLINES2 & localNum.ToString & ","
        End If
        SHADE = ""
        If localFlag Then
            ReDim Preserve DISPLAYLINES2(UBound(DISPLAYLINES2) + 1)
            DISPLAYLINES2(UBound(DISPLAYLINES2)) = vbCrLf
            CURRENT_LINE = CURRENT_LINE & vbCrLf 'restore, in case.
        End If
        If LINEtoDISPLAY <> "" Then
            CURRENT_LINE = localLine    'restore it.
            LINEtoDISPLAY = ""
        End If
    End Sub
    Public Sub SHOWDISPLAY()
        Dim localLoop, lineLength As Integer
        Dim tempLine As String

        If DISPLAYLINES1 Is Nothing Then GoTo SecondDisplay
        If UBound(DISPLAYLINES1) < 2 Then
            ReDim Preserve DISPLAYLINES1(UBound(DISPLAYLINES1) + 1)
            DISPLAYLINES1(UBound(DISPLAYLINES1)) = "No mismatches found in " & UBound(CONTENTS1).ToString & " lines."
        End If
        BLUELINES = BLUELINES1
        REDLINES = REDLINES1
        GREENLINES = GREENLINES1
        BROWNLINES = BROWNLINES1
        'REDLINES = ","
        lineLength = 1
        For localLoop = 1 To UBound(DISPLAYLINES1)
            If Len(DISPLAYLINES1(localLoop)) > lineLength Then
                lineLength = Len(DISPLAYLINES1(localLoop))
            End If

            If localLoop <> UBound(DISPLAYLINES1) Then
            Else
                tempLine = DISPLAYLINES1(UBound(DISPLAYLINES1))   'last  line
                If Len(tempLine) > 2 Then tempLine = Strings.Left(tempLine, Len(tempLine) - 2) 'strip off last linefeed
            End If
        Next localLoop
        DISPLAYLINES = DISPLAYLINES1
        DISPLAYMAXCHARS = lineLength + 1
        DISPLAYLINELENGTH = DISPLAYMAXCHARS
        DISPLAYDATAFILE = "\DisplayOptions1.txt"
        'DISPLAYLINEINDEX = 1
        If DISPLAYOBJECT1 Is Nothing Then
            DISPLAYOBJECT1 = New Display
            DISPLAYOBJECT1.LineCount = UBound(DISPLAYLINES1)
            DISPLAYOBJECT1.Show()
            Call DISPLAYOBJECT1.UpdateContent()
        Else
            'DISPLAYLINELENGTH = DISPLAYOBJECT1.ZoomLineLength    'preserve zoom
            DISPLAYOBJECT1.LineCount = UBound(DISPLAYLINES1)
            Call DISPLAYOBJECT1.UpdateContent()
        End If

SecondDisplay:
        If DISPLAYLINES2 Is Nothing Then Exit Sub
        If UBound(DISPLAYLINES2) < 2 Then
            ReDim Preserve DISPLAYLINES2(UBound(DISPLAYLINES2) + 1)
            DISPLAYLINES2(UBound(DISPLAYLINES2)) = "No mismatches found in " & UBound(CONTENTS2).ToString & " lines."
        End If
        BLUELINES = BLUELINES2
        REDLINES = REDLINES2
        GREENLINES = GREENLINES2
        BROWNLINES = BROWNLINES2
        'BLUELINES = ","
        lineLength = 1
        For localLoop = 1 To UBound(DISPLAYLINES2)
            If Len(DISPLAYLINES2(localLoop)) > lineLength Then
                lineLength = Len(DISPLAYLINES2(localLoop))
            End If

            If localLoop <> UBound(DISPLAYLINES2) Then
            Else
                tempLine = DISPLAYLINES2(UBound(DISPLAYLINES2))   'last  line
                If Len(tempLine) > 2 Then tempLine = Strings.Left(tempLine, Len(tempLine) - 2) 'strip off last linefeed
            End If
        Next localLoop
        DISPLAYLINES = DISPLAYLINES2
        DISPLAYMAXCHARS = lineLength + 1
        DISPLAYLINELENGTH = DISPLAYMAXCHARS
        DISPLAYDATAFILE = "\DisplayOptions2.txt"
        'DISPLAYLINEINDEX = 1
        If DISPLAYOBJECT2 Is Nothing Then
            DISPLAYOBJECT2 = New Display
            DISPLAYOBJECT2.LineCount = UBound(DISPLAYLINES2)
            DISPLAYOBJECT2.Show()
            Call DISPLAYOBJECT2.UpdateContent()
        Else
            'DISPLAYLINELENGTH = DISPLAYOBJECT2.ZoomLineLength    'preserve zoom
            DISPLAYOBJECT2.LineCount = UBound(DISPLAYLINES2)
            Call DISPLAYOBJECT2.UpdateContent()
        End If
        If Not SPEAK("ding") Then Beep()

    End Sub
    Public Function StripLeadingSpaces(sourceLine As String) As String 'get rid of leading spaces in sourceline.
        StripLeadingSpaces = sourceLine 'error default
        Do
            If Len(sourceLine) = 0 Then Exit Do
            If Strings.Left(sourceLine, 1) = " " Then
                sourceLine = Strings.Right(sourceLine, Len(sourceLine) - 1)
            Else
                Exit Do
            End If
        Loop
        StripLeadingSpaces = sourceLine
        Return StripLeadingSpaces
    End Function
    'BYTE0 = EVENTTIME.Day.ToString("00")       'day of the month. Formatted to print day with leading zero if single digit
    'BYTE1 = EVENTTIME.Month.ToString("00")     'month
    'BYTE2 = EVENTTIME.Year.ToString("0000")     'month

    Public Sub COMPAREFILES1()
        Dim localFlag As Boolean
        Dim fileNum As Integer
        Dim localWord As String

        ReDim CONTENTS1(0)
        ReDim CONTENTS2(0)
        localFlag = False
        If Not My.Computer.FileSystem.FileExists(SELECTEDFILE1) Then
            localFlag = True
        End If
        If Not My.Computer.FileSystem.FileExists(SELECTEDFILE2) Then
            localFlag = True
        End If
        If localFlag Then
            MsgBox("Files not found", MsgBoxStyle.Critical, "")
            Exit Sub
        End If
        fileNum = FreeFile()
        FileOpen(fileNum, SELECTEDFILE1, OpenMode.Input)
        Do Until EOF(fileNum)
            localWord = LineInput(fileNum)
            ReDim Preserve CONTENTS1(UBound(CONTENTS1) + 1)
            CONTENTS1(UBound(CONTENTS1)) = localWord
        Loop
        FileClose(fileNum)
        fileNum = FreeFile()
        FileOpen(fileNum, SELECTEDFILE2, OpenMode.Input)
        Do Until EOF(fileNum)
            localWord = LineInput(fileNum)
            ReDim Preserve CONTENTS2(UBound(CONTENTS2) + 1)
            CONTENTS2(UBound(CONTENTS2)) = localWord
        Loop
        FileClose(fileNum)
        Call COMPAREFILES2()

    End Sub
    Private Sub COMPAREFILES2() 'compares CONTENTS1() with CONTENTS2().
        Dim poinTer1, poinTer2 As Integer
        Dim processedLine1, localWord1 As String

        UsedReach = REACH
        SYNCHB_USED = False   'Used as a debug display on main window
        SYNCHBRECORD = ""
        ReDim DISPLAYLINES1(0)
        ReDim DISPLAYLINES2(0)
        CURRENT_LINE = "File date = " & FILE1DATE
        Call ADD_DISPLAYLINES1()
        CURRENT_LINE = "File date = " & FILE2DATE
        Call ADD_DISPLAYLINES2()
        REDLINES = ","
        BLUELINES1 = ","
        BLUELINES2 = ","
        REDLINES1 = ","
        REDLINES2 = ","
        GREENLINES1 = ","
        BROWNLINES1 = ","
        GREENLINES2 = ","
        BROWNLINES2 = ","

        ORIGINAL1 = CONTENTS1
        ORIGINAL2 = CONTENTS2
        ReDim CONTENTS1(0)
        ReDim CONTENTS2(0)
        'Now create SHADOW1 and 2 as a reduced version of CONTENTS1 and 2, depending on spaces, blanklines, and tabs
        'CONTENTS1 and 2 are used for sorting, but SHADOW 1 and 2 must be used for display.
        ReDim SHADOW1(0)
        ReDim SHADOW2(0)
        ReDim LINENUM1(0)
        ReDim LINENUM2(0)
        For poinTer1 = 1 To UBound(ORIGINAL1)
            processedLine1 = ORIGINAL1(poinTer1)
            If SHOWTABS Then
                If IGNORESPACES Then
                    processedLine1 = Replace(processedLine1, " ", "")  'remove spaces
                End If
                processedLine1 = Replace(processedLine1, Chr(9), " <tab> ")  'show tabs
                localWord1 = Replace(processedLine1, " <tab> ", "")
                If localWord1 = "" And SKIPBLANKLINES Then
                Else
                    ReDim Preserve CONTENTS1(UBound(CONTENTS1) + 1) 'CONTENTS is used for search
                    CONTENTS1(UBound(CONTENTS1)) = processedLine1
                    ReDim Preserve SHADOW1(UBound(SHADOW1) + 1)
                    SHADOW1(UBound(SHADOW1)) = processedLine1 'SHADOW is used for display
                    ReDim Preserve LINENUM1(UBound(LINENUM1) + 1)
                    LINENUM1(UBound(LINENUM1)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                End If

            ElseIf IGNORESPACES Or SKIPBLANKLINES Or IGNORETABS Then
                'if processedLine1 is empty and skipblanklines then do nothing,
                'elseif processedLine1 is empty and ignorespaces then Contents = processedLine1 and shadow = original.
                'else contents=origina', shadow = original
                If IGNORETABS Then
                    processedLine1 = Replace(processedLine1, Chr(9), " ")  'remove tabs
                End If
                localWord1 = Replace(processedLine1, " ", "")  'remove spaces
                If localWord1 = "" And SKIPBLANKLINES Then
                ElseIf processedLine1 <> "" And Not IGNORESPACES Then
                    ReDim Preserve CONTENTS1(UBound(CONTENTS1) + 1)
                    CONTENTS1(UBound(CONTENTS1)) = processedLine1
                    ReDim Preserve SHADOW1(UBound(SHADOW1) + 1)
                    SHADOW1(UBound(SHADOW1)) = processedLine1 'SHADOW is used for display
                    ReDim Preserve LINENUM1(UBound(LINENUM1) + 1)
                    LINENUM1(UBound(LINENUM1)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                Else
                    processedLine1 = Replace(processedLine1, " ", "")  'remove spaces
                    ReDim Preserve CONTENTS1(UBound(CONTENTS1) + 1)
                    CONTENTS1(UBound(CONTENTS1)) = processedLine1
                    ReDim Preserve SHADOW1(UBound(SHADOW1) + 1)
                    SHADOW1(UBound(SHADOW1)) = ORIGINAL1(poinTer1) 'for display
                    ReDim Preserve LINENUM1(UBound(LINENUM1) + 1)
                    LINENUM1(UBound(LINENUM1)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                End If
            Else
                ReDim Preserve CONTENTS1(UBound(CONTENTS1) + 1)
                CONTENTS1(UBound(CONTENTS1)) = ORIGINAL1(poinTer1)
                ReDim Preserve SHADOW1(UBound(SHADOW1) + 1)
                SHADOW1(UBound(SHADOW1)) = ORIGINAL1(poinTer1) 'for display
                ReDim Preserve LINENUM1(UBound(LINENUM1) + 1)
                LINENUM1(UBound(LINENUM1)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
            End If
        Next
        For poinTer1 = 1 To UBound(ORIGINAL2)
            processedLine1 = ORIGINAL2(poinTer1)
            If SHOWTABS Then
                If IGNORESPACES Then
                    processedLine1 = Replace(processedLine1, " ", "")  'remove spaces
                End If
                processedLine1 = Replace(processedLine1, Chr(9), " <tab> ")  'show tabs
                localWord1 = Replace(processedLine1, " <tab> ", "")
                If localWord1 = "" And SKIPBLANKLINES Then
                Else
                    ReDim Preserve CONTENTS2(UBound(CONTENTS2) + 1)
                    CONTENTS2(UBound(CONTENTS2)) = processedLine1
                    ReDim Preserve SHADOW2(UBound(SHADOW2) + 1)
                    SHADOW2(UBound(SHADOW2)) = processedLine1 'for display
                    ReDim Preserve LINENUM2(UBound(LINENUM2) + 1)
                    LINENUM2(UBound(LINENUM2)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                End If

            ElseIf IGNORESPACES Or SKIPBLANKLINES Or IGNORETABS Then
                If IGNORETABS Then
                    processedLine1 = Replace(processedLine1, Chr(9), " ")  'remove spaces
                End If
                localWord1 = Replace(processedLine1, " ", "")  'remove spaces
                If localWord1 = "" And SKIPBLANKLINES Then
                ElseIf processedLine1 <> "" And Not IGNORESPACES Then
                    ReDim Preserve CONTENTS2(UBound(CONTENTS2) + 1)
                    CONTENTS2(UBound(CONTENTS2)) = processedLine1
                    ReDim Preserve SHADOW2(UBound(SHADOW2) + 1)
                    SHADOW2(UBound(SHADOW2)) = processedLine1 'for display
                    ReDim Preserve LINENUM2(UBound(LINENUM2) + 1)
                    LINENUM2(UBound(LINENUM2)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                Else
                    processedLine1 = Replace(processedLine1, " ", "")  'remove spaces
                    ReDim Preserve CONTENTS2(UBound(CONTENTS2) + 1)
                    CONTENTS2(UBound(CONTENTS2)) = processedLine1
                    ReDim Preserve SHADOW2(UBound(SHADOW2) + 1)
                    SHADOW2(UBound(SHADOW2)) = ORIGINAL2(poinTer1) 'for display
                    ReDim Preserve LINENUM2(UBound(LINENUM2) + 1)
                    LINENUM2(UBound(LINENUM2)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
                End If
            Else
                ReDim Preserve CONTENTS2(UBound(CONTENTS2) + 1)
                CONTENTS2(UBound(CONTENTS2)) = ORIGINAL2(poinTer1)
                ReDim Preserve SHADOW2(UBound(SHADOW2) + 1)
                SHADOW2(UBound(SHADOW2)) = ORIGINAL2(poinTer1) 'for display
                ReDim Preserve LINENUM2(UBound(LINENUM2) + 1)
                LINENUM2(UBound(LINENUM2)) = Strings.Right("     " & poinTer1.ToString, 5) & " "
            End If
        Next

        poinTer1 = 1
        poinTer2 = 1
        Do
            If poinTer1 > UBound(CONTENTS1) Or poinTer2 > UBound(CONTENTS2) Then
                Call NoSynch(poinTer1 - 1, poinTer2 - 1)
                Call SHOWDISPLAY()
                Exit Sub
            End If

            If CONTENTS1(poinTer1) <> CONTENTS2(poinTer2) Then  'lines not same.
                CURRENT_LINE = LINENUM1(poinTer1 - 1) & SHADOW1(poinTer1 - 1)
                Call ADD_DISPLAYLINES1()
                CURRENT_LINE = LINENUM2(poinTer2 - 1) & SHADOW2(poinTer2 - 1)
                Call ADD_DISPLAYLINES2()
                'need to find next synch line. INDEX1, INDEX1END,  INDEX2, INDEX2END
                INDEX1 = poinTer1
                INDEX2 = poinTer2
                UsedReach = REACH
                INDEX1END = UBound(CONTENTS1)   'worst case default
                INDEX2END = UBound(CONTENTS2)

                If INDEX1 + 1 > INDEX1END Or INDEX2 + 1 > INDEX2END Then
                    Call FindNextSynchB() 'Have just displayed the common line.
                    SYNCHBRECORD = SYNCHBRECORD & "(" & INDEX1.ToString & " " & INDEX2.ToString & ") "
                Else
                    Call FindNextSynchA()
                End If

                If Not SynchWorked Then 'never seem to get here.
                    Call FindNextSynchB() 'Have just displayed the common line.
                    SYNCHBRECORD = SYNCHBRECORD & "(" & INDEX1.ToString & " " & INDEX2.ToString & ") "
                End If

                If INDEX1END = UBound(CONTENTS1) Or INDEX2END = UBound(CONTENTS2) Then
                    Call NoSynch(poinTer1, poinTer2)
                    Exit Do
                End If

                If INDEX1END - INDEX1 = INDEX2END - INDEX2 Then     'Equal number of lines, but different
                    For localcount = poinTer1 To INDEX1END - 1
                        CURRENT_LINE = localcount.ToString & " " & SHADOW1(localcount)
                        CURRENT_LINE = LINENUM1(localcount) & SHADOW1(localcount)
                        Call ADDBLUELINE1()
                    Next
                    For localcount = poinTer2 To INDEX2END - 1
                        CURRENT_LINE = localcount.ToString & " " & SHADOW2(localcount)
                        CURRENT_LINE = LINENUM2(localcount) & SHADOW2(localcount)
                        Call ADDREDLINE2()
                    Next
                Else
                    'Synch found
                    Call ShowFirstLines()
                    Call ShowSecondLines()
                End If
                'now add the next common line, from CONTENTS1
                CURRENT_LINE = (INDEX1END).ToString & " " & SHADOW1(INDEX1END)
                CURRENT_LINE = LINENUM1(INDEX1END) & SHADOW1(INDEX1END)
                Call ADD_DISPLAYLINES1()
                CURRENT_LINE = (INDEX2END).ToString & " " & SHADOW2(INDEX2END)
                CURRENT_LINE = LINENUM2(INDEX2END) & SHADOW2(INDEX2END)
                Call ADD_DISPLAYLINES2()
                CURRENT_LINE = "============================================= "
                Call ADD_DISPLAYLINES1()
                Call ADD_DISPLAYLINES2()
                poinTer1 = INDEX1END
                poinTer2 = INDEX2END
            End If
            poinTer1 = poinTer1 + 1
            poinTer2 = poinTer2 + 1
        Loop
        Call SHOWDISPLAY()
    End Sub

    Private Sub FindNextSynchA()
        'INDEX1 and INDEX2 point to the first non-matching lines in CONTENTS1, CONTENTS2.
        'there could be extra lines in CONTENTS1, or in CONTENTS2.
        'Grab two lines from CONTENTS1, and search through entire CONTENTS2 foer a match.
        'If found, record the match indexes, and exit the search. If not, record ubound(contents1 and 2)
        'then increment the start index from CONTENTS1's two lines, and repeat the procedure.
        'Search CONTENTS 2 up to the match previously found. If a match is found earlier,
        'exit the search and replace the match indices.
        'Repeat until the gap in CONTENTS1 search is bigger than the gap in CONTENTS2 search.
        Dim match1Start, match2Start, localNum1, localNum2 As Integer
        Dim firstWord, secondWord As String

        Dim startOffset1, startOffset2 As Integer

        Dim localWord1, localWord2 As String
        Dim localFlag As Boolean

        match1Start = UBound(CONTENTS1)   'worst case default
        match2Start = UBound(CONTENTS2)

        NUM1 = 1000000 'maximum number of lines ALLOWED?
        startOffset1 = 0
        startOffset2 = 0

        Do
            firstWord = CONTENTS1(INDEX1 + startOffset1)
            secondWord = CONTENTS1(INDEX1 + startOffset1 + 1)
            localFlag = False
            For loop2 = INDEX2 To UBound(CONTENTS2) - 1
                localWord1 = CONTENTS2(loop2)
                localWord2 = CONTENTS2(loop2 + 1)
                If firstWord = localWord1 Then
                    If secondWord = localWord2 Then  'have found a match. 
                        'save index1+startoffset in match1Start, if smaller than match1Start.
                        'save loop2 in match2start, if smaller than match2 start
                        If INDEX1 + startOffset1 < match1Start Or loop2 < match2Start Then
                            NUM2 = startOffset1 + loop2
                            If NUM2 < NUM1 Then 'This match looks closer than previously found matches.
                                NUM1 = NUM2
                                localNum1 = INDEX1 + startOffset1    'debug
                                localNum2 = loop2        'debug
                                match1Start = INDEX1 + startOffset1
                                match2Start = loop2
                                SynchWorked = True
                                localFlag = True
                                Exit For
                            End If
                        End If
                        Exit For  'Match found, but later
                    End If
                End If
            Next
            startOffset1 = startOffset1 + 1
            If (INDEX1 + startOffset1) >= UBound(CONTENTS1) - 1 Then Exit Do
        Loop

        If match1Start <> UBound(CONTENTS1) Then
            INDEX1END = match1Start 'first line of equaity
            INDEX2END = match2Start
        End If
    End Sub
    Private Sub FindNextSynchB() 'Given INDEX1 and INDEX2 as start points into CONTENTS1 and CONTENTS2.
        'Return INDEX1END and INDEX2END as the first line where synch is found.
        Dim localCount1, localCount2, foundSynch1, foundSynch2 As Integer
        Dim startWord, endWord As String
        Dim inx1End1, inx2End1, inx1End2, inx2End2, steP1, steP2, diFF1, diFF2 As Integer
        Dim inx1Found, inx2Found As Boolean

        SYNCHB_USED = True
        'Assume CONTENTS2 has extra lines inserted.
        foundSynch1 = INDEX1
        foundSynch2 = INDEX2
        inx1End1 = UBound(CONTENTS1) 'default
        inx2End1 = UBound(CONTENTS2) 'default
        inx1Found = False
        inx2Found = False
        Do
            startWord = CONTENTS1(foundSynch1)   'The line in contents1 to search a match for.
            For localCount2 = INDEX2 To UBound(CONTENTS2)
                endWord = CONTENTS2(localCount2)
                If endWord = startWord Then 'synch found
                    inx1End1 = foundSynch1
                    inx2End1 = localCount2
                    inx1Found = True
                    a1 = 1
                    Exit Do
                End If
            Next
            foundSynch1 = foundSynch1 + 1
            If foundSynch1 > UBound(CONTENTS1) Then Exit Do
        Loop
        Do
            'now look for extra lines in CONTENTS1
            startWord = CONTENTS2(foundSynch2)   'The first line in contents2 which doesn't match
            For localCount1 = INDEX1 To UBound(CONTENTS1)
                endWord = CONTENTS1(localCount1)
                If endWord = startWord Then 'synch found
                    inx1End2 = localCount1
                    inx2End2 = foundSynch2
                    inx2Found = True
                    Exit Do
                End If
            Next
            foundSynch2 = foundSynch2 + 1
            If foundSynch2 > UBound(CONTENTS2) Then Exit Do
        Loop

        'Combinations of inxfound: 0,1=use end2, 1,0 use end1, 0,0 use ubound(), 1,1 choose below.
        If inx1Found = False And inx2Found = False Then
            INDEX1END = UBound(CONTENTS1)
            INDEX2END = UBound(CONTENTS2)
            GoTo LastLine
        End If
        If inx1Found = False And inx2Found = True Then 'use End2
            INDEX1END = inx2End1
            INDEX2END = inx2End2
            GoTo LastLine
        End If
        If inx1Found = True And inx2Found = False Then 'use End1
            INDEX1END = inx1End1
            INDEX2END = inx1End2
            GoTo LastLine
        End If

        'steP1 = lesser of [inx1end1 - INDEX1 compared with inx1end2 - INDEX1]. Use the lesser step
        'steP2 lesser of [inx2end1 - INDEX2 compared with inx2end2 - INDEX2]. use the lesser step.
        'choose the lesser of steP1 and steP2.
        'If steP1 < steP2 then choose wich matches - [inx1end1 - INDEX1 or inx1end2 - INDEX1]
        'If steP1 > steP2 then choose wich matches - [inx2end1 - INDEX2 or inx2end2 - INDEX2]
        diFF1 = inx1End1 - INDEX1
        diFF2 = inx1End2 - INDEX1
        steP1 = diFF2
        If diFF1 < diFF2 Then steP1 = diFF1

        diFF1 = inx2End1 - INDEX2
        diFF2 = inx2End2 - INDEX2
        steP2 = diFF2
        If diFF1 < diFF2 Then steP2 = diFF1

        If steP1 < steP2 And steP1 <> 0 Then  'use inx1
            If inx1End1 - INDEX1 = steP1 Then
                INDEX1END = inx1End1
                INDEX2END = inx2End1
            Else 'use inx2
                INDEX1END = inx1End2
                INDEX2END = inx2End2
            End If
        ElseIf steP2 <> 0 Then
            If inx2End1 - INDEX2 = steP2 Then
                INDEX1END = inx1End1
                INDEX2END = inx2End1
            Else 'use inx2
                INDEX1END = inx1End2
                INDEX2END = inx2End2
            End If
        Else
            If inx1End1 - INDEX1 = steP1 Then
                INDEX1END = inx1End1
                INDEX2END = inx2End1
            Else 'use inx2
                INDEX1END = inx1End2
                INDEX2END = inx2End2
            End If
        End If
LastLine:
    End Sub
    Private Sub NoSynch(Ptr1 As Integer, Ptr2 As Integer)
        'Show lines from Pointer1 to UBound(CONTENTS1) in blue, in DISPLAYLINES1
        'Then show lines from poinTer2 to UBound(CONTENTS2) in red, in DISPLAYLINES2
        Dim localCount1, localCount2 As Integer

        If INDEX1END <> UBound(CONTENTS1) And INDEX2END <> UBound(CONTENTS2) Then
            CURRENT_LINE = (Ptr1 - 1).ToString & " " & CONTENTS1(Ptr1 - 1)
            CURRENT_LINE = LINENUM1(Ptr1 - 1) & CONTENTS1(Ptr1 - 1)
            Call ADD_DISPLAYLINES1()
            CURRENT_LINE = (Ptr1).ToString & " " & CONTENTS1(Ptr1)
            CURRENT_LINE = LINENUM1(Ptr1) & CONTENTS1(Ptr1)
            Call ADDBLUELINE1()

            CURRENT_LINE = (Ptr2 - 1).ToString & " " & CONTENTS2(Ptr2 - 1)
            CURRENT_LINE = LINENUM2(Ptr2 - 1) & CONTENTS2(Ptr2 - 1)
            Call ADD_DISPLAYLINES2()
            CURRENT_LINE = (Ptr2).ToString & " " & CONTENTS2(Ptr2)
            CURRENT_LINE = LINENUM2(Ptr2) & CONTENTS2(Ptr2)
            Call ADDREDLINE2()

        End If
        CURRENT_LINE = "NO FURTHER SYNCH FOUND."
        Call ADD_DISPLAYLINES1()
        Call ADD_DISPLAYLINES2()
        localCount1 = UBound(CONTENTS1) - Ptr1 + 1
        CURRENT_LINE = "First file has " & localCount1.ToString & " more lines."
        Call ADDGREENLINE1()
        localCount2 = UBound(CONTENTS2) - Ptr2 + 1
        CURRENT_LINE = "Second file has " & localCount2.ToString & " more lines."
        Call ADDGREENLINE2()
    End Sub
    Private Sub ShowFirstLines()
        'first file has extra lines
        'Show the missed lines from file2 first, in brown. 
        Dim localCount1 As Integer
        For localCount1 = INDEX1 To INDEX1END - 1
            CURRENT_LINE = LINENUM1(localCount1) & SHADOW1(localCount1)
            Call ADDBLUELINE1()    'ADDBROWNLINE2 ??
        Next
        If INDEX1END = UBound(CONTENTS1) Then
            CURRENT_LINE = LINENUM1(INDEX1END) & SHADOW1(INDEX1END)
            Call ADDBLUELINE1()    'ADDBROWNLINE2 ??
        End If
    End Sub
    Private Sub ShowSecondLines()
        'Synch found, but second file has extra lines.
        'Show the missed lines from file1 first, in brown. 
        Dim localCount2 As Integer

        For localCount2 = INDEX2 To INDEX2END - 1
            CURRENT_LINE = (localCount2).ToString & " " & SHADOW2(localCount2)
            CURRENT_LINE = LINENUM2(localCount2) & SHADOW2(localCount2)
            Call ADDREDLINE2()
        Next
        If INDEX2END = UBound(CONTENTS2) Then
            CURRENT_LINE = (INDEX2END).ToString & " " & SHADOW2(INDEX2END)
            CURRENT_LINE = LINENUM2(INDEX1END) & SHADOW2(INDEX1END)
            Call ADDREDLINE2()    'ADDBROWNLINE2 ??
        End If
    End Sub
End Module
