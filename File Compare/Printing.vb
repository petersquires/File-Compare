Imports System
Imports System.IO
Imports System.IO.Ports
Imports Microsoft.VisualBasic
Imports System.Timers
'Imports Microsoft.VisualBasic.Devices
Imports System.Threading
Imports System.Windows.Forms
Imports System.Media
Imports System.Drawing.Point
'Imports System.Drawing.Printing
Imports System.Printing
Imports System.Math



Module Printing
    Public StartPage, EndPage, MaxNamesLen As Integer
    Public MyPaperSize, MyOrientation, LinesPerPage, MyDefaultFont, LocalFontSize, MyMargin As String
    'PaperSize = ["Letter" "A5" "A4" "A3"]
    ' Orientation = ["Portrait" "Landscape"] 
    'Linesperpage  = "40"
    'MyDefaultFont = [ "Lucida Console" "Arial" Times"]
    Public ChosenDefault As String
    Public FILE2PRINT As String  'the name of the saved file to print.
    Public PRINTBUFFER() As String
    Public PageBuffer() As String
    Public streamToPrint As StreamReader
    Public PageNumber As Integer
    Public PrinterFailed As Boolean
    Public SelectedCourse As String

    '***************************************************************************************************************
    Public PrinterCalled As Boolean
    Private NEW_THREAD As Thread
    Private NEW_THREAD_LOADED As Boolean

    Public LineMaxLength As Integer
    Public LineCount, PageCount, CharCount, TotalPages, ExtraLines As Integer
    'Public HeaderLine1, HeaderLine2, CatPlacing, tempFileName As String
    Public HeaderLine1, HeaderLine2, CatPlacing As String
    Public CategoryFile(), TagnumFile(), PrintTagFile(), TagNamesFile() As String
    Public PPaperSize, POrientation, PLinesPerPage, PDefaultFont, PFontSize As String      'PORTRAIT setup defaults for RESULTS
    Public LPaperSize, LOrientation, LLinesPerPage, LDefaultFont, LFontSize As String      'landscape setup
    Public WPF_PLinesPerPage, WPF_RLinesPerPage, WPF_CLinesPerPage, WPF_WLinesPerPage, WPF_RPLinesPerPage, WPF_WPLinesPerPage, WPF_CPLinesPerPage As String  'Overrides the non-WPF ones.
    Public RPaperSize, RLinesPerPage, RDefaultFont, RFontSize As String          'RESULTS LANDSCAPE setup
    Public RPPaperSize, RPLinesPerPage, RPDefaultFont, RPFontSize As String      'RESULTS PORTRAIT setup
    Public WPaperSize, WLinesPerPage, WDefaultFont, WFontSize As String          'WINNERS LANDSCAPE setup
    Public WPPaperSize, WPLinesPerPage, WPDefaultFont, WPFontSize As String      'WINNERS PORTRAIT setup
    Public CPaperSize, CLinesPerPage, CDefaultFont, CFontSize As String          'CATEGORIES LANDSCAPE setup
    Public CPPaperSize, CPLinesPerPage, CPDefaultFont, CPFontSize As String          'CATEGORIES PORTRAIT setup
    Public ROrientation, WOrientation, COrientation As String
    Public OPaperSize, OOrientation, OLinesPerPage, ODefaultFont, OLocalFontSize As String
    Public place, TeamNum, Score, FinishOrEqual, CategoryNames(), CatString, TagCategories As String
    Public CategoryRankings(), EqualCount, LastRanking(), CategoryEqualCount() As Long
    Public Equal, FoundCategory, DoingFormLoad, ResultsDebug, FirstLoad, NamesFound As Boolean
    Public CriterionFile, PunchList(), tagList(), CriterionResults(), NamesList, TempWrite(), tempLine As String
    Public PunchAverage() As Long
    Public AverageTime, StartTime As Date
    Public CourseName, TeamLine1, TeamLine2, TeamLine3, PrintfileName, WinnersPrint() As String
    Public CourseData(), NamesFillerLine, EqualCatsList, DummyFile(), EqualList(), TabsWithEqual() As String
    Public UnitedCourseNames(), TagAndNames(), UnitedCourseName, UnitedSummary() As String
    Public lowER, highER, TotalFlag, CombineNames As Boolean
    Public LowResultIndex, LastLowResultIndex As Integer
    Public PrinterMessage, SaveFileName, LocalPath As String
    Public DesiredOrientation, ResultsType, TotalNames() As String
    'Public LinesPerPage, LocalFontSize As String
    Public mnuFinishTime_Checked, mnuGROSS_Checked, mnuNAMES_Checked, mnuHideUnscored_Checked, mnuCatResults_Checked As Boolean
    Public Check2_Checked, mnuAllResults_Checked, mnuWinners_Checked, Check4_Checked, ROLLINGDISPLAY As Boolean

    Public CoursesToSearch(), RESCORETAGS, PrintCourseTitle, ScoredFileDate(), StartTimeOfDirname As String
    Public ScrollOpen, ScrollRescoring, BackGroundResults As Boolean
    Public SCROLLHEADER, SCROLLHEADER1 As String
    Private HeaderFileName As String
    Private PrintStart, PrintEnd As Long
    Public PrintThreadEnabled As Boolean
    Public ThreadWarningKilled, ShowPrintSpeed As Boolean
    Public ThreadBuffer As String

    Private Sub Count3Tabs()        'Strips CURRENT_LINE up tp 3 tabs. All to left of 3 tabs, incl. Tab, in BYTE1. Remainder in CURRENT_LINE
        Dim Removed As Integer
        BYTE0 = ""      'default.
        BYTE1 = ""      'default
        If Len(CURRENT_LINE) = 0 Then Exit Sub
        Removed = 0
        Do      'Find the start of text
            BYTE1 = BYTE1 & Strings.Left(CURRENT_LINE, 1)
            If Strings.Left(CURRENT_LINE, 1) = Chr(9) Then
                CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - 1)   'remove  tab.
                Removed = Removed + 1
                If Removed = 3 Then
                    Exit Sub
                End If
            Else
                CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - 1)   'remove data
            End If
            If Len(CURRENT_LINE) = 0 Then Exit Sub
        Loop
    End Sub

    Public Sub SendToPrinter()
        'LinesPerPage is string. LineCount, PageCount As Integer. HeaderLine As String
        Dim localNum, localPos, localCount As Integer
        Dim localDouble As Double
        Dim localMessage As String
        Dim localLinesPerPage As String
        Dim localFlag As Boolean = False
        localMessage = ""


        ReDim DISPLAYLINES1(0)
        CURRENT_LINE = HeaderLine2
        CharCount = 0
        Call AddToDisplay()   'adds to displaylines()
        For localCount = 1 To UBound(WRITEFILE)
            CharCount = -1      'space the line across 1 to the right.
            CURRENT_LINE = " " & WRITEFILE(localCount)
            Call AddToDisplay()
        Next localCount
        Call SHOWDISPLAY()


        If Check2_Checked = True Or (ScrollRescoring And (PrintfileName = "CATRESULTS" Or PrintfileName = "ALL")) Then   'PRINT
            If mnuAllResults_Checked = True And PrintfileName <> "ALL" Then Exit Sub
            If mnuWinners_Checked = True And PrintfileName <> "SECTION" Then Exit Sub
            If mnuCatResults_Checked = True And PrintfileName <> "CATRESULTS" Then Exit Sub

            STORERESULTS = True 'send to default printer

            If Not ScrollRescoring Then
                SENDPRINT("")                     'space line
                LineCount = LineCount + 1
                SENDPRINT(Chr(9) & HeaderLine1)
                LineCount = LineCount + 1
                SENDPRINT("")                     'space line
                LineCount = LineCount + 1
                SENDPRINT(Chr(9) & HeaderLine2)
                LineCount = LineCount + 1
            End If

            For localCount = 1 To UBound(WRITEFILE)      'Format the page breaks here
                CURRENT_LINE = Chr(9) & WRITEFILE(localCount)
                If InStr(1, WRITEFILE(localCount), Chr(13), CompareMethod.Text) <> 0 Then LineCount = LineCount + 1
                SENDPRINT(CURRENT_LINE)
                LineCount = LineCount + 1
                If LinesPerPage Is Nothing Then LinesPerPage = "80" 'An error if Rolling Results, and Menu used.
                localLinesPerPage = LinesPerPage
                If LinesPerPage = "Auto" Then localLinesPerPage = UBound(WRITEFILE).ToString
                localNum = CInt(LineCount Mod Val(localLinesPerPage))  'number of line printed so far
                localPos = 1    'normal max. number of lines left spare at bottom
                If PrintfileName = "CATRESULTS" Then
                    localDouble = Val(localLinesPerPage) - localNum   'lines left this page
                    If CURRENT_LINE = Chr(9) Then
                        If localDouble < 4 Then
                            localPos = localNum + 1     'force a page break.
                            LineCount = CInt(LineCount + localDouble) - 1
                        End If
                        If Strings.Left(WRITEFILE(localCount - 1), 6) = "------" And localDouble < 10 Then
                            localPos = localNum + 1     'force a page break.
                            LineCount = CInt(LineCount + localDouble) - 1
                        End If
                    End If
                End If

                If localNum <= localPos And Not ScrollRescoring And LinesPerPage <> "Auto" Then 'page overflow.
                    NEXTPAGE()
                    PageCount = PageCount + 1
                    If LinesPerPage = "Auto" Then
                        HeaderLine1 = Chr(9) & HeaderFileName & " Results." & Chr(9) & Chr(9) & "PAGE " & Str$(PageCount + 1)
                    Else
                        HeaderLine1 = Chr(9) & HeaderFileName & " Results." & Chr(9) & Chr(9) & "PAGE " & Str$(PageCount + 1) & " of " & Str$(TotalPages)
                    End If
                    SENDPRINT("")                    'space line
                    LineCount = LineCount + 1
                    SENDPRINT(Chr(9) & HeaderLine1)
                    LineCount = LineCount + 1
                    SENDPRINT("")                    'space line
                    LineCount = LineCount + 1
                    SENDPRINT(Chr(9) & HeaderLine2)
                    LineCount = LineCount + 1
                End If
            Next localCount
            ENDPRINT()      'prints it. Needs MyPaperSize, 
        End If
        If localFlag = True Then Check2_Checked = False

    End Sub

    Private Sub AddToDisplay()      'adds current_line to display stack, after adding a filler and line end.
        Dim fiLLer, tabFreeLine, localChar As String
        Dim TabSize As Integer

        fiLLer = "   "
        tabFreeLine = ""
        TabSize = 8
        Do While Len(CURRENT_LINE) > 0
            localChar = Strings.Left(CURRENT_LINE, 1)
            If localChar = Chr(9) Then
                localChar = Strings.Left("        ", (TabSize - ((CharCount) Mod TabSize)))
                CharCount = 0
            Else
                CharCount = CharCount + 1
            End If
            tabFreeLine = tabFreeLine & localChar
            CURRENT_LINE = Strings.Right(CURRENT_LINE, Len(CURRENT_LINE) - 1)
        Loop

        CURRENT_LINE = tabFreeLine

        ReDim Preserve DISPLAYLINES1(UBound(DISPLAYLINES1) + 1)
        DISPLAYLINES1(UBound(DISPLAYLINES1)) = CURRENT_LINE & fiLLer & vbCrLf
    End Sub

    Public Function SENDPRINT(lineOfPrint As String) As Integer
        Try
            If UBound(PRINTBUFFER) <= 1 Then
                PageNumber = 1
            End If
        Catch ex As Exception
            ReDim PRINTBUFFER(0)
            PageNumber = 1
        End Try
        'ReDim DEBUGPRINTLINES(0)
        Try

            ReDim Preserve PRINTBUFFER(UBound(PRINTBUFFER) + 1)
                PRINTBUFFER(UBound(PRINTBUFFER)) = lineOfPrint & vbCr & vbLf

                ReDim Preserve DEBUGPRINTLINES(UBound(DEBUGPRINTLINES) + 1)
            DEBUGPRINTLINES(UBound(DEBUGPRINTLINES)) = lineOfPrint & vbCr & vbLf
            If UBound(DEBUGPRINTLINES) > 2000 Then ReDim DEBUGPRINTLINES(0) 'safety catch
        Catch ex As Exception
        End Try
        Return 0
    End Function
    Public Function ENDPRINT() As Integer
        Dim localNum As Integer = 1
        localNum = ENDPRINT1()
        Return localNum
    End Function

    Private Function ENDPRINT1() As Integer
        Dim localFlag As Boolean
        Static timeNumber As Integer = 0
        Dim localLoop, fileNum As Integer
        Static simPageCount As Integer = 0
        're-instate default printer.

        localFlag = False
        MyPaperSize = "A4"
        MyOrientation = "Portrait"
        LinesPerPage = "Auto"
        MyDefaultFont = "Lucida Console"
        LocalFontSize = "12"
        MyMargin = "18"      '70


        FILE2PRINT = PARENTPATH & "\PrintingBuffer.txt"
        ThreadBuffer = PARENTPATH & vbCrLf
        ThreadBuffer = ThreadBuffer & MyPaperSize & vbCrLf
        ThreadBuffer = ThreadBuffer & MyOrientation & vbCrLf
        ThreadBuffer = ThreadBuffer & LinesPerPage & vbCrLf
        ThreadBuffer = ThreadBuffer & MyDefaultFont & vbCrLf
        ThreadBuffer = ThreadBuffer & LocalFontSize & vbCrLf
        ThreadBuffer = ThreadBuffer & ChosenDefault & vbCrLf
        ThreadBuffer = ThreadBuffer & "Margin=" & MyMargin & vbCrLf
        For localLoop = 1 To UBound(PRINTBUFFER)
            ThreadBuffer = ThreadBuffer & PRINTBUFFER(localLoop)    'PRINTBUFFER already includes vbCrLf
        Next
        'PrintStart = Date.Now.Ticks

        Try
            If Not PRINTDEBUG And STORERESULTS Then  'Save and PRINT the document
                fileNum = FreeFile()
                FileOpen(fileNum, PARENTPATH & "\PrintingBuffer.txt", OpenMode.Output)
                For localIndex As Integer = 1 To UBound(PRINTBUFFER)
                    BYTE9 = PRINTBUFFER(localIndex)  'debug  for inspection
                    Print(fileNum, PRINTBUFFER(localIndex))
                Next
                FileClose(fileNum)
                Call ThreadPrint(ThreadBuffer)
            End If
        Catch ex As Exception

        End Try

        ReDim PRINTBUFFER(0)
        Return 0
    End Function

    Public Function NEXTPAGE() As Integer
        If Not PRINTDEBUG And STORERESULTS Then
            'Printer.NewPage()
            'ReDim Preserve PRINTBUFFER(UBound(PRINTBUFFER) + 1)
            'PRINTBUFFER(UBound(PRINTBUFFER)) = "PAGE " & CStr(PageNumber) & vbCr & vbLf
            ReDim Preserve PRINTBUFFER(UBound(PRINTBUFFER) + 1)
            PRINTBUFFER(UBound(PRINTBUFFER)) = "<NEWPAGE>" & vbCr & vbLf
        End If

        'ReDim Preserve DEBUGPRINTLINES(UBound(DEBUGPRINTLINES) + 1)
        'DEBUGPRINTLINES(UBound(DEBUGPRINTLINES)) = "PAGE " & CStr(PageNumber) & vbCr & vbLf
        If STORERESULTS Then
            ReDim Preserve DEBUGPRINTLINES(UBound(DEBUGPRINTLINES) + 1)
            DEBUGPRINTLINES(UBound(DEBUGPRINTLINES)) = "- - - - - - NEXT  PAGE - - - - - " & vbCr & vbLf
            PageNumber = PageNumber + 1
        End If
        Return 0
    End Function


End Module

