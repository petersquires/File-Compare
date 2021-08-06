
Imports System.IO
Imports System.Printing
Module PrintThread

    Private PrintStart, PrintEnd As Long
    Private ThreadPageBuffer(), PrintCaption As String
    'Private PrintPageCount As Integer = 1
    Private ThreadPrintBuffer(), ThreadFilePrintName, THREADPARENT As String
    Private PrintFailed As Boolean
    Private ThisPaperSize, ThisOrientation, ThisLinesPerPage, ThisDefaultFont, ThisLocalFontSize, ThisPrinter, ThisMargin As String
    Private SpacedTabLine, TabbedLine As String
    Private PrinterFont As String

    Public Sub ThreadPrint(ByVal data As Object) '        Run until the application tries to shutdown
        Dim localString, localChar, firstLine As String
        Dim localPos, loopCount As Integer
        Dim reportNum As Integer
        Dim elaPsed As Double


        ReDim ThreadPrintBuffer(0)
        localString = data.ToString
        loopCount = 0
        firstLine = "SQRX7YT"
        ThisMargin = ""

        Do
            If Len(localString) < 1 Then Exit Do
            localPos = InStr(localString, vbCrLf)
            If localPos = 0 Then Exit Do
            localChar = Strings.Left(localString, localPos - 1)
            localString = Strings.Right(localString, Len(localString) - localPos - 1)
            Select Case loopCount
                Case 0
                    THREADPARENT = localChar
                Case 1
                    ThisPaperSize = localChar
                Case 2
                    ThisOrientation = localChar
                Case 3
                    ThisLinesPerPage = localChar
                Case 4
                    ThisDefaultFont = localChar
                Case 5
                    ThisLocalFontSize = localChar
                Case 6
                    ThisPrinter = localChar
                Case 7
                    If Strings.Left(localChar, 7) = "Margin=" Then
                        ThisMargin = Strings.Right(localChar, Len(localChar) - 7)
                    Else
                        firstLine = localChar
                        ReDim Preserve ThreadPrintBuffer(UBound(ThreadPrintBuffer) + 1)
                        ThreadPrintBuffer(UBound(ThreadPrintBuffer)) = localChar
                    End If
                Case 9
                    If firstLine = "SQRX7YT" Then firstLine = localChar
                    ReDim Preserve ThreadPrintBuffer(UBound(ThreadPrintBuffer) + 1)
                    ThreadPrintBuffer(UBound(ThreadPrintBuffer)) = localChar
                Case Else
                    ReDim Preserve ThreadPrintBuffer(UBound(ThreadPrintBuffer) + 1)
                    ThreadPrintBuffer(UBound(ThreadPrintBuffer)) = localChar
            End Select
            loopCount = loopCount + 1
        Loop
        PrintStart = Date.Now.Ticks
        Call PrintThisList() 'prints PrintingBuffer.txt according to the settings. Doesn't actually read PrintingBuffer; uses ThreadPrintBuffer
        PrintEnd = Date.Now.Ticks

        If ShowPrintSpeed Then
            elaPsed = (PrintEnd - PrintStart) / 10000000
            reportNum = FreeFile()
            FileOpen(reportNum, THREADPARENT & "\PrintSpeed" & ".txt", OpenMode.Output)
            Print(reportNum, "Printer is " & ThisPrinter & vbCrLf)
            Print(reportNum, "First line is  " & firstLine & vbCrLf)
            Print(reportNum, "Elapsed time in seconds =  " & CStr(elaPsed) & vbCrLf)
            FileClose(reportNum)
        End If
    End Sub
    Private Sub PrintThisList()    ' 'prints ThreadPrintBuffer. Calls PrintOnePage whenever the line count exceeds ThisLinesPerPage, or whenever <NEWPAGE> is found.
        Dim localCount, sourceIndex As Integer
        Dim localLine As String

        PrintFailed = False

        Dim pageTotalLines As Integer = CInt(Val(ThisLinesPerPage))
        If pageTotalLines = 0 Then pageTotalLines = UBound(ThreadPrintBuffer)
        sourceIndex = 1

        Do
            localCount = 1
            ReDim ThreadPageBuffer(0)
            localLine = ""
            If sourceIndex >= UBound(ThreadPrintBuffer) Then Exit Do
            Do
                If sourceIndex > UBound(ThreadPrintBuffer) Then Exit Do
                TabbedLine = ThreadPrintBuffer(sourceIndex)
                Call SpaceTheTabs() 'Input is LineWithTabs, and returns TabSpacedLine
                If SpacedTabLine = "<NEWPAGE>" Or localCount = pageTotalLines Or sourceIndex = UBound(ThreadPrintBuffer) Then 'print this page
                    If SpacedTabLine <> "<NEWPAGE>" Then
                        ReDim Preserve ThreadPageBuffer(UBound(ThreadPageBuffer) + 1)
                        ThreadPageBuffer(UBound(ThreadPageBuffer)) = SpacedTabLine
                        localLine = localLine & SpacedTabLine & vbCrLf
                    Else        '="<NEWPAGE>"
                        If localCount = 1 Then  'NEWPAGE inserted externally. Code added March 2018
                            sourceIndex = sourceIndex + 1
                            Exit Do
                        End If
                    End If

                    Call PrintOnePage(localLine)    'Prints ThreadPageBuffer as a single page
                    If PrintFailed Then
                        PrintFailed = False
                        Exit Sub
                    End If
                    sourceIndex = sourceIndex + 1
                    Exit Do
                Else
                    ReDim Preserve ThreadPageBuffer(UBound(ThreadPageBuffer) + 1)
                    ThreadPageBuffer(UBound(ThreadPageBuffer)) = SpacedTabLine
                    localLine = localLine & SpacedTabLine & vbCrLf
                    localCount = localCount + 1
                End If
                sourceIndex = sourceIndex + 1
            Loop
        Loop
    End Sub
    Private Sub PrintOnePage(thisText As String)  'prints ThreadPageBuffer (an array of lines) with ThisPaperSize, ThisOrientation, ThisLinesPerPage, ThisDefaultFont, ThisLocalFontSize, ThisPrinter, ThisMargin
        Dim localLoop, paperSizeIndex, localPos As Integer
        Dim myDefaultPrinter, findPaperSize, localWord, localChar As String
        Dim thisPageWidth, thisPageHeight, holdDouble As Double
        'Static PrintQueue As Integer = 0
        Static PrintPageCount As Integer = 1
        Try
            'Find installed printers. Don't need this here but it is a sample of how to do it.
            For Each thing As String In System.Drawing.Printing.PrinterSettings.InstalledPrinters
                localWord = thing   '"Brother DCP-J4120DW Printer", "Send To OneNote 2016", "Foxit PhantomPDF Printer"
            Next

            'Open Printdialog and Find default printer.  Don't need default here but this shows how.
            Dim pDiag As New PrintDialog()
            myDefaultPrinter = pDiag.PrintQueue.Name            '"Brother DCP-J4120DW Printer", "Send To OneNote 2016", "Foxit PhantomPDF Printer"

            'Set new printer
            pDiag.PrintQueue = New PrintQueue(New PrintServer(), ThisPrinter)
            pDiag.PageRangeSelection = PageRangeSelection.AllPages
            Dim myQueue As PrintQueue = pDiag.PrintQueue
            'BYTE2 = myQueue.Name            '"Brother DCP-J4120DW Printer", "Send To OneNote 2016", "Foxit PhantomPDF Printer"

            Dim myCap As PrintCapabilities = myQueue.GetPrintCapabilities
            ''Find orientation options. We do not use this
            'Dim myCol = myCap.PageOrientationCapability 'two orientations so = 2
            'Dim myArray As Array = myCol.ToArray
            'For localLoop = 0 To UBound(myArray)
            'localWord = myArray(localLoop).ToString 'Portrait, Landscape  for foxit pdf
            'Next

            findPaperSize = "ISO" & ThisPaperSize & " "
            paperSizeIndex = -1

            thisPageWidth = 793     'A4 DEFAULT
            thisPageHeight = 1122   'A4 default

            'FInd the index for the font type we have selected, and find the page height and width which matches that size.
            Dim myPaper = myCap.PageMediaSizeCapability
            Dim mySizes As Array = myPaper.ToArray
            For localLoop = 0 To UBound(mySizes)
                localWord = mySizes(localLoop).ToString      '101 sizes
                If InStr(localWord, findPaperSize) <> 0 Then
                    paperSizeIndex = localLoop
                    '"ISOA3 (1122.51968503937 x 1587.40157480315)"
                    '"ISOA4 (793.700787401575 x 1122.51968503937)"
                    '"ISOA5 (559.370078740158 x 793.700787401575)"
                    '"ISOA6 (396.850393700787 x 559.370078740158)"
                    localPos = InStr(localWord, "(")
                    If localPos <> 0 Then
                        localWord = Strings.Right(localWord, (Len(localWord) - localPos))
                        localPos = InStr(localWord, ".")
                        If localPos <> 0 Then
                            localChar = Strings.Left(localWord, localPos - 1)
                            thisPageWidth = CDbl(localChar)
                        End If
                        localPos = InStr(localWord, "x", CompareMethod.Text)
                        If localPos <> 0 Then
                            localWord = Strings.Right(localWord, (Len(localWord) - localPos - 1))
                            localPos = InStr(localWord, ".")
                            If localPos <> 0 Then
                                localChar = Strings.Left(localWord, localPos - 1)
                                thisPageHeight = CDbl(localChar)
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next

            Dim myTicket As PrintTicket = myQueue.UserPrintTicket
            If UCase(ThisOrientation) = "LANDSCAPE" Then
                myTicket.PageOrientation = PageOrientation.Landscape
                holdDouble = thisPageWidth
                thisPageWidth = thisPageHeight
                thisPageHeight = holdDouble
            Else
                myTicket.PageOrientation = PageOrientation.Portrait
            End If

            If paperSizeIndex <> -1 Then
                myTicket.PageMediaSize = myCap.PageMediaSizeCapability(paperSizeIndex)      'from above
            End If

            'Create the document, passing a new paragraph and new run using text
            Dim doc As New FlowDocument(New Paragraph(New Run(thisText)))
            Dim myPaginatorSource As IDocumentPaginatorSource
            myPaginatorSource = CType(doc, IDocumentPaginatorSource)

            ' Dim mySize As Size = myPaginatorSource.DocumentPaginator.PageSize   '816,1056  for Brother; 816, 1056 for Foxit  Size before margins. A4 = 793,1122
            Dim marGin As Double = 15       'default mm
            If ThisMargin <> "" Then marGin = CDbl(ThisMargin)
            marGin = marGin * 3.78               'Units of mm. Need units of emSize = 1/96 Inch. emSize = marGin in mm *96/25.4  = marGin in mm * 3.78
            doc.PagePadding = New Thickness(marGin) 'Creates left and top margin around the page. May have to find printable area.
            doc.PageWidth = thisPageWidth - marGin    'setting this causes wrap at the width setting. PageWidth = 
            doc.PageHeight = thisPageHeight - marGin       'Dimensions are in emSize units of 1/96th Inch.
            doc.FontFamily = New FontFamily(ThisDefaultFont)
            doc.FontSize = CDbl(ThisLocalFontSize)
            'doc.FontStyle = FontStyles.Italic
            doc.FontStyle = FontStyles.Normal
            doc.ColumnWidth = thisPageWidth
            'Dim aBool As Boolean = doc.IsColumnWidthFlexible   'Default is True
            doc.TextAlignment = TextAlignment.Left
            Dim aBool As Boolean = doc.IsHyphenationEnabled     'false
            Dim bBool As Boolean = doc.IsOptimalParagraphEnabled    'false
            'doc.LineStackingStrategy = LineStackingStrategy.MaxHeight '1       'Default
            'doc.LineStackingStrategy = LineStackingStrategy.BlockLineHeight   '0
            'Send the document to the printer
            If PrintCaption = "" Or PrintCaption Is Nothing Or ThisPrinter = "" Then
                a1 = 1  'debug trap
            End If

            PrintCaption = "RogaineSCORE Printing " & PrintPageCount.ToString
            PrintPageCount = PrintPageCount + 1
            pDiag.PrintDocument(myPaginatorSource.DocumentPaginator, PrintCaption) 'pDiag.PrintDocument(CType(doc, IDocumentPaginatorSource).DocumentPaginator, printCaption)

        Catch ex As Exception
            PrintFailed = True
        End Try
    End Sub

    Private Sub SpaceTheTabs() 'Input is LineWithTabs, and returns TabSpacedLine
        'Where a tab is found, it is replaced one to 8 spaces up to a position which is a multiple of 8.
        Dim aReplace, bReplace, cReplace As Integer
        Dim CharPointer As Long
        Dim localChar As String

        SpacedTabLine = ""
        TabbedLine = StripLeadingSpaces(TabbedLine)
        aReplace = Len(TabbedLine)
        If aReplace = 0 Then Exit Sub
        CharPointer = 1
        For bReplace = 1 To aReplace
            'CharacterPointer = CharacterPointer + 1
            localChar = Strings.Left(Right(TabbedLine, aReplace - bReplace + 1), 1)  'Position in line is a-b+1
            If localChar = Chr(9) Then
                'TabSpacedLine = TabSpacedLine & " "
                'CharacterPointer = CharacterPointer + 1
                cReplace = CInt(9 - (CharPointer Mod 8))
                If cReplace = 9 Then cReplace = 1
                For J As Integer = 1 To cReplace
                    SpacedTabLine = SpacedTabLine & " "
                    CharPointer = CharPointer + 1
                Next J
            Else
                SpacedTabLine = SpacedTabLine & localChar
                CharPointer = CharPointer + 1
            End If
        Next bReplace
    End Sub

End Module
