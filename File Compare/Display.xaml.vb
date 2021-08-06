Imports System
Imports System.IO
Imports System.IO.Ports
Imports Microsoft.VisualBasic
Imports System.Timers
Imports System.Threading
Imports System.Windows.Forms
'Imports System.Drawing
Imports System.Drawing.Point    'Need to add System.Drawing.dll to references.

Imports System.Math
Imports Microsoft.VisualBasic.Strings
Imports System.Windows.Threading
Imports System.Globalization


Public Class Display
    Public Shared NewTitle As String
    Public Shared CallingForm As Object
    Public Shared LineCount, MenuMax As Integer
    Public Shared LabelText As String
    Public Shared EndOfScore As Integer    'The END line if coming from ReadTags 


    Public ZoomLineLength As Integer

    Public MouseEnabled As Boolean
    Public ReadTagsObject As Object
    Private Reply, KeyCode As Integer
    Private Button As Integer = 0
    Private Shift As Integer = 0
    Private X, Y As Single
    Private MyLabelText As String
    Private LinesVisible, MenuWidth, LinesMax, HoldLineLength As Integer
    Private PrintLines() As String
    Private Xleft, Ytop, Xwidth, Yheight As Double
    Private DLeft, DTop, DWidth, DHeight As Integer         'screen location and size
    Private X1Loc, Y1Loc, RedLine, EndLine, HoverLine, DisplayLineIndex As Integer
    Private WriteHeight As Single
    Private Timer1 As DispatcherTimer
    Private Timer1HandlerLoaded, MeLoaded As Boolean
    Private MyFont As Double
    Private MyDisplayLines(), OptionFileName As String
    Private MyREDlines, MyGREENlines, MyBLUElines, MyBROWNlines As String

    Private Sub Me_Sizechanged() Handles Me.SizeChanged     'called before Loaded !
        If Not MeLoaded Then Exit Sub
        Call Form_Resizing()
    End Sub
    Private Sub Form_LoadHandler() Handles Me.Loaded
        OptionFileName = DISPLAYDATAFILE        'the name of the file holding sacreen startup position
        Call Form_Load()
        If DISPLAYLINES IsNot Nothing Then
            MyDisplayLines = DISPLAYLINES
        End If
        ZoomLineLength = 0
        Me.Show()
    End Sub
    Private Sub Form_Load()
        Dim screenRight, screenLeft, screenBottom, screenTop As Integer
        LinesVisible = 16 'arbitrary default
        VScroll1.Minimum = 1        'max is determined by SHOWDISPLAY in CODE file.
        VScroll1.LargeChange = 8
        VScroll1.SmallChange = 1
        MenuWidth = 40  'default number of chars long.
        DisplayLineIndex = 1
        HoldLineLength = 0

        screenLeft = My.Computer.Screen.WorkingArea.Left   '0
        screenRight = My.Computer.Screen.WorkingArea.Right
        screenBottom = My.Computer.Screen.WorkingArea.Bottom
        screenTop = My.Computer.Screen.WorkingArea.Top   '1040
        'defaults
        DLeft = screenLeft + 10   '10
        DTop = CInt(screenTop + 10)
        DWidth = 548
        DHeight = 449

        Call LoadDisplaySettings()    'if saved from last time
        'Safety check.
        If DLeft < screenLeft Or DLeft > screenRight - DWidth Then
            DLeft = screenLeft + 10
            DWidth = 548
        End If
        If DTop < screenTop Or DTop > screenBottom - DHeight Then
            DTop = CInt(screenTop + 10)
            DHeight = 449
        End If
        If DLeft < 10 Then DLeft = 10
        If DTop < 10 Then DTop = 10
        If DWidth < 200 Then DWidth = 548
        If DHeight < 200 Then DHeight = 449
        Me.Left = DLeft
        Me.Top = DTop
        Me.Width = DWidth
        Me.Height = DHeight
        CLICKLINE = -10
        Timer1 = New DispatcherTimer
        AddHandler Timer1.Tick, AddressOf Timer1_Tick
        Timer1HandlerLoaded = True
        Timer1.Interval = TimeSpan.FromMilliseconds(20)
        Timer1.Start()      'form re-sizing
        If SCORECOLOURS Is Nothing Then ReDim SCORECOLOURS(0)
        MnuAutoFit.ToolTip = "If checked, font size is scaled so that the widest line will fit into the window." & vbCrLf &
            "As you re-size the window, the font scales automatically."

        MnuFixedWidth.ToolTip = ""
        MnuReduceWidth.ToolTip = "Zooms in by increasing the font size. Long Lines may be truncated." & vbCrLf &
            "Auto-font size is disabled."      'ZOOM IN
        MnuIncreaseWidth.ToolTip = "Zooms out by reducing the font size." & vbCrLf &
            "Auto-font size is disabled."           'ZOOM OUT
        MaimMenu.Visibility = Visibility.Visible
        Mnufile.Visibility = Visibility.Visible
        NewTitle = Me.Title
        MeLoaded = True
    End Sub

    Private Sub WindowMoved() Handles Me.LocationChanged
        If Not Timer1HandlerLoaded Then Exit Sub
        DLeft = CInt(Me.Left)
        DTop = CInt(Me.Top)
        DWidth = CInt(Me.Width)
        DHeight = CInt(Me.Height)
    End Sub
    Public Sub UpdateContent()
        MyDisplayLines = DISPLAYLINES
        MyREDlines = REDLINES
        MyGREENlines = GREENLINES
        MyBLUElines = BLUELINES
        MyBROWNlines = BROWNLINES
        Call Form_Resizing()
    End Sub
    Public Sub Form_Resizing()
        Dim localLoop, localPointer, fontWidth, fontSize As Integer
        Dim localstring, localWord, testString As String
        Dim textHeight, textLength, canvasHeight, canvasWidth As Double
        Dim localFlag As Boolean

        Dim localColor As Brush
        Dim localBrush As Brush
        Dim myCulture As System.Globalization.CultureInfo
        Dim formattedText As FormattedText


        If MyDisplayLines Is Nothing And DISPLAYLINES IsNot Nothing Then
            MyDisplayLines = DISPLAYLINES
        End If
        If MyREDlines Is Nothing Then MyREDlines = REDLINES
        If MyGREENlines Is Nothing Then MyGREENlines = GREENLINES
        If MyBLUElines Is Nothing Then MyBLUElines = BLUELINES
        If MyBROWNlines Is Nothing Then MyBROWNlines = BROWNLINES

        If Not Timer1HandlerLoaded Then Exit Sub
        Timer1.Stop()

        If MyDisplayLines Is Nothing Then Exit Sub
        myCulture = System.Globalization.CultureInfo.CurrentCulture
        canvasWidth = LabelCanvas.ActualWidth   '650 to 1258
        canvasHeight = LabelCanvas.ActualHeight   '416

        RedLine = -1    'redline is set to the line number of the InLines we want colored red.
        fontSize = 36
        If HoldLineLength = 0 Then HoldLineLength = DISPLAYLINELENGTH
        localBrush = Brushes.Black
        Do      'Try to get fontsize which just fits. emSize is device-independent = 1/96 inches. fontsize is size in points, each point is 1/72 inches. 
            testString = "A B C D E 1234567890"      '20 characters
            formattedText = New FormattedText(testString, myCulture, System.Windows.FlowDirection.LeftToRight, New Typeface("Lucida Console"), fontSize * 96.0 / 72.0, localBrush)
            'formattedText.SetFontSize(fontSize * (96.0 / 72.0), 0, Len(testString))
            formattedText.SetFontWeight(FontWeights.Normal, 0, Len(testString))
            textHeight = formattedText.Height   '21.333textLength = formattedText.Width    '2030.82  for fontsize = 32, 1015 for fontsize 16,3
            textLength = formattedText.Width    '643 for font 40, 578 for font 32 (emSize = 48). 48/96 * 20 = 10" wide = 720 pixels
            'fontWidth in pixels = (emSize/96)*72  =  (fontSize * 96.0 / 72.0) *72/96  = fontSize

            fontWidth = CInt(Math.Ceiling(textLength / 20))  'char width
            If fontSize * DISPLAYLINELENGTH <= canvasWidth Then
                If fontSize >= 40 Then DISPLAYLINELENGTH = HoldLineLength   '73 to 18
                Exit Do
            End If
            fontSize = fontSize - 1
            If fontSize <= 2 Then Exit Do
        Loop
        fontSize = CInt(Math.Round(1.7 * fontSize))   'debug fiddle
        Dim heightFiddle As Double = 0.74    'debug fiddle
        LinesVisible = CInt(Math.Floor((canvasHeight / textHeight) * heightFiddle) - 0)

        If MnuAutoFit.IsChecked = True Then MenuWidth = DISPLAYMAXCHARS
        LineCount = UBound(MyDisplayLines)
        localFlag = False
        If LineCount = 0 Then Exit Sub

        If LineCount <= LinesVisible Then
            VScroll1.Visibility = Visibility.Hidden
            DisplayLineIndex = 1
        Else
            VScroll1.Visibility = Visibility.Visible
            VScroll1.Maximum = UBound(MyDisplayLines) - LinesVisible '+ 7
            VScroll1.Value = DisplayLineIndex
            If (UBound(MyDisplayLines) > 0) And (UBound(MyDisplayLines) > LinesVisible) Then
                If DisplayLineIndex > UBound(MyDisplayLines) Then
                    DisplayLineIndex = CInt(VScroll1.Maximum) - LinesVisible + 1    'never seem to get here
                End If
            End If
        End If

        RedLine = -1
        EndLine = 0
        localstring = MyDisplayLines(1)
        Label1.Inlines.Clear()
        Label1.Text = ""
        Label1.Inlines.Add(New Run(localstring))        'header line

        For localLoop = DisplayLineIndex + 1 To DisplayLineIndex + LinesVisible 'UBound(DISPLAYLINES)
            localColor = New SolidColorBrush(Colors.Black)
            If localLoop > UBound(MyDisplayLines) Then Exit For
            localWord = MyDisplayLines(localLoop)
            If Strings.Left(localWord, 3) = "END" Then
                EndLine = localLoop
                MyDisplayLines(localLoop) = "END" & vbTab & vbTab & vbTab & vbTab & vbTab & vbCrLf 'make line longer so that mouseover highlights long line
            End If
            If localWord = "   " & vbCrLf And EndLine = 0 And localLoop > 4 Then
                EndLine = localLoop
            End If
            localPointer = localLoop - 4
            If localPointer <= UBound(SCORECOLOURS) And localPointer >= 0 Then
                localColor = SCORECOLOURS(localPointer)
                If localColor IsNot Nothing Then
                    If localColor.ToString = "" Then localColor = New SolidColorBrush(Colors.Black)
                Else
                    localColor = New SolidColorBrush(Colors.Black)
                End If
            End If
            BYTE8 = localLoop.ToString
            If InStr(MyREDlines, "," & BYTE8 & ",") <> 0 Then
                localColor = New SolidColorBrush(color:=Colors.Red)
            End If
            If InStr(MyBLUElines, "," & BYTE8 & ",") <> 0 Then
                localColor = New SolidColorBrush(color:=Colors.Blue)
            End If
            If InStr(MyGREENlines, "," & BYTE8 & ",") <> 0 Then
                localColor = New SolidColorBrush(color:=Colors.Green)
            End If
            If InStr(MyBROWNlines, "," & BYTE8 & ",") <> 0 Then
                localColor = New SolidColorBrush(color:=Colors.Brown)
            End If
            localstring = MyDisplayLines(localLoop)
            Label1.Inlines.Add(New Run(localstring))
            Label1.Inlines(Label1.Inlines.Count - 1).Foreground = localColor
        Next

        For localLoop = 0 To Label1.Inlines.Count - 1
            Label1.Inlines(localLoop).FontSize = fontSize
            If localLoop = RedLine Then Label1.Inlines(localLoop).Foreground = New SolidColorBrush(color:=Colors.Red)
        Next

        If Not Timer1HandlerLoaded Then Exit Sub
        DLeft = CInt(Me.Left)
        DTop = CInt(Me.Top)
        DWidth = CInt(Me.Width)
        DHeight = CInt(Me.Height)

    End Sub


    Private Sub Form_UnloadHandler(sender As Object, e As EventArgs) Handles Me.Closing
        Call SaveDisplaySettings()
        DISPLAYOBJECT1 = Nothing
    End Sub
    Public Sub SaveDisplaySettings()
        Dim fileNum As Integer
        Dim localFile As String
        Dim localFlag As Boolean

        If DLeft = 0 Then Exit Sub  'if closed from an external source and me is already closed.
        localFile = PARENTPATH & OptionFileName
        localFlag = FileIO.FileSystem.FileExists(localFile)
        If localFlag Then
            File.SetAttributes(localFile, FileAttributes.Normal)
            FileIO.FileSystem.DeleteFile(localFile)
        End If

        fileNum = FreeFile()
        FileOpen(fileNum, localFile, OpenMode.Output)
        Print(fileNum, "MyLeft = " & DLeft & vbCr & vbLf)
        Print(fileNum, "MyTop = " & DTop & vbCr & vbLf)
        Print(fileNum, "MyWidth = " & DWidth & vbCr & vbLf)
        Print(fileNum, "MyHeight = " & DHeight & vbCr & vbLf)
        FileClose((fileNum))
    End Sub
    Private Sub LoadDisplaySettings()
        Dim fileNum, localPos As Integer
        Dim localFile, fileLine, localChar, localNum As String
        Dim localFlag As Boolean

        localFile = PARENTPATH & OptionFileName
        localFlag = FileIO.FileSystem.FileExists(localFile)
        If localFlag Then
            fileNum = FreeFile()
            FileOpen(fileNum, localFile, OpenMode.Input)
            Do Until EOF(fileNum)
                fileLine = LineInput(fileNum)
                localPos = InStr(fileLine, " = ")
                If localPos <> 0 Then
                    localChar = Strings.Left(fileLine, localPos - 1)    'MyLeft, MyTop, MyWidth,MyHeight
                    localNum = Trim(Strings.Right(fileLine, Len(fileLine) - localPos - 1))
                    Select Case localChar
                        Case "MyLeft"
                            DLeft = CInt(localNum)
                        Case "MyTop"
                            DTop = CInt(localNum)
                        Case "MyWidth"
                            DWidth = CInt(localNum)
                        Case "MyHeight"
                            DHeight = CInt(localNum)
                    End Select
                End If
            Loop
            FileClose(fileNum)  'BYTE3 should contain last PUNCH line. Should start with a space unless Tag is empty.
        End If
    End Sub

    Private Sub MnuAutoFit_ClickHandler() Handles MnuAutoFit.Click
        Call MnuAutoFit_Click()
    End Sub
    Private Sub MnuAutoFit_Click()
        If MnuAutoFit.IsChecked = True Then
            DISPLAYLINELENGTH = DISPLAYMAXCHARS
            Call Form_Resizing()
        End If
    End Sub
    Private Sub MnuCLOSE_ClickHandler() Handles MnuCLOSE.Click
        Me.Close()
    End Sub
    Private Sub MnuIncreaseWidth_ClickHandler() Handles MnuIncreaseWidth.Click      'ZOOM OUT
        Call MnuIncreaseWidth_Click()
    End Sub
    Private Sub MnuIncreaseWidth_Click()    'makes text smaller.
        MnuAutoFit.IsChecked = False
        DISPLAYLINELENGTH = DISPLAYLINELENGTH + CInt(DISPLAYLINELENGTH * 0.25)
        ZoomLineLength = DISPLAYLINELENGTH
        'If DISPLAYLINELENGTH > DISPLAYMAXCHARS Then DISPLAYLINELENGTH = DISPLAYMAXCHARS
        Call Form_Resizing()
    End Sub
    Private Sub MnuReduceWidth_ClickHandler() Handles MnuReduceWidth.Click      'ZOOM IN
        Call MnuReduceWidth_Click()
    End Sub
    Private Sub MnuReduceWidth_Click()      'makes text larger
        MnuAutoFit.IsChecked = False
        HoldLineLength = DISPLAYLINELENGTH
        DISPLAYLINELENGTH = DISPLAYLINELENGTH - CInt(DISPLAYLINELENGTH * 0.2)
        ZoomLineLength = DISPLAYLINELENGTH
        Call Form_Resizing()
    End Sub
    Private Sub ZoomIn() Handles MnuZOOMin.Click
        Call MnuReduceWidth_Click()
    End Sub
    Private Sub MnuPrint2Printer() Handles MnuPrintFile.Click
        Dim localLoop, fileNum As Integer
        Dim tempLine, localFile As String
        ReDim DEBUGPRINTLINES(0)
        For localLoop = 1 To UBound(MyDisplayLines)
            tempLine = (MyDisplayLines(localLoop))
            tempLine = Strings.Left(tempLine, Len(tempLine) - 2) 'remove linefeeds
            Do
                If Strings.Right(tempLine, 1) <> " " Then
                    Exit Do
                Else
                    tempLine = Strings.Left(tempLine, Len(tempLine) - 1) 'remove trailing spaces.
                End If
            Loop
            ReDim Preserve DEBUGPRINTLINES(UBound(DEBUGPRINTLINES) + 1)
            DEBUGPRINTLINES(UBound(DEBUGPRINTLINES)) = tempLine
        Next localLoop
        'Write to file ScreenDump.txt
        localFile = PARENTPATH & "\ScreenDump.txt"
        fileNum = FreeFile()
        FileOpen(fileNum, localFile, OpenMode.Output)
        For localLoop = 1 To UBound(DEBUGPRINTLINES)
            tempLine = DEBUGPRINTLINES(localLoop)
            Print(fileNum, tempLine & vbCr & vbLf)
        Next localLoop
        FileClose((fileNum))
        SPEAK("ding")
    End Sub
    Private Sub MnuPrint_ClickHandler() Handles MnuPrint.Click
        Call MnuPrint_Click()
    End Sub
    Private Sub MnuPrint_Click()
        Dim localLoop As Integer
        Dim tempLine As String
        If UBound(MyDisplayLines) < 1 Then Exit Sub
        ReDim PrintLines(UBound(MyDisplayLines))
        ReDim DEBUGPRINTLINES(0)
        STORERESULTS = True 'send to default printer
        For localLoop = 1 To UBound(MyDisplayLines)
            tempLine = (MyDisplayLines(localLoop))
            tempLine = Strings.Left(tempLine, Len(tempLine) - 2) 'remove linefeeds
            Do
                If Strings.Right(tempLine, 1) <> " " Then
                    Exit Do
                Else
                    tempLine = Strings.Left(tempLine, Len(tempLine) - 1) 'remove trailing spaces.
                End If
            Loop
            PrintLines(localLoop) = tempLine
            SENDPRINT(tempLine)
        Next localLoop
        ENDPRINT()

        Call DoDummyPrint()

    End Sub
    Private Sub MnuClearScreen_Click(sender As Object, e As EventArgs) Handles MnuClearScreen.Click
        ReDim MyDisplayLines(1)
        MyDisplayLines(1) = "                                        "
        DISPLAYMAXCHARS = 40
        DISPLAYLINELENGTH = DISPLAYMAXCHARS
        DisplayLineLength = 1
        Call Form_Resizing()
    End Sub

    Private Sub VScroll1_ChangeHandler() Handles VScroll1.ValueChanged
        Call ScrollShow()
    End Sub
    Private Sub Vscroll1_Scroll() Handles VScroll1.Scroll
        Call ScrollShow()
    End Sub
    Private Sub ScrollShow()        'Have DISPLAYLINES as source, and LinesVisible as how many we can show.
        'Maybe also show first line and last line.
        If MyDisplayLines IsNot Nothing Then
            DisplayLineIndex = CInt(VScroll1.Value)
            Dim localString, stringShort As String
            stringShort = ""
            localString = ""
            If DisplayLineIndex + LinesVisible > UBound(MyDisplayLines) - 0 Then
                DisplayLineIndex = UBound(MyDisplayLines) - LinesVisible + 1
                If DisplayLineIndex < 1 Then DisplayLineIndex = 1
            End If
            Call Form_Resizing()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Call Form_Resizing()
    End Sub
    Private Sub Mouse_scroll(sender As Object, e As MouseWheelEventArgs) Handles Me.MouseWheel

        If e.Delta > 0 Then
            If VScroll1.Value > VScroll1.Minimum Then VScroll1.Value = VScroll1.Value - 1
        Else
            If VScroll1.Value < VScroll1.Maximum Then VScroll1.Value = VScroll1.Value + 1
        End If
    End Sub
    Private Sub NotUsedExamples()       'A storage for useful experimental code used during development
        Dim localWord, testString As String
        Dim myColor As Color
        Dim localBrush As Brush
        Dim textLen, lineNum As Integer
        Dim textHeight, textLength, canvasHeight, canvasWidth As Double
        Dim myCulture As System.Globalization.CultureInfo
        Dim formattedText As FormattedText
        Dim localFlag As Boolean
        Dim mySize As Double

        myCulture = System.Globalization.CultureInfo.CurrentCulture
        canvasWidth = LabelCanvas.ActualWidth   '650 to 1258
        canvasHeight = LabelCanvas.ActualHeight   '416

        RedLine = -1

        Label1.Inlines.Clear()
        Label1.Text = ""
        For localLoop = 1 To UBound(MyDisplayLines)
            localWord = MyDisplayLines(localLoop)
            Label1.Inlines.Add(New Run(localWord))
        Next
        testString = "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor"
        MyFont = 16     'fontsize   'default test
        myColor = Colors.Black
        ' Create the initial formatted text string.
        localBrush = Brushes.Black
        formattedText = New FormattedText(testString, myCulture, System.Windows.FlowDirection.LeftToRight, New Typeface("Lucida Console"), MyFont * 96.0 / 72.0, localBrush)
        ' Use a larger font size beginning at the first (zero-based) character and continuing for 5 characters.
        ' The font size is calculated in terms of points -- not as device-independent pixels.
        formattedText.SetFontSize(MyFont * (96.0 / 72.0), 0, textLen)
        ' Use a Bold font weight beginning at the 6th character and continuing for 11 characters.
        formattedText.SetFontWeight(FontWeights.Normal, 0, textLen)
        ' Use a linear gradient brush beginning at the 6th character and continuing for 11 characters.
        'formattedText.SetForegroundBrush(New LinearGradientBrush(Colors.Orange, Colors.Teal, 90.0), 0, textLen)
        ' Use an Italic font style beginning at the 28th character and continuing for 28 characters.
        'formattedText.SetFontStyle(FontStyles.Italic, 28, 28)

        textHeight = formattedText.Height   '21.333textLength = formattedText.Width    '2030.82  for fontsize = 32, 1015 for fontsize 16,3
        textLength = formattedText.Width    '2030.82  for fontsize = 32, 1015 for fontsize 16,

        'examples of fromatting individual lines.
        Label1.Inlines.Add(New Run("  The first line" & vbCrLf))
        Label1.Inlines.Add(New Run("  1. Not Bold text" & vbCrLf))
        Label1.Inlines.Add(New Bold(New Run("  2. is Bold text" & vbCrLf)))
        lineNum = Label1.Inlines.Count - 1
        Label1.Inlines(lineNum).Foreground = New SolidColorBrush(color:=Colors.Green)
        'Label1.Foreground = New SolidColorBrush(color:=Colors.Red)      'applies to entire block.
        ''Label1.Inlines.f(New SolidColorBrush(color:=Colors.Red))
        'Label1.Foreground = CType(Brushes.Navy, System.Media.)
        Label1.FontSize = 30        'applies to whole block
        'Label1.FontFamily = FontFamilyProperty.
        Label1.Inlines.Add(New Run("  3. random is not bold text" & vbCrLf))
        Label1.Inlines.Add(New Run("  4. random is not bold text" & vbCrLf))
        Label1.FontSize = 16       'over-rides the earlier setting and applies to whole block
        Label1.Inlines.Add(New Bold(New Run("  5. Hello my Second bold Run" & vbCrLf)))
        Label1.Inlines.Add(New Bold(New Run("  6. Hello my THird bold Run" & vbCrLf)))

        For localLoop = 0 To Label1.Inlines.Count - 1       'highlight the background of each line as the mouse is over it.
            'ability to color eah line individually
            If localLoop = 9 Then
                Label1.Inlines(localLoop).Foreground = New SolidColorBrush(color:=Colors.Red)
                Label1.Inlines(localLoop).FontSize = 24
            End If
            If localLoop = 6 Then
                Label1.Inlines(localLoop).Foreground = New SolidColorBrush(color:=Colors.Blue)
                Label1.Inlines(localLoop).FontSize = 18
                mySize = FontFamily.LineSpacing 'Always equals 1
            End If
            If localLoop = 7 Then
                Label1.Inlines(localLoop).Foreground = New SolidColorBrush(color:=Colors.Green)
                Label1.Inlines(localLoop).FontSize = 12
                localFlag = Label1.Inlines(localLoop).IsMouseOver
            End If
            If localLoop = 8 Then
                Label1.Inlines(localLoop).Foreground = New SolidColorBrush(color:=Colors.Magenta)
                Label1.Inlines(localLoop).FontSize = 10
            End If
        Next
    End Sub

    Private Sub Display_Activated() Handles Me.Activated
        If MeLoaded Then Me.Title = NewTitle
    End Sub
End Class

