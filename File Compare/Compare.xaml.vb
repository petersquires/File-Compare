Imports System.Windows
Imports System.IO
Imports System.ComponentModel
Imports System.Printing     'NOTE: Had to add a reference in Project... Add Reference...

Class MainWindow
    Private HoldPath1, HoldPath2, FileTypes, SortSequence, ClickedFileTypes, ChosenFileTypes, SavedHistory() As String
    Private reply As MsgBoxResult
    Private OptionsFileName As String
    Private Pointer1, Pointer2 As Integer
    Private FileToggle As Boolean


    Private Structure FileData
        Public FileFullName As String       'full name with extension but no path
        Public FileName As String       'full name with extension but no path
        Public ShortName As String      'Name without extension
        Public Extension As String
        Public FileDate As Date
    End Structure
    Private Files() As FileData
    Private AllFiles1() As FileData
    Private AllFiles2() As FileData
    Private Files1() As FileData
    Private Files2() As FileData
    Private Files2Sort(), SortedFiles() As FileData


    Private MyRatio, ExpandedRatio, TitleHeight, ExtraHeight, ExtraWidth, FirstWidth, Magnification As Double
    Private FormJustLoaded, DontChangeDisplayOnly As Boolean

    Private Sub Me_Sizechanged() Handles Me.SizeChanged
        Dim localHeight As Double
        If Not FormJustLoaded Then Exit Sub
        localHeight = (MyRatio * Width + TitleHeight)
        Height = localHeight

    End Sub

    Private Sub Form_LoadHandler() Handles Me.Loaded
        Call Form_Load()
        FileExtensions.Visibility = Visibility.Hidden   'initial default
        ReDim Files(0)
        ReDim AllFiles1(0)
        ReDim AllFiles2(0)
        ReDim Files1(0)
        ReDim Files2(0)
        ReDim Files2Sort(0)
        ReDim SortedFiles(0)

        ReDim SavedHistory(0)
        OptionsFileName = "CompareOptions.txt" 'defaut
        Call ReadHistory()
        If OptionsFileName <> "CompareOptions.txt" Then Call ReadHistory()
        ListBox1.SelectedIndex = Pointer1
        ListBox2.SelectedIndex = Pointer2
        If SOURCEPATH1 <> "" Then
            PATH1.Text = SOURCEPATH1
            Call GetFiles1()    'Loads AllFiles1() with ALL files in SOURCEPATH1.
            Call LoadFiles1() 'Loads Files1() from AllFiles1(), and ChosenFileTypes 
        End If
        If SOURCEPATH2 <> "" Then
            PATH2.Text = SOURCEPATH2
            Call GetFiles2()    'Loads AllFiles2() with ALL files in SOURCEPATH1.
            Call LoadFiles2() 'Loads Files2() from AllFiles2(), and ChosenFileTypes 
        End If
        If ChosenFileTypes <> ",*," Then
            SelectedFilesOption.IsChecked = True
            Call SelectedFilesChosen()
        End If
        'Now find the default printer, and load into line 4.
        Dim pDiag As New PrintDialog()
        ChosenDefault = pDiag.PrintQueue.Name            '"Brother DCP-J4120DW Printer", "Send To OneNote 2016", "Foxit PhantomPDF Printer"

    End Sub
    Private Sub Form_Load()  'Width="415" Height="358" Ratio = 1.159
        'For Window W/H Ratio   = 0.6; 1.0;  1.16,  1.2, 1.39; 1.5; 1.66; 1.83; 2.0, 
        'TitleHeight            = 6,   26,   27,   28,  32,   34,  36,   39,   41
        'Border                 = ?,   ? ,   1,    2,   2,   3,    6,    5,    6
        Dim borderHeight As Double
        TitleHeight = 26
        borderHeight = 1
        Height = Height + borderHeight  'BackCanvas.Height ??
        MyRatio = (Height - TitleHeight) / (Width)  '(height + borderHeight - TitleHeight)
        FormJustLoaded = True
        PARENTPATH = FileIO.FileSystem.CurrentDirectory
        HoldPath1 = PARENTPATH  'default
        HoldPath2 = PARENTPATH  'default
    End Sub

    Private Sub TryCompare() Handles CompareFiles.Click
        Dim localNum, localPos As Integer
        Dim worD1, worD2 As String
        SELECTEDFILE1 = ""
        SELECTEDFILE2 = ""
        Try
            worD1 = ListBox1.SelectedItem.ToString  'includes date. Must remove.
            worD2 = ListBox2.SelectedItem.ToString
            localNum = ListBox1.SelectedIndex + 1
            worD1 = Files1(localNum).FileName
            localNum = ListBox2.SelectedIndex + 1
            worD2 = Files2(localNum).FileName
        Catch ex As Exception
            MsgBox("A file must be selected for each list.", vbCritical, "SELECT AN ITEM")
            Exit Sub
        End Try
        For localNum = 1 To UBound(Files1)
            If Files1(localNum).FileName = worD1 Then
                SELECTEDFILE1 = Files1(localNum).FileFullName
                FIRSTSELECTED = SELECTEDFILE1
                BYTE8 = Files1(localNum).FileDate.ToShortDateString
                BYTE9 = Files1(localNum).FileDate.ToLongDateString
                BYTEA = Files1(localNum).FileDate.ToLongTimeString
                localPos = InStr(BYTE9, ",")
                BYTE9 = Strings.Right(BYTE9, Len(BYTE9) - localPos)
                FILE1DATE = BYTE8 & "   " & BYTEA
                'FILE1DATE = BYTE9 & ",   " & BYTEA
            End If
        Next
        For localNum = 1 To UBound(Files2)
            If Files2(localNum).FileName = worD2 Then
                SELECTEDFILE2 = Files2(localNum).FileFullName
                SECONDSELECTED = SELECTEDFILE2
                BYTE8 = Files2(localNum).FileDate.ToShortDateString
                BYTE9 = Files2(localNum).FileDate.ToLongDateString
                BYTEA = Files2(localNum).FileDate.ToLongTimeString
                localPos = InStr(BYTE9, ",")
                BYTE9 = Strings.Right(BYTE9, Len(BYTE9) - localPos)
                FILE2DATE = BYTE8 & "   " & BYTEA
                'FILE2DATE = BYTE9 & ",   " & BYTEA
            End If
        Next
        SwapButton.Visibility = Visibility.Visible
        FileToggle = False
        REACH = CInt(ReachValue.Text)
        SIZEOFBLOCK = CInt(BlockValue.Text)
        SynchB.Content = "SynchB ??"
        Call TryDisplay()  '        Calls COMPAREFILES1() in module CODE
        If SYNCHB_USED Then SynchB.Content = SYNCHBRECORD ' "SynchB used."
    End Sub
    Private Sub TryDisplay() Handles SwapButton.Click
        IGNORESPACES = False
        SHOWTABS = False
        SKIPBLANKLINES = False
        IGNORETABS = False
        If NoSpace.IsChecked Then IGNORESPACES = True
        If TabsShow.IsChecked Then SHOWTABS = True
        If BlankSkip.IsChecked Then SKIPBLANKLINES = True
        If NoTab.IsChecked Then IGNORETABS = True

        If FileToggle Then
            SELECTEDFILE1 = SECONDSELECTED
            SELECTEDFILE2 = FIRSTSELECTED
            SwapButton.Background = New SolidColorBrush(Colors.Red)
            SwapButton.Foreground = New SolidColorBrush(Colors.Black)
            SwapButton.Content = "BACK"
        Else
            SELECTEDFILE1 = FIRSTSELECTED
            SELECTEDFILE2 = SECONDSELECTED
            SwapButton.Background = New SolidColorBrush(Colors.Blue)
            SwapButton.Foreground = New SolidColorBrush(Colors.White)
            SwapButton.Content = "SWAP"
        End If
        FileToggle = Not FileToggle
        Call COMPAREFILES1()    'In module Code.vb
    End Sub

    Private Sub Save_as() Handles MnuSAVEAS.Click
        Dim dialogFlag? As Boolean
        Dim localPos, fileNum As Integer
        Dim localWord, localLine As String
        Dim myDialog As New Microsoft.Win32.SaveFileDialog

        myDialog.InitialDirectory = PARENTPATH
        myDialog.Title = "Enter a new file name and click SAVE..."
        myDialog.FileName = OptionsFileName
        dialogFlag = myDialog.ShowDialog()
        If dialogFlag Then
            localWord = myDialog.FileName      'Selected file and path
            localPos = InStrRev(localWord, "\",, CompareMethod.Binary)
            OptionsFileName = Strings.Right(localWord, Len(localWord) - localPos)
            Call LoadSaveHistory()
            fileNum = FreeFile()
            FileOpen(fileNum, localWord, OpenMode.Output)
            For localLoop = 1 To UBound(SavedHistory)
                localLine = SavedHistory(localLoop)
                Print(fileNum, localLine & vbCr & vbLf)
            Next localLoop
            FileClose(fileNum)
        End If
    End Sub
    Private Sub Save_Existing() Handles MnuSAVE.Click
        Dim fileNum As Integer
        Dim localLine, localWord As String

        localWord = PARENTPATH & "\" & OptionsFileName
        Call LoadSaveHistory()
        fileNum = FreeFile()
        FileOpen(fileNum, localWord, OpenMode.Output)
        For localLoop = 1 To UBound(SavedHistory)
            localLine = SavedHistory(localLoop)
            Print(fileNum, localLine & vbCr & vbLf)
        Next localLoop
        FileClose(fileNum)
    End Sub
    Private Sub LoadOptions() Handles MnuOPEN.Click
        Dim dialogFlag? As Boolean
        Dim localPos As Integer
        Dim localWord, localFile As String

        Dim myDialog As New Microsoft.Win32.OpenFileDialog With {
            .InitialDirectory = PARENTPATH,
            .Multiselect = False,
            .DefaultExt = "txt",
            .Filter = "TXT files (*.txt)|*.txt",
            .Title = "DOUBLE_CLICK the file you want to open."
        }
        myDialog.DefaultExt = "txt"
        myDialog.FileName = OptionsFileName 'the name to search for
        myDialog.ShowReadOnly = False

        dialogFlag = myDialog.ShowDialog()
        If dialogFlag Then
            localWord = myDialog.FileName      'Selected file and path
            localPos = InStrRev(localWord, "\",, CompareMethod.Binary)
            localFile = Strings.Right(localWord, Len(localWord) - localPos)
            OptionsFileName = localFile
            Call ReadHistory()  'Reads OptionsFileName
            Call ShowSourcePath1()
            Call ShowSourcePath2()
            Call GetFiles1()    'Loads AllFiles1() with ALL files in SOURCEPATH1.
            Call LoadFiles1() 'Loads Files1() from AllFiles1(), and ChosenFileTypes 
            Call GetFiles2()    'Loads AllFiles1() with ALL files in SOURCEPATH1.
            Call LoadFiles2() 'Loads Files1() from AllFiles1(), and ChosenFileTypes 
            ListBox1.SelectedIndex = Pointer1
            ListBox2.SelectedIndex = Pointer2
            OptionFile.Content = localFile
            OptionsFileName = localFile
        End If
    End Sub
    Private Sub PATH1Clicked() Handles PATH1.PreviewMouseDoubleClick  'The PATH1 box
        Call ChoosePath1()
    End Sub
    Private Sub ChoosePath1() Handles PATH1Select.PreviewMouseLeftButtonUp 'the LABEL
        'Must Add reference to project in Project Properties..
        'C:\Program Files(x86)\Reference assemblies\microsoft\Framework\.NETframework\v4.6.1\System.Windows.Forms.dll

        'Dim myDialog As New Microsoft.Win32.OpenFileDialog
        Dim myDialog As New System.Windows.Forms.FolderBrowserDialog
        Dim reSult As String

        'FileTypes = ",VB,XAML,TXT,ASM,"      'Must choose this later from a selection box.
        myDialog.Description = "DOUBLE_CLICK the folder for Source 1."
        myDialog.ShowNewFolderButton = False   'or True
        myDialog.RootFolder = Environment.SpecialFolder.Desktop
        myDialog.SelectedPath = HoldPath1
        reSult = myDialog.ShowDialog().ToString      'Returns Cancel or OK

        If reSult = "Cancel" Then
            Exit Sub
        Else
            Search1.Text = ""
            Search2.Text = ""
            SOURCEPATH1 = myDialog.SelectedPath
            HoldPath1 = SOURCEPATH1
            Call ShowSourcePath1()
            Call FindFileTypes()
            Call GetFiles1()    'Loads AllFile1() with all files in SOURCEPATH1.
            'now load the listbox with the sorted files. Need to define a sort order; txt,vb,asm etc, then sort with date
            Call LoadFiles1() 'Loads Files1() from AllFiles1(), if extensions in ChosenFileTypes match.
        End If
    End Sub
    Private Sub PATH2Clicked() Handles PATH2.PreviewMouseDoubleClick  'The PATH2 box
        Call ChoosePath2()
    End Sub
    Private Sub ChoosePath2() Handles PATH2Select.PreviewMouseLeftButtonUp 'the LABEL
        'Must Add reference to project in Project Properties..
        'C:\Program Files(x86)\Reference assemblies\microsoft\Framework\.NETframework\v4.6.1\System.Windows.Forms.dll

        'Dim myDialog As New Microsoft.Win32.OpenfileDialog
        Dim myDialog As New System.Windows.Forms.FolderBrowserDialog
        Dim reSult As String

        'FileTypes = ",VB,XAML,TXT,ASM,"      'Must choose this later from a selection box.
        myDialog.Description = "DOUBLE_CLICK the folder for Source 2."
        myDialog.ShowNewFolderButton = False   'or True
        myDialog.SelectedPath = HoldPath2
        myDialog.RootFolder = Environment.SpecialFolder.Desktop
        reSult = myDialog.ShowDialog().ToString      'Returns Cancel or OK

        If reSult = "Cancel" Then
            Exit Sub
        Else
            Search1.Text = ""
            Search2.Text = ""
            SOURCEPATH2 = myDialog.SelectedPath
            HoldPath2 = SOURCEPATH2
            Call ShowSourcePath2()
            Call FindFileTypes()
            Call GetFiles2()    'Loads AllFiles2() with all files in SOURCEPATH2.
            'now load the listbox with the sorted files. Need to define a sort order; txt,vb,asm etc, then sort with date
            Call LoadFiles2() 'Loads Files2() from AllFiles2(), if extensions in ChosenFileTypes match.
        End If
    End Sub

    Private Sub HideTags() Handles NoTab.Click
        TabsShow.IsChecked = False
    End Sub
    Private Sub DisplayTabs() Handles TabsShow.Click
        NoTab.IsChecked = False
    End Sub

    Private Sub ShowSourcePath1()
        Dim localPos, localPos2, lineLengthMax As Integer
        Dim localName, localShort As String
        'Load Folder into PATH1 Textbox.
        lineLengthMax = 30 'Max. characters in PATH1 text box.
        localName = SOURCEPATH1
        localShort = ""
        localPos = 1
        Do
            localPos2 = Strings.InStr(localPos, localName, "\")     'start from localpos2
            If localPos2 > lineLengthMax Then
                localShort = localShort & Strings.Left(localName, localPos - 2) & vbCrLf
                localName = Strings.Right(localName, Len(localName) - localPos + 2)
                localPos = 1
            ElseIf localPos2 = 0 Then
                localShort = localShort & localName
                Exit Do
            Else
                localPos = localPos2 + 1
            End If
        Loop
        PATH1.Text = localShort
    End Sub
    Private Sub ShowSourcePath2()
        Dim localPos, localPos2, lineLengthMax As Integer
        Dim localName, localShort As String
        'Load Folder into PATH1 Textbox.
        lineLengthMax = 30 'Max. characters in PATH1 text box.
        localName = SOURCEPATH2
        localShort = ""
        localPos = 1
        Do
            localPos2 = Strings.InStr(localPos, localName, "\")     'start from localpos2
            If localPos2 > lineLengthMax Then
                localShort = localShort & Strings.Left(localName, localPos - 2) & vbCrLf
                localName = Strings.Right(localName, Len(localName) - localPos + 2)
                localPos = 1
            ElseIf localPos2 = 0 Then
                localShort = localShort & localName
                Exit Do
            Else
                localPos = localPos2 + 1
            End If
        Loop
        PATH2.Text = localShort
    End Sub

    Private Sub AllTypesChosen() Handles AllFilesOption.Click
        FileExtensions.Visibility = Visibility.Hidden
    End Sub
    Private Sub SelectedFilesChosen() Handles SelectedFilesOption.Click
        FileExtensions.Visibility = Visibility.Visible
    End Sub
    Private Sub FindFileTypes() 'If SelectedFilesOption is checked, Searches Check1 to Check12, plus text1, for desired file extensions.
        'If AllFilesOption is checked, desired files are *.*
        'Returns types in ChosenFileTypes as comma separated extensions.
        'SelectedFilesOption chosen.
        Dim localWord As String
        Dim localPos As Integer

        ClickedFileTypes = ","
        If Check1.IsChecked Then ClickedFileTypes = ClickedFileTypes & "txt," 'Check1.Content & ","
        If Check2.IsChecked Then ClickedFileTypes = ClickedFileTypes & "vb,"
        If Check3.IsChecked Then ClickedFileTypes = ClickedFileTypes & "xaml,"
        If Check4.IsChecked Then ClickedFileTypes = ClickedFileTypes & "ach,"

        If Check5.IsChecked Then ClickedFileTypes = ClickedFileTypes & "asm,"
        If Check6.IsChecked Then ClickedFileTypes = ClickedFileTypes & "inc,"
        If Check7.IsChecked Then ClickedFileTypes = ClickedFileTypes & "mcs,"
        If Check8.IsChecked Then ClickedFileTypes = ClickedFileTypes & "mcp,"

        If Check9.IsChecked Then ClickedFileTypes = ClickedFileTypes & "ino,"
        If Check10.IsChecked Then ClickedFileTypes = ClickedFileTypes & "raw,"
        If Check11.IsChecked Then ClickedFileTypes = ClickedFileTypes & "c,"
        If Check12.IsChecked Then ClickedFileTypes = ClickedFileTypes & "h,"
        'Look at Text1.text
        localWord = Text1.Text
        If localWord <> "" Then
            localPos = InStrRev(localWord, ".")
            If localPos <> 0 Then localWord = Strings.Right(localWord, Len(localWord) - localPos)
            ClickedFileTypes = ClickedFileTypes & localWord & ","
        End If

        ChosenFileTypes = ClickedFileTypes
        If AllFilesOption.IsChecked Then
            ChosenFileTypes = ",*,"
        End If
    End Sub
    Private Sub GetFiles1() 'Creates Files() from SOURCEPATH1; only if in ChosenFileTypes 
        Dim localFullName, localName, fileExt, localShort As String
        Dim localPos As Integer
        Dim localDate As Date

        If SOURCEPATH1 = "" Then Exit Sub
        Search1.Text = ""
        ReDim AllFiles1(0)
        Dim fileList As Array
        fileList = Directory.GetFiles(SOURCEPATH1)
        For Each fileName As String In fileList   'fileName gives full path
            localFullName = fileName    'gives full path.   FileFullName
            localPos = InStrRev(localFullName, "\", , CompareMethod.Binary)
            localName = Strings.Right(localFullName, Len(localFullName) - localPos)  'FileName with extension but no path
            localPos = InStrRev(localName, ".", -1, CompareMethod.Binary)
            If localPos <> 0 Then
                fileExt = Strings.Right(localName, Len(localName) - localPos)
                'If InStr(UCase(ChosenFileTypes), "," & UCase(fileExt) & ",") <> 0 Or ChosenFileTypes = ",*," Then
                localShort = Strings.Left(localName, localPos - 1)  'Name without extension
                    localDate = FileIO.FileSystem.GetFileInfo(localFullName).LastWriteTime
                ReDim Preserve AllFiles1(UBound(AllFiles1) + 1)
                With AllFiles1(UBound(AllFiles1))
                    .FileFullName = localFullName
                    .FileName = localName       'A_Version.txt
                    .ShortName = localShort     'A_Version
                    .Extension = fileExt        'txt
                    .FileDate = localDate
                End With
                'End If
            End If
        Next fileName
    End Sub
    Private Sub GetFiles2() 'Creates Files() from SOURCEPATH1; only if in ChosenFileTypes 
        Dim localFullName, localName, fileExt, localShort As String
        Dim localPos As Integer
        Dim localDate As Date

        If SOURCEPATH2 = "" Then Exit Sub
        Search2.Text = ""
        ReDim AllFiles2(0)
        Dim fileList As Array
        fileList = Directory.GetFiles(SOURCEPATH2)
        For Each fileName As String In fileList   'fileName gives full path
            localFullName = fileName    'gives full path.   FileFullName
            localPos = InStrRev(localFullName, "\", , CompareMethod.Binary)
            localName = Strings.Right(localFullName, Len(localFullName) - localPos)  'FileName with extension but no path
            localPos = InStrRev(localName, ".", -1, CompareMethod.Binary)
            If localPos <> 0 Then
                fileExt = Strings.Right(localName, Len(localName) - localPos)
                'If InStr(UCase(ChosenFileTypes), "," & UCase(fileExt) & ",") <> 0 Or ChosenFileTypes = ",*," Then
                localShort = Strings.Left(localName, localPos - 1)  'Name without extension
                localDate = FileIO.FileSystem.GetFileInfo(localFullName).LastWriteTime
                ReDim Preserve AllFiles2(UBound(AllFiles2) + 1)
                With AllFiles2(UBound(AllFiles2))
                    .FileFullName = localFullName
                    .FileName = localName       'A_Version.txt
                    .ShortName = localShort     'A_Version
                    .Extension = fileExt        'txt
                    .FileDate = localDate
                End With
                'End If
            End If
        Next fileName
    End Sub
    Private Sub UpdateFiles1() Handles ListBox1.MouseDoubleClick
        Call FindFileTypes()  'Loads ChosenFileTypes as comma-separated string. "*," or "txt,vb,ach," etc.
        Call LoadFiles1()
    End Sub
    Private Sub LoadFiles1() 'Loads Files1() from AllFiles1(), if extensions in ChosenFileTypes match.
        Dim localNum As Integer
        Dim localExt As String

        Search1.Text = ""
        ReDim Files1(0)
        For localNum = 1 To UBound(AllFiles1)
            localExt = "," & AllFiles1(localNum).Extension & ","
            If InStr(ChosenFileTypes, localExt) <> 0 Or ChosenFileTypes = ",*," Then
                ReDim Preserve Files1(UBound(Files1) + 1)
                Files1(UBound(Files1)) = AllFiles1(localNum)
            End If
        Next
        'now load the files into Listbox1.
        Call LoadList1()
    End Sub
    Private Sub UpdateFiles2() Handles ListBox2.MouseDoubleClick
        Call FindFileTypes()  'Loads ChosenFileTypes as comma-separated string. "*," or "txt,vb,ach," etc.
        Call LoadFiles2()
    End Sub
    Private Sub LoadFiles2() 'Loads Files1() from AllFiles2(), if extensions in ChosenFileTypes match.
        Dim localNum As Integer
        Dim localExt As String

        Search2.Text = ""
        ReDim Files2(0)
        For localNum = 1 To UBound(AllFiles2)
            localExt = "," & AllFiles2(localNum).Extension & ","
            If InStr(ChosenFileTypes, localExt) <> 0 Or ChosenFileTypes = ",*," Then
                ReDim Preserve Files2(UBound(Files2) + 1)
                Files2(UBound(Files2)) = AllFiles2(localNum)
            End If
        Next
        'now load the files into Listbox1.
        Call LoadList2()
    End Sub
    Private Sub List1Changed() Handles ListBox1.SelectionChanged
        If ListBox1.SelectedItem IsNot Nothing And ListBox2.SelectedItem IsNot Nothing Then
            CompareFiles.Visibility = Visibility.Visible
        Else
            CompareFiles.Visibility = Visibility.Hidden
        End If
    End Sub
    Private Sub List2Changed() Handles ListBox2.SelectionChanged
        If ListBox1.SelectedItem IsNot Nothing And ListBox2.SelectedItem IsNot Nothing Then
            CompareFiles.Visibility = Visibility.Visible
        Else
            CompareFiles.Visibility = Visibility.Hidden
        End If
    End Sub

    Private Sub Search1_Enter(sender As Object, e As KeyEventArgs) Handles Search1.PreviewKeyUp
        If e.Key = Key.Enter Then
            Call SearchList1()
        End If
    End Sub
    Private Sub Search1_DoubleClick() Handles Search1.PreviewMouseDoubleClick
        If Search1.Text = "" Then Exit Sub
        Call SearchList1()
    End Sub
    Private Sub SearchList1() 'Handles Search1.PreviewMouseDoubleClick and Search1 <ENTER>
        'searches Listbox1 for entries with text in Search1, and scrolls to them in listbox1.

        Dim localWord As String
        Dim thisCount, localNum As Integer
        Dim lastObject As Object

        If UBound(Files1) < 1 Then Exit Sub
        'ListBox1.SelectedItem.Clear()
        lastObject = Nothing
        For thisCount = 0 To (ListBox1.Items.Count - 1)
            localWord = UCase(ListBox1.Items(thisCount).ToString)
            If Search1.Text <> "" Then
                If InStr(localWord, UCase(Search1.Text), CompareMethod.Text) <> 0 Then  'found TAG CODE match
                    localNum = thisCount
                    'ListBox1.SelectedItems.Add(ListBox1.Items(thisCount))
                    lastObject = ListBox1.Items(thisCount)
                End If
            End If
        Next thisCount
        ListBox1.ScrollIntoView(lastObject)
        ListBox1.SelectedIndex = localNum
    End Sub
    Private Sub Search2_Enter(sender As Object, e As KeyEventArgs) Handles Search2.PreviewKeyUp
        If e.Key = Key.Enter Then
            Call SearchList2()
        End If
    End Sub
    Private Sub Search2_DoubleClick() Handles Search2.PreviewMouseDoubleClick
        If Search2.Text = "" Then Exit Sub
        Call SearchList2()
    End Sub
    Private Sub SearchList2() 'Handles Search1.PreviewMouseDoubleClick and Search1 <ENTER>
        'searches Listbox1 for entries with text in Search1, and scrolls to them in listbox1.

        Dim localWord As String
        Dim thisCount, localNum As Integer
        Dim lastObject As Object

        If UBound(Files2) < 1 Then Exit Sub
        'ListBox1.SelectedItem.Clear()
        lastObject = Nothing
        For thisCount = 0 To (ListBox2.Items.Count - 1)
            localWord = UCase(ListBox2.Items(thisCount).ToString)
            If Search2.Text <> "" Then
                If InStr(localWord, UCase(Search2.Text), CompareMethod.Text) <> 0 Then  'found TAG CODE match
                    localNum = thisCount
                    'ListBox1.SelectedItems.Add(ListBox1.Items(thisCount))
                    lastObject = ListBox2.Items(thisCount)
                End If
            End If
        Next thisCount
        ListBox2.ScrollIntoView(lastObject)
        ListBox2.SelectedIndex = localNum
    End Sub

    Private Sub LoadList1()     'Reloads Listbox1 with all contents from Files1()
        Dim localNum As Integer
        Dim localName, localTime, localDay, localMonth, localYear As String
        Dim localDate As Date

        Search1.Text = ""
        If ListBox1.Items.Count > 0 Then
            ListBox1.Items.Clear()
        End If
        If UBound(Files1) > 0 Then
            For localNum = 1 To UBound(Files1)
                localName = Files1(localNum).FileName
                localName = Strings.Left(localName & "                              ", 25)
                localDate = Files1(localNum).FileDate
                'localTime = localDate.ToShortDateString

                localDay = localDate.Day.ToString("00")       'day of the month. Formatted to print day with leading zero if single digit
                localMonth = localDate.Month.ToString("00")     'month
                localYear = localDate.Year.ToString("0000")     'year
                localTime = localDay & "/" & localMonth & "/" & localYear
                ListBox1.Items.Add(localName & " " & localTime)
            Next
        End If
    End Sub
    Private Sub LoadList2()     'Reloads Listbox2 with all contents from Files2()
        Dim localNum As Integer
        Dim localName, localTime, localDay, localMonth, localYear As String
        Dim localDate As Date

        Search2.Text = ""
        If ListBox2.Items.Count > 0 Then
            ListBox2.Items.Clear()
        End If
        If UBound(Files2) > 0 Then
            For localNum = 1 To UBound(Files2)
                localName = Files2(localNum).FileName
                localName = Strings.Left(localName & "                              ", 25)
                localDate = Files2(localNum).FileDate
                'localTime = localDate.ToShortDateString

                localDay = localDate.Day.ToString("00")       'day of the month. Formatted to print day with leading zero if single digit
                localMonth = localDate.Month.ToString("00")     'month
                localYear = localDate.Year.ToString("0000")     'year
                localTime = localDay & "/" & localMonth & "/" & localYear
                ListBox2.Items.Add(localName & "  " & localTime)
            Next
        End If
    End Sub
    Private Sub SortByDate() Handles MnuSortOnDate.Click
        Search1.Text = ""
        Files2Sort = Files1
        Call SortFiles("DATE")
        Files1 = SortedFiles
        Call LoadList1()

        Search2.Text = ""
        Files2Sort = Files2
        Call SortFiles("DATE")
        Files2 = SortedFiles
        Call LoadList2()
    End Sub
    Private Sub SortByType() Handles MnuSortOnType.Click
        Search1.Text = ""
        Files2Sort = Files1
        Call SortFiles("TYPE")
        Files1 = SortedFiles
        Call LoadList1()

        Search2.Text = ""
        Files2Sort = Files2
        Call SortFiles("TYPE")
        Files2 = SortedFiles
        Call LoadList2()

    End Sub
    Private Sub SortFiles(tyPe As String)  'Sorts the files in Files2sort() and returns them in  SortedFiles.
        'type can be Date or Type.
        Dim localLoop, outerLoop, expanDloop As Integer
        Dim localFlag As Boolean
        Dim dAte1, dAte2 As Date
        Dim localLong As Long
        Dim exT1, exT2 As String


        SortedFiles = Files2Sort      'default
        If UCase(tyPe) <> "DATE" And UCase(tyPe) <> "TYPE" Then Exit Sub
        If UBound(Files2Sort) < 2 Then Exit Sub

        If UCase(tyPe) = "DATE" Then
            ReDim SortedFiles(1)
            SortedFiles(1) = Files2Sort(1)      'a starting point.
            For localLoop = 2 To UBound(Files2Sort)
                dAte2 = Files2Sort(localLoop).FileDate
                localFlag = False
                ReDim Preserve SortedFiles(UBound(SortedFiles) + 1)
                For outerLoop = 1 To UBound(SortedFiles)
                    dAte1 = SortedFiles(outerLoop).FileDate
                    localLong = DateDiff(DateInterval.Second, dAte2, dAte1)
                    If localLong <= 0 Then 'This file is newer.Insert it  here
                        localFlag = True
                        For expanDloop = UBound(SortedFiles) To outerLoop + 1 Step -1
                            SortedFiles(expanDloop) = SortedFiles(expanDloop - 1)
                        Next
                        SortedFiles(outerLoop) = Files2Sort(localLoop)
                        Exit For
                    End If
                Next outerLoop
                If Not localFlag Then
                    SortedFiles(UBound(SortedFiles)) = Files2Sort(localLoop)
                End If
            Next
        End If 'UCase(tyPe) = "DATE"
        If UCase(tyPe) = "TYPE" Then
            ReDim SortedFiles(0)
            For localLoop = 1 To UBound(Files2Sort)
                exT1 = UCase(Files2Sort(localLoop).Extension)
                If exT1 <> "" Then
                    For outerLoop = localLoop To UBound(Files2Sort)
                        exT2 = UCase(Files2Sort(outerLoop).Extension)
                        If exT1 = exT2 Then
                            ReDim Preserve SortedFiles(UBound(SortedFiles) + 1)
                            SortedFiles(UBound(SortedFiles)) = Files2Sort(outerLoop)
                            Files2Sort(outerLoop).Extension = ""
                        End If
                    Next
                End If
            Next
        End If

    End Sub
    Private Sub ReadHistory()
        Dim localFlag As Boolean
        Dim fileNum As Integer
        Dim localLine As String

        'LOAD DEFAULTS.
        Call FindFileTypes()  'loads defaults for ClickedFileTypes and ChosenFileTypes
        Pointer1 = -1
        Pointer2 = -1
        localFlag = My.Computer.FileSystem.FileExists(PARENTPATH & "\" & OptionsFileName)
        If localFlag Then
            ReDim SavedHistory(0)
            fileNum = FreeFile()  'any number
            FileOpen(fileNum, PARENTPATH & "\" & OptionsFileName, OpenMode.Input)
            Do Until EOF(fileNum)
                localLine = LineInput(fileNum)
                ReDim Preserve SavedHistory(UBound(SavedHistory) + 1)
                SavedHistory(UBound(SavedHistory)) = localLine
            Loop
            FileClose((fileNum))
            For localloop = 1 To UBound(SavedHistory)
                localLine = SavedHistory(localloop)
                If InStr(localLine, "HoldPath1=") <> 0 Then HoldPath1 = Strings.Right(localLine, Len(localLine) - 10)
                If InStr(localLine, "HoldPath2=") <> 0 Then HoldPath2 = Strings.Right(localLine, Len(localLine) - 10)
                If InStr(localLine, "ClickedFileTypes=") <> 0 Then ClickedFileTypes = Strings.Right(localLine, Len(localLine) - 17)
                If InStr(localLine, "ChosenFileTypes=") <> 0 Then ChosenFileTypes = Strings.Right(localLine, Len(localLine) - 16)
                If InStr(localLine, "SOURCEPATH1=") <> 0 Then SOURCEPATH1 = Strings.Right(localLine, Len(localLine) - 12)
                If InStr(localLine, "SOURCEPATH2=") <> 0 Then SOURCEPATH2 = Strings.Right(localLine, Len(localLine) - 12)
                If InStr(localLine, "INDEX1=") <> 0 Then Pointer1 = CInt(Val(Strings.Right(localLine, Len(localLine) - 7)))
                If InStr(localLine, "INDEX2=") <> 0 Then Pointer2 = CInt(Val(Strings.Right(localLine, Len(localLine) - 7)))
                If InStr(localLine, "OPTIONSFILENAME=") <> 0 Then OptionsFileName = Strings.Right(localLine, Len(localLine) - 16)
            Next
        End If
        OptionFile.Content = OptionsFileName
        'Process ClickedFileTypes. ,txt,mcs,h, 
        If InStr(ClickedFileTypes, ",txt,") <> 0 Then Check1.IsChecked = True
        If InStr(ClickedFileTypes, ",vb,") <> 0 Then Check2.IsChecked = True
        If InStr(ClickedFileTypes, ",xaml,") <> 0 Then Check3.IsChecked = True
        If InStr(ClickedFileTypes, ",ach,") <> 0 Then Check4.IsChecked = True

        If InStr(ClickedFileTypes, ",asm,") <> 0 Then Check5.IsChecked = True
        If InStr(ClickedFileTypes, ",inc,") <> 0 Then Check6.IsChecked = True
        If InStr(ClickedFileTypes, ",mcs,") <> 0 Then Check7.IsChecked = True
        If InStr(ClickedFileTypes, ",mcp,") <> 0 Then Check8.IsChecked = True

        If InStr(ClickedFileTypes, ",ino,") <> 0 Then Check9.IsChecked = True
        If InStr(ClickedFileTypes, ",raw,") <> 0 Then Check10.IsChecked = True
        If InStr(ClickedFileTypes, ",mcs,") <> 0 Then Check11.IsChecked = True
        If InStr(ClickedFileTypes, ",mcp,") <> 0 Then Check12.IsChecked = True
    End Sub
    Private Sub SaveHistory()
        Dim fileNum, localLoop As Integer
        Dim localLine As String

        Call LoadSaveHistory()
        fileNum = FreeFile()
        FileOpen(fileNum, PARENTPATH & "\" & "CompareOptions.txt", OpenMode.Output)
        For localLoop = 1 To UBound(SavedHistory)
            localLine = SavedHistory(localLoop)
            Print(fileNum, localLine & vbCr & vbLf)
        Next localLoop
        FileClose((fileNum))

    End Sub
    Private Sub LoadSaveHistory()
        Call FindFileTypes()    'loads CLickedFileTypes and ChosenFileTypes. ChosenFileTypes = clicked or ",*,"
        ReDim SavedHistory(11)
        SavedHistory(1) = "PARENTPATH=" & PARENTPATH
        SavedHistory(2) = "HoldPath1=" & HoldPath1
        SavedHistory(3) = "HoldPath2=" & HoldPath2
        SavedHistory(4) = "ClickedFileTypes=" & ClickedFileTypes
        SavedHistory(5) = "ChosenFileTypes=" & ChosenFileTypes
        SavedHistory(6) = "SOURCEPATH1=" & SOURCEPATH1
        SavedHistory(7) = "SOURCEPATH2=" & SOURCEPATH2
        SavedHistory(8) = "INDEX1=" & ListBox1.SelectedIndex.ToString
        SavedHistory(9) = "INDEX2=" & ListBox2.SelectedIndex.ToString
        SavedHistory(10) = "OPTIONSFILENAME=" & OptionsFileName
        SavedHistory(11) = "END:="
    End Sub
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Call SaveHistory()
    End Sub
    Private Sub CLOSEclicked() Handles MnuCLOSE.Click
        Me.Close()
    End Sub
    Private Sub OpenEvent() Handles MnuOpenEvent.Click
        Process.Start("explorer.exe", "/root," & PARENTPATH)  'Opens at PARENTPATH
    End Sub
    Private Sub MnuEventPath_Click(sender As Object, e As EventArgs) Handles MnuEventPath.Click
        'MESSAGEACTIVE = True
        reply = MsgBox("Your event Folder is  " & vbCrLf &
        PARENTPATH & vbCrLf &
        "", vbOKOnly, "EVENT PATH.")
    End Sub




    'BYTE0 = EVENTTIME.Day.ToString("00")       'day of the month. Formatted to print day with leading zero if single digit
    'BYTE1 = EVENTTIME.Month.ToString("00")     'month
    'BYTE2 = EVENTTIME.Year.ToString("0000")     'month

End Class
