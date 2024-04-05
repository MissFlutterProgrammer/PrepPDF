Imports System.Collections.Specialized
Imports System.Globalization
Imports System.IO
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas
'Imports iText.Kernel.Pdf.Canvas.Parser
'Imports iText.Kernel.Pdf.Canvas.Parser.Filter
'Imports iText.Kernel.Pdf.Canvas.Parser.Listener
Imports iText.Kernel.Pdf.Xobject
Imports iText.Kernel.Utils

Public Class Form1
    Public sw As IO.StreamWriter
    Dim LeftPageNum As String = "0"
    Dim RightPageNum As String = "0"
    Dim ItsABatchRun As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.ComboBoxTitle.DisplayMember = "Key"
        Me.ComboBoxTitle.ValueMember = "Text"
        Me.ComboBoxTitle.Items.Add(New DictionaryEntry("Daily Mail", "DailyMail"))
        Me.ComboBoxTitle.Items.Add(New DictionaryEntry("The Mail On Sunday", "TMOS"))
        ComboBoxTitle.SelectedIndex = 0
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -7, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -6, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -5, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -4, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -3, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -2, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, -1, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateTime.Now.ToString("dd-MM-yyyy"))
        ComboBoxIssue.Items.Add(DateAdd(DateInterval.Day, 1, DateTime.Now).ToString("dd-MM-yyyy"))
        ComboBoxIssue.SelectedIndex = 8

        ComboBoxEdition.Items.Add(1)
        ComboBoxEdition.Items.Add(2)
        ComboBoxEdition.SelectedIndex = 0

        ComboBoxRegion.DisplayMember = "Key"
        ComboBoxRegion.ValueMember = "Text"
        ComboBoxRegion.Items.Add(New DictionaryEntry("Ireland", "IRE (IDM)"))
        ComboBoxRegion.Items.Add(New DictionaryEntry("National", "NAT "))
        ComboBoxRegion.Items.Add(New DictionaryEntry("Scotland", "SCT (SCT)"))
        ComboBoxRegion.SelectedIndex = 1
        Application.DoEvents()
        Dim CommandArguments As String = Environment.CommandLine()
        Dim ArgumentCount As Integer = 0

        'first argument is application.exe name
        If Mid(CommandArguments, 1, (Application.ProductName & ".exe").Length()).ToLower() = Application.ProductName.ToLower() & ".exe" Then
            Dim Position As Integer
            Position = InStr(CommandArguments.ToLower(), "title=", CompareMethod.Text)
            If Position > 0 Then
                Position = Position + 6
                'got start point of title, read upto the next space or &
                txtTitle.Text = ""
                ArgumentCount = ArgumentCount + 1
                Do While Mid(CommandArguments, Position, 1) <> " " And Position <= CommandArguments.Length()
                    txtTitle.Text = txtTitle.Text + Mid(CommandArguments, Position, 1)
                    Position = Position + 1
                Loop
            End If
            Application.DoEvents()
            Position = 0
            Position = InStr(CommandArguments.ToLower(), "issue=", CompareMethod.Text)
            If Position > 0 Then
                Position = Position + 6
                'got start point of issue, read upto the next space or &
                txtIssue.Text = ""
                ArgumentCount = ArgumentCount + 1
                Do While Mid(CommandArguments, Position, 1) <> " " And Position <= CommandArguments.Length()
                    txtIssue.Text = txtIssue.Text + Mid(CommandArguments, Position, 1)
                    Position = Position + 1
                Loop
            End If
            Application.DoEvents()
            Position = 0
            Position = InStr(CommandArguments.ToLower(), "edition=", CompareMethod.Text)
            If Position > 0 Then
                Position = Position + 8
                'got start point of edition, read upto the next space or &
                txtEdition.Text = ""
                ArgumentCount = ArgumentCount + 1
                Do While Mid(CommandArguments, Position, 1) <> " " And Position <= CommandArguments.Length()
                    txtEdition.Text = txtEdition.Text + Mid(CommandArguments, Position, 1)
                    Position = Position + 1
                Loop
            End If
            Application.DoEvents()
            Position = 0
            Position = InStr(CommandArguments.ToLower(), "region=", CompareMethod.Text)
            If Position > 0 Then
                Position = Position + 7
                'got start point of edition, read upto the next space or &
                txtRegion.Text = ""
                ArgumentCount = ArgumentCount + 1
                Do While Mid(CommandArguments, Position, 1) <> " " And Position <= CommandArguments.Length()
                    txtRegion.Text = txtRegion.Text + Mid(CommandArguments, Position, 1)
                    Position = Position + 1
                Loop
            End If
            Application.DoEvents()
        End If

        'validate the parameters we've got
        Dim ParametersOK As Boolean = True
        If ArgumentCount > 0 Then

            'check whether txtTitle.text is in comboboxtitle
            Dim index As Integer
            Dim KeyValue As DictionaryEntry
            For index = 0 To ComboBoxTitle.Items.Count - 1
                KeyValue = CType(ComboBoxTitle.Items(index), DictionaryEntry)
                If txtTitle.Text.ToLower() = KeyValue.Value.ToString().ToLower() Then
                    ComboBoxTitle.SelectedIndex = index
                    'ParametersOK = True
                    lblCommand.Text = "Title OK"

                    Exit For
                Else
                    ParametersOK = False
                    lblCommand.Text = "Title Invalid"
                End If
            Next
            Application.DoEvents()
            If txtIssue.Text.ToLower() <> "all" Then
                Dim formats() As String = {"dd-MM-yy", "dd-MM-yyyy", "ddMMyyyy", "dd/MM/yy", "dd/MM/yyyy", "dd\MM\yy", "dd\MM\yyyy"}
                Dim thisDt As DateTime
                ' this should work with all format strings above
                If DateTime.TryParseExact(txtIssue.Text, formats, Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, thisDt) Then
                    txtIssue.Text = thisDt.ToString("dd-MM-yyyy")
                    index = -1
                    index = ComboBoxIssue.FindString(txtIssue.Text)
                    If index <> -1 Then
                        ComboBoxIssue.SelectedIndex = index
                        ' ParametersOK = True
                        lblCommand.Text = lblCommand.Text + vbLf + "Issue OK"
                    Else
                        ParametersOK = False
                        lblCommand.Text = lblCommand.Text + vbLf + "Issue Invalid"
                    End If
                Else
                    ParametersOK = False
                    lblCommand.Text = lblCommand.Text + vbLf + "Issue Invalid"
                End If
                Application.DoEvents()
                'If ParametersOK Then
                index = -1
                index = ComboBoxEdition.FindString(txtEdition.Text)
                If index <> -1 Then
                    ComboBoxEdition.SelectedIndex = index
                    'ParametersOK = True
                    lblCommand.Text = lblCommand.Text + vbLf + "Edition OK"
                Else
                    ParametersOK = False
                    lblCommand.Text = lblCommand.Text + vbLf + "Edition Invalid"
                End If
                'End If
                Application.DoEvents()
                'If ParametersOK Then

                For index = 0 To ComboBoxRegion.Items.Count - 1
                    KeyValue = CType(ComboBoxRegion.Items(index), DictionaryEntry)
                    'MsgBox(KeyValue.Key.ToString().ToLower())
                    If txtRegion.Text.ToLower() = KeyValue.Key.ToString().ToLower() Then
                        ComboBoxRegion.SelectedIndex = index
                        'ParametersOK = True
                        lblCommand.Text = lblCommand.Text + vbLf + "Region OK"
                        Exit For
                    Else
                        If index = ComboBoxRegion.Items.Count - 1 Then
                            ParametersOK = False
                            lblCommand.Text = lblCommand.Text + vbLf + "Region Invalid"
                        End If
                    End If

                Next
                'End If
                Application.DoEvents()
            Else
                'all found on issue so batch run Today+1day for all txtTitle.text
                'ItsABatchRun = True
                'lblCommand.Text = lblCommand.Text + vbLf + "Starting Batch Run"
                'Me.Height = 217
                'Me.Width = 788
                'Button1.Location = New Point(687, 140)
                Application.DoEvents()
                Me.Button10.PerformClick()
                'lblCommand.Text = lblCommand.Text + vbLf + "Finished Batch Run"
                Me.Close()
            End If


        End If
        'MsgBox(lblCommand.Text)
        If ParametersOK And ArgumentCount = 4 Then
            ItsABatchRun = True
            'lblCommand.Text = lblCommand.Text + vbLf + "Starting Batch Run"
            'Me.Height = 217
            'Me.Width = 788
            'Button1.Location = New Point(687, 140)
            Application.DoEvents()
            Call BatchRun_V2()
            'lblCommand.Text = lblCommand.Text + vbLf + "Finished Batch Run"
            Me.Close()
        Else
            If ArgumentCount > 4 Then
                'lblCommand.Text = lblCommand.Text + vbLf + "Invalid Parameters Can't Batch Run"
                Application.DoEvents()
            End If
        End If

    End Sub

    Public Sub WriteLogFile(ByVal LogLine As String)

        Dim FilePath As String
        If ItsABatchRun Then
            FilePath = "D:\PrepPDF\" & txtTitle.Text & "-" & Replace(txtIssue.Text, "-", "") & "-" & txtEdition.Text & "-" & txtRegion.Text & ".log"
        Else
            FilePath = "D:\PrepPDF\PrepPDF.log"
        End If

        If Not File.Exists(FilePath) Then
            sw = File.CreateText(FilePath)
            sw.WriteLine(DateTime.Now.ToString("yyyyMMdd  HH:mm:ss") & " - Log File Created.")
        Else
            sw = File.AppendText(FilePath)
        End If
        sw.WriteLine(DateTime.Now.ToString("yyyyMMdd  HH:mm:ss") & ": " & LogLine)
        sw.Close()


    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False

        Dim TimeStart As DateTime
        TimeStart = Now()

        Dim success As Boolean = False
        Dim IgnoredPages As Int16 = 0
        Dim NewPages As Int16 = 0
        Dim ModifiedPages As Int16 = 0

        Dim filepath As String = ""
        Dim fName As String = ""
        Dim SourceDir As DirectoryInfo
        Dim CurrentPage As Int16 = 0
        Dim input_files_list As New List(Of String)
        Dim FullBookFolder As String


        Using sw As New StreamWriter(File.Open("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxTitle.SelectedItem.Value & ".log", FileMode.OpenOrCreate))
            For Edition = 1 To 4

                For Each PubRegion In {"Ireland", "National", "Scotland"}
                    FullBookFolder = "FullBook"

                    If Directory.Exists("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion & "\" & FullBookFolder) Then
                        input_files_list = New List(Of String)
                        SourceDir = New DirectoryInfo("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion & "\" & FullBookFolder)
                        For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf").OrderBy(Function(file) file.FullName)
                            sw.WriteLine(childFile.FullName)

                            input_files_list.Add(childFile.FullName)

                        Next
                        success = Merge_PDF_Files(input_files_list, "D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & ComboBoxTitle.SelectedItem.Value & "_" & ComboBoxIssue.Text & "_" & CStr(Edition) & "_" & PubRegion & ".pdf")
                        'Optional: handling errors'   
                        If success Then
                            sw.WriteLine("Files merged")
                        Else
                            sw.WriteLine("Error merging files")
                        End If

                    End If
                    CurrentPage = 0
                Next

            Next

            sw.Close()

            Me.Enabled = True
            Me.Cursor = Cursors.Default



            '  lblTime.Text = "Started: " & TimeStart.ToString("HH:mm:ss") & vbLf & "Finished: " & Now.ToString("HH:mm:ss")
            '  lblWhat.Text = "New:  " & CStr(NewPages) & vbLf & "Modified: " & CStr(ModifiedPages) & vbLf & "Ignored: " & CStr(IgnoredPages)





        End Using
    End Sub

    'Public Sub FindPageNumbers(ThisPDFPage As PdfPage)
    '    Dim toFindPageNumbers As iText.Kernel.Geom.Rectangle
    '    toFindPageNumbers = New iText.Kernel.Geom.Rectangle(ThisPDFPage.GetPageSize())
    '    toFindPageNumbers.SetY(ThisPDFPage.GetPageSize().GetHeight() - 100)
    '    Dim regionFilter As TextRegionEventFilter = New TextRegionEventFilter(toFindPageNumbers)
    '    Dim strategy As ITextExtractionStrategy = New FilteredTextEventListener(New LocationTextExtractionStrategy(), regionFilter)
    '    Dim strText As String = PdfTextExtractor.GetTextFromPage(ThisPDFPage, strategy)


    '    If strText.Length > 10 Then

    '        Dim strpos As Integer

    '        Dim numcount As Integer = 0
    '        Dim charcount As Integer = 0
    '        strpos = InStr(strText.ToLower(), "page", CompareMethod.Text)
    '        strpos = strpos + 4
    '        Do While strpos <= strText.Length
    '            If IsNumeric(Mid(strText, strpos, 1)) Then
    '                LeftPageNum = LeftPageNum & Mid(strText, strpos, 1)
    '                numcount = numcount + 1
    '            Else
    '                If numcount <> 0 Then
    '                    Exit Do
    '                End If
    '                charcount = charcount + 1
    '                If charcount = 4 Then
    '                    Exit Do
    '                End If
    '            End If
    '            strpos = strpos + 1
    '        Loop
    '        strpos = InStr(strpos, strText.ToLower(), "page", CompareMethod.Text)
    '        strpos = strpos + 4
    '        numcount = 0
    '        charcount = 0
    '        Do While strpos <= strText.Length
    '            If IsNumeric(Mid(strText, strpos, 1)) Then
    '                RightPageNum = RightPageNum & Mid(strText, strpos, 1)
    '                numcount = numcount + 1
    '            Else
    '                If numcount <> 0 Then
    '                    Exit Do
    '                End If
    '                charcount = charcount + 1
    '                If charcount = 4 Then
    '                    Exit Do
    '                End If
    '            End If
    '            strpos = strpos + 1
    '        Loop
    '    End If
    'End Sub

    Public Function getHalfPageSize(pagesize As iText.Kernel.Geom.Rectangle) As iText.Kernel.Geom.Rectangle

        Dim Width As Single
        Dim Height As Single

        Width = pagesize.GetWidth()
        Height = pagesize.GetHeight()
        Return New iText.Kernel.Geom.Rectangle(Width / 2, Height)

    End Function
    Public Function Merge_PDF_Files(ByVal input_files As List(Of String), ByVal output_file As String) As Boolean
        Dim Input_Document As PdfDocument = Nothing
        Dim Output_Document As PdfDocument = Nothing
        Dim Merger As PdfMerger

        Try
            Output_Document = New iText.Kernel.Pdf.PdfDocument(New iText.Kernel.Pdf.PdfWriter(output_file))
            'Create the output file (Document) from a Merger stream'
            Merger = New PdfMerger(Output_Document)

            'Merge each input PDF file to the output document'
            For Each file As String In input_files
                Input_Document = New PdfDocument(New PdfReader(file))
                Merger.Merge(Input_Document, 1, Input_Document.GetNumberOfPages())
                Input_Document.Close()
            Next

            Output_Document.Close()
            Return True
        Catch ex As Exception
            'catch Exception if needed'
            If Input_Document IsNot Nothing Then Input_Document.Close()
            If Output_Document IsNot Nothing Then Output_Document.Close()
            File.Delete(output_file)

            Return False
        End Try
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Sub BatchRun_V2()

        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        Try
            Dim PageFolder As String = "NAT"
            Dim FullBookFolder As String
            Dim ThisPDFDocument As PdfDocument
            Dim ThisPDFPage As PdfPage
            Dim oNewPage As PdfPage
            Dim ThisPageSize As iText.Kernel.Geom.Rectangle
            Dim NumOfPDFPages As Integer
            Dim PhysicalPageNum As Integer
            Dim SourceDir As DirectoryInfo
            Dim NewFileName As String
            Dim EvenPage As String
            Dim OddPage As String
            Dim DateWritten As DateTime
            Dim DateCreated As DateTime
            Dim CompareWritten As DateTime
            Dim PageInFolder() As String
            Dim PagesToProcess As Integer
            Dim P As Integer
            Dim LastFileWritten As String = "0000000"
            Dim LastFileRead As String
            Dim InsertPage As String = ""
            Dim InsertPageNumber As Integer
            Dim FirstInsertPage As Boolean = False
            Dim InsertPageCount As Integer = 0

            Dim Edition As String
            Edition = txtEdition.Text
            'Edition = "1"

            Dim PubRegion As String
            PubRegion = txtRegion.Text
            'PubRegion = "Ireland"

            If Directory.Exists("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion) Then
                If PubRegion = "Ireland" Then
                    PageFolder = "IRE (IDM)"
                End If
                If PubRegion = "National" Then
                    PageFolder = "NAT "
                End If
                If PubRegion = "Scotland" Then
                    PageFolder = "SCT (SCT)"
                End If
                FullBookFolder = "D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion & "\" & "FullBook"
                If Not Directory.Exists(FullBookFolder) Then
                    Directory.CreateDirectory(FullBookFolder)
                End If

                For Each LocalDir In My.Computer.FileSystem.GetDirectories("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion, FileIO.SearchOption.SearchTopLevelOnly, "*")
                    If LocalDir.ToLower.EndsWith("fullbook") Then
                        Continue For
                    End If
                    SourceDir = New DirectoryInfo(LocalDir)
                    'If SourceDir.Name.ToLower.StartsWith("in") Then
                    ' Continue For
                    'End If

                    PagesToProcess = SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).Count
                    ReDim PageInFolder(PagesToProcess - 1)
                    P = -1
                    For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Length).ThenBy(Function(fi) fi.Name)
                        P = P + 1
                        PageInFolder(P) = childFile.Name
                        If Not SourceDir.Name.ToLower.StartsWith("in") Then
                            LastFileRead = childFile.Name
                        End If
                    Next

                    For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Length).ThenBy(Function(fi) fi.Name)

                        InsertPage = ""
                        If Mid(SourceDir.Name, 1, 2) = "IN" Then
                            InsertPage = "IN"
                            If Not FirstInsertPage Then
                                InsertPageNumber = CInt(Mid(LastFileRead, Len(LastFileRead) - 6, 3))
                                FirstInsertPage = True
                            End If
                        End If
                        NewFileName = SourceDir.FullName & "\" & InsertPage & Mid(childFile.Name, Len(childFile.Name) - 6, 3) & " " & Edition & " " & PageFolder & " " & Mid(Replace(ComboBoxIssue.Text, "-", ""), 1, 4) & ".indd"
                        If File.Exists(NewFileName) Then   'Removing this line prevent the indd exists check therefore now a strict newest file wins senario

                            'matched .pdf to .indd, so got first file in folder
                            'if remaining files in folder are same sizes, then each filename has page number
                            'first file will have as many pages as files in folder
                            'if not matching file size treat as a single or spread page evens left odds right, single page either applies

                            'make sure starting P matches the index of the childFile found
                            For P = 0 To PageInFolder.Count - 1
                                If PageInFolder(P) = childFile.Name Then
                                    Exit For
                                End If
                            Next

                            DateCreated = childFile.CreationTime
                            DateCreated = New DateTime(DateCreated.Year, DateCreated.Month, DateCreated.Day, DateCreated.Hour, DateCreated.Minute, DateCreated.Second)
                            DateWritten = childFile.LastWriteTime
                            DateWritten = New DateTime(DateWritten.Year, DateWritten.Month, DateWritten.Day, DateWritten.Hour, DateWritten.Minute, DateWritten.Second)
                            'CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & Replace(childFile.Name.ToLower(), ".pdf", ".bak"))
                            If InsertPage = "IN" Then
                                InsertPageNumber = InsertPageNumber + 1
                                CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF")
                            Else
                                CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & childFile.Name.ToLower())
                            End If

                            CompareWritten = New DateTime(CompareWritten.Year, CompareWritten.Month, CompareWritten.Day, CompareWritten.Hour, CompareWritten.Minute, CompareWritten.Second)

                            'is incoming file newer than previously processed
                            If (DateTime.Compare(DateWritten, CompareWritten) > 0) And (DateDiff(DateInterval.Second, CompareWritten, DateWritten) > 25) Then

                                ThisPDFDocument = New PdfDocument(New PdfReader(childFile))
                                NumOfPDFPages = ThisPDFDocument.GetNumberOfPages()
                                'get physical page number from file name
                                PhysicalPageNum = CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3))

                                For x = 1 To NumOfPDFPages
                                    ThisPDFPage = ThisPDFDocument.GetPage(x)
                                    ThisPageSize = ThisPDFPage.GetPageSize()
                                    'WriteLogFile(PubRegion & "|" & Edition & "|" & childFile.Name & "|" & CStr(NumOfPDFPages) & "|" & CStr(PhysicalPageNum) & "|" & childFile.Length.ToString() & "|" & childFile.LastWriteTime.ToString("dd-MM-yyyy HH:mm:ss"))

                                    'is it already a single page 827.7w x 1105.5h Points
                                    If ThisPageSize.GetHeight() < 1120 And ThisPageSize.GetWidth() < 840 Then
                                        'its a single just copy accross
                                        If NumOfPDFPages = 1 Then
                                            If InsertPage = "IN" Then
                                                'InsertPageNumber = InsertPageNumber + 1
                                                NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                            Else

                                                NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                            End If
                                            childFile.CopyTo(NewFileName, True)
                                            File.SetCreationTime(NewFileName, DateCreated)
                                            File.SetLastWriteTime(NewFileName, DateWritten)
                                            LastFileWritten = NewFileName
                                        Else
                                            'Write 2nd and subsequent ThisPDFPage to new file name
                                            If InsertPage = "IN" Then
                                                InsertPageNumber = InsertPageNumber + 1
                                                NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                            Else

                                                NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                            End If

                                            Using pdfNew = New PdfDocument(New PdfWriter(NewFileName))
                                                oNewPage = ThisPDFPage.CopyTo(pdfNew)
                                                pdfNew.AddPage(oNewPage)
                                                pdfNew.Close()
                                            End Using
                                            File.SetCreationTime(NewFileName, DateCreated)
                                            File.SetLastWriteTime(NewFileName, DateWritten)
                                            LastFileWritten = NewFileName

                                        End If

                                        P = P + 1
                                    Else
                                        If ThisPageSize.GetHeight() < ((1105.5 * 2) - 15) And ThisPageSize.GetWidth() > 840 Then
                                            'Its a spread even page take the left odd page take the right
                                            If InsertPage = "IN" Then
                                                ' InsertPageNumber = InsertPageNumber + 1
                                                EvenPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                            Else
                                                EvenPage = FullBookFolder & "\" & PageInFolder(P)
                                            End If

                                            P = P + 1
                                            If P > (PageInFolder.Count - 1) Then
                                                'If Theres no other file in array take page number from last page and increase by 1
                                                P = P - 1
                                                If InsertPage = "IN" Then
                                                    'InsertPageNumber = InsertPageNumber + 1
                                                    OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                Else
                                                    OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3) + 1), "000") & ".PDF"
                                                End If
                                            Else
                                                If InsertPage = "IN" Then
                                                    InsertPageNumber = InsertPageNumber + 1
                                                    OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                Else
                                                    OddPage = FullBookFolder & "\" & PageInFolder(P)
                                                End If
                                            End If


                                            Dim toMove As iText.Kernel.Geom.Rectangle
                                            toMove = New iText.Kernel.Geom.Rectangle(getHalfPageSize(ThisPDFPage.GetPageSize()))

                                            Dim pdfDoc As New PdfDocument(New PdfWriter(EvenPage, New WriterProperties().SetFullCompressionMode(True)))
                                            pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                            pdfDoc.AddNewPage()

                                            'Left Side
                                            Dim t_canvas1 As New PdfFormXObject(toMove)
                                            Dim canvas1 As New PdfCanvas(t_canvas1, pdfDoc)
                                            Dim pageXObject As PdfFormXObject
                                            pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                            canvas1.Rectangle(0, 0, toMove.GetWidth(), toMove.GetHeight())
                                            canvas1.EoClip()
                                            canvas1.EndPath()
                                            canvas1.AddXObjectAt(pageXObject, 0, 0)
                                            Dim canvas As New PdfCanvas(pdfDoc.GetFirstPage())
                                            canvas.AddXObjectAt(t_canvas1, 0, 0)

                                            pdfDoc.Close()
                                            File.SetCreationTime(EvenPage, DateCreated)
                                            File.SetLastWriteTime(EvenPage, DateWritten)
                                            'Right Side
                                            pdfDoc = New PdfDocument(New PdfWriter(OddPage, New WriterProperties().SetFullCompressionMode(True)))
                                            pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                            pdfDoc.AddNewPage()

                                            toMove.SetX(ThisPageSize.GetWidth() - toMove.GetWidth())

                                            Dim t_canvas2 As New PdfFormXObject(toMove)
                                            Dim canvas2 As New PdfCanvas(t_canvas2, pdfDoc)
                                            pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                            canvas2.Rectangle(toMove.GetLeft(), toMove.GetBottom(), toMove.GetWidth(), toMove.GetHeight())
                                            canvas2.Clip()
                                            canvas2.EndPath()
                                            canvas2.AddXObjectAt(pageXObject, 0, 0)

                                            canvas = New PdfCanvas(pdfDoc.GetFirstPage())
                                            canvas.AddXObjectAt(t_canvas2, 0, 0)

                                            pdfDoc.Close()
                                            File.SetCreationTime(OddPage, DateCreated)
                                            File.SetLastWriteTime(OddPage, DateWritten)
                                            LastFileWritten = OddPage
                                            P = P + 1
                                        Else
                                            'WTF have we got here
                                        End If
                                    End If
                                Next
                            Else
                                If InsertPage = "IN" Then
                                    InsertPageCount = 0
                                    InsertPageCount = GetNumberOfSinglePages(childFile)
                                    If InsertPageCount >= 2 Then
                                        InsertPageNumber = InsertPageNumber + (InsertPageCount - 1)
                                    End If


                                End If
                            End If
                        End If   'This matches the indd file exists check
                    Next
                Next
                Call FillInTheBlanks(FullBookFolder, ComboBoxTitle.SelectedItem.Value, ComboBoxIssue.Text, LastFileWritten)
            End If
        Catch ex As System.IO.FileNotFoundException
            WriteLogFile("Error: " & ex.Message)
        Catch ex As Exception
            WriteLogFile("Error: " & ex.Message)
        End Try
        Me.Cursor = Cursors.Default
        Me.Enabled = True
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click

        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        Try
            Dim PageFolder As String = "NAT"
            Dim FullBookFolder As String
            Dim ThisPDFDocument As PdfDocument
            Dim ThisPDFPage As PdfPage
            Dim oNewPage As PdfPage
            Dim ThisPageSize As iText.Kernel.Geom.Rectangle
            Dim NumOfPDFPages As Integer
            Dim PhysicalPageNum As Integer
            Dim SourceDir As DirectoryInfo
            Dim NewFileName As String

            Dim EvenPage As String
            Dim OddPage As String
            Dim DateWritten As DateTime
            Dim DateCreated As DateTime
            Dim CompareWritten As DateTime
            Dim PageInFolder() As String
            Dim PagesToProcess As Integer
            Dim P As Integer
            Dim LastFileWritten As String = "0000000"
            Dim LastFileRead As String
            Dim InsertPage As String = ""
            Dim InsertPageNumber As Integer
            Dim FirstInsertPage As Boolean = False
            Dim InsertPageCount As Integer = 0

            ' Dim Edition As String
            ' Dim PubRegion As String

            'Edition = ComboBoxEdition.Text
            'PubRegion = ComboBoxRegion.Text


            For Each Edition In {"1", "2"}
                For Each PubRegion In {"Ireland", "National", "Scotland"}
                    FirstInsertPage = False
                    If Directory.Exists("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion) Then
                        'WriteLogFile("Title: " + ComboBoxTitle.SelectedItem.Value + "  Issue: " + ComboBoxIssue.Text + "  Edition: " + Edition.ToString() + "  Region: " + PubRegion)
                        If PubRegion = "Ireland" Then
                            PageFolder = "IRE (IDM)"
                        End If
                        If PubRegion = "National" Then
                            PageFolder = "NAT "
                        End If
                        If PubRegion = "Scotland" Then
                            PageFolder = "SCT (SCT)"
                        End If
                        CheckForDuplicateInsertPDF(Edition, PubRegion)
                        FullBookFolder = "D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion & "\" & "FullBook"
                        If Not Directory.Exists(FullBookFolder) Then
                            Directory.CreateDirectory(FullBookFolder)
                        End If

                        For Each LocalDir In My.Computer.FileSystem.GetDirectories("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion, FileIO.SearchOption.SearchTopLevelOnly, "*")
                            If LocalDir.ToLower.EndsWith("fullbook") Then
                                Continue For
                            End If
                            SourceDir = New DirectoryInfo(LocalDir)
                            'If SourceDir.Name.ToLower.StartsWith("in") Then
                            ' Continue For
                            'End If

                            PagesToProcess = SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).Count
                            ReDim PageInFolder(PagesToProcess - 1)
                            P = -1
                            For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Length).ThenBy(Function(fi) fi.Name)
                                P = P + 1
                                PageInFolder(P) = childFile.Name
                                If Not SourceDir.Name.ToLower.StartsWith("in") Then
                                    LastFileRead = childFile.Name
                                End If
                            Next
                            'LastFileRead = childFile.Name

                            For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Length).ThenBy(Function(fi) fi.Name)

                                InsertPage = ""
                                If Mid(SourceDir.Name, 1, 2) = "IN" Then
                                    InsertPage = "IN"
                                    If Not FirstInsertPage Then
                                        InsertPageNumber = CInt(Mid(LastFileRead, Len(LastFileRead) - 6, 3))
                                        FirstInsertPage = True
                                    End If
                                End If
                                NewFileName = SourceDir.FullName & "\" & InsertPage & Mid(childFile.Name, Len(childFile.Name) - 6, 3) & " " & Edition & " " & PageFolder & " " & Mid(Replace(ComboBoxIssue.Text, "-", ""), 1, 4) & ".indd"
                                If File.Exists(NewFileName) Then   'Removing this line prevent the indd exists check therefore now a strict newest file wins senario

                                    'matched .pdf to .indd, so got first file in folder
                                    'if remaining files in folder are same sizes, then each filename has page number
                                    'first file will have as many pages as files in folder
                                    'if not matching file size treat as a single or spread page evens left odds right, single page either applies

                                    'make sure starting P matches the index of the childFile found
                                    For P = 0 To PageInFolder.Count - 1
                                        If PageInFolder(P) = childFile.Name Then
                                            Exit For
                                        End If
                                    Next

                                    DateCreated = childFile.CreationTime
                                    DateCreated = New DateTime(DateCreated.Year, DateCreated.Month, DateCreated.Day, DateCreated.Hour, DateCreated.Minute, DateCreated.Second)
                                    DateWritten = childFile.LastWriteTime
                                    DateWritten = New DateTime(DateWritten.Year, DateWritten.Month, DateWritten.Day, DateWritten.Hour, DateWritten.Minute, DateWritten.Second)
                                    'CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & Replace(childFile.Name.ToLower(), ".pdf", ".bak"))
                                    If InsertPage = "IN" Then
                                        InsertPageNumber = InsertPageNumber + 1
                                        CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF")
                                    Else
                                        CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & childFile.Name.ToLower())
                                    End If

                                    CompareWritten = New DateTime(CompareWritten.Year, CompareWritten.Month, CompareWritten.Day, CompareWritten.Hour, CompareWritten.Minute, CompareWritten.Second)

                                    'is incoming file newer than previously processed
                                    If (DateTime.Compare(DateWritten, CompareWritten) > 0) And (DateDiff(DateInterval.Second, CompareWritten, DateWritten) > 25) Then

                                        ThisPDFDocument = New PdfDocument(New PdfReader(childFile))
                                        NumOfPDFPages = ThisPDFDocument.GetNumberOfPages()
                                        'get physical page number from file name
                                        PhysicalPageNum = CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3))

                                        For x = 1 To NumOfPDFPages
                                            ThisPDFPage = ThisPDFDocument.GetPage(x)
                                            ThisPageSize = ThisPDFPage.GetPageSize()
                                            'WriteLogFile(PubRegion & "|" & Edition & "|" & childFile.Name & "|" & CStr(NumOfPDFPages) & "|" & CStr(PhysicalPageNum) & "|" & childFile.Length.ToString() & "|" & childFile.LastWriteTime.ToString("dd-MM-yyyy HH:mm:ss"))

                                            'is it already a single page 827.7w x 1105.5h Points
                                            If ThisPageSize.GetHeight() < 1120 And ThisPageSize.GetWidth() < 840 Then
                                                'its a single just copy accross
                                                If NumOfPDFPages = 1 Then
                                                    If InsertPage = "IN" Then
                                                        'InsertPageNumber = InsertPageNumber + 1
                                                        NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"

                                                    Else

                                                        NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                                    End If
                                                    childFile.CopyTo(NewFileName, True)
                                                    File.SetCreationTime(NewFileName, DateCreated)
                                                    File.SetLastWriteTime(NewFileName, DateWritten)
                                                    LastFileWritten = NewFileName
                                                Else
                                                    'Write 2nd and subsequent ThisPDFPage to new file name
                                                    If InsertPage = "IN" Then
                                                        InsertPageNumber = InsertPageNumber + 1
                                                        NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                    Else

                                                        NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                                    End If

                                                    Using pdfNew = New PdfDocument(New PdfWriter(NewFileName))
                                                        oNewPage = ThisPDFPage.CopyTo(pdfNew)
                                                        pdfNew.AddPage(oNewPage)
                                                        pdfNew.Close()
                                                    End Using
                                                    File.SetCreationTime(NewFileName, DateCreated)
                                                    File.SetLastWriteTime(NewFileName, DateWritten)
                                                    LastFileWritten = NewFileName

                                                End If

                                                P = P + 1

                                            Else
                                                If ThisPageSize.GetHeight() < ((1105.5 * 2) - 15) And ThisPageSize.GetWidth() > 840 Then
                                                    'Its a spread even page take the left odd page take the right
                                                    If InsertPage = "IN" Then
                                                        ' InsertPageNumber = InsertPageNumber + 1

                                                        EvenPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                    Else
                                                        EvenPage = FullBookFolder & "\" & PageInFolder(P)
                                                    End If

                                                    P = P + 1
                                                    If P > (PageInFolder.Count - 1) Then
                                                        'If Theres no other file in array take page number from last page and increase by 1
                                                        P = P - 1
                                                        If InsertPage = "IN" Then
                                                            'InsertPageNumber = InsertPageNumber + 1

                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                        Else
                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3) + 1), "000") & ".PDF"
                                                        End If
                                                    Else
                                                        If InsertPage = "IN" Then
                                                            InsertPageNumber = InsertPageNumber + 1

                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"

                                                        Else
                                                            OddPage = FullBookFolder & "\" & PageInFolder(P)
                                                        End If
                                                    End If

                                                    Dim toMove As iText.Kernel.Geom.Rectangle
                                                    toMove = New iText.Kernel.Geom.Rectangle(getHalfPageSize(ThisPDFPage.GetPageSize()))

                                                    Dim pdfDoc As New PdfDocument(New PdfWriter(EvenPage, New WriterProperties().SetFullCompressionMode(True)))
                                                    pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                                    pdfDoc.AddNewPage()

                                                    'Left Side
                                                    Dim t_canvas1 As New PdfFormXObject(toMove)
                                                    Dim canvas1 As New PdfCanvas(t_canvas1, pdfDoc)
                                                    Dim pageXObject As PdfFormXObject
                                                    pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                                    canvas1.Rectangle(0, 0, toMove.GetWidth(), toMove.GetHeight())
                                                    canvas1.EoClip()
                                                    canvas1.EndPath()
                                                    canvas1.AddXObjectAt(pageXObject, 0, 0)
                                                    Dim canvas As New PdfCanvas(pdfDoc.GetFirstPage())
                                                    canvas.AddXObjectAt(t_canvas1, 0, 0)

                                                    pdfDoc.Close()
                                                    File.SetCreationTime(EvenPage, DateCreated)
                                                    File.SetLastWriteTime(EvenPage, DateWritten)
                                                    'Right Side
                                                    pdfDoc = New PdfDocument(New PdfWriter(OddPage, New WriterProperties().SetFullCompressionMode(True)))
                                                    pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                                    pdfDoc.AddNewPage()

                                                    toMove.SetX(ThisPageSize.GetWidth() - toMove.GetWidth())

                                                    Dim t_canvas2 As New PdfFormXObject(toMove)
                                                    Dim canvas2 As New PdfCanvas(t_canvas2, pdfDoc)
                                                    pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                                    canvas2.Rectangle(toMove.GetLeft(), toMove.GetBottom(), toMove.GetWidth(), toMove.GetHeight())
                                                    canvas2.Clip()
                                                    canvas2.EndPath()
                                                    canvas2.AddXObjectAt(pageXObject, 0, 0)

                                                    canvas = New PdfCanvas(pdfDoc.GetFirstPage())
                                                    canvas.AddXObjectAt(t_canvas2, 0, 0)

                                                    pdfDoc.Close()
                                                    File.SetCreationTime(OddPage, DateCreated)
                                                    File.SetLastWriteTime(OddPage, DateWritten)
                                                    LastFileWritten = OddPage
                                                    P = P + 1

                                                Else
                                                    'WTF have we got here

                                                End If


                                            End If

                                        Next
                                    Else
                                        If InsertPage = "IN" Then
                                            InsertPageCount = 0
                                            InsertPageCount = GetNumberOfSinglePages(childFile)
                                            If InsertPageCount >= 2 Then
                                                InsertPageNumber = InsertPageNumber + (InsertPageCount - 1)
                                            End If


                                        End If
                                    End If


                                End If    'This End if matches the indd file exists check

                            Next
                        Next
                        Call FillInTheBlanks(FullBookFolder, ComboBoxTitle.SelectedItem.Value, ComboBoxIssue.Text, LastFileWritten)
                    End If
                Next
            Next
        Catch ex As System.IO.FileNotFoundException
            WriteLogFile("Error: " & ex.Message)
        Catch ex As Exception
            WriteLogFile("Error: " & ex.Message)
        End Try

        Me.Cursor = Cursors.Default
        Me.Enabled = True


    End Sub
    Sub FillInTheBlanks(FullBookFolder As String, Title As String, Issue As String, LastFileWritten As String)

        Dim SourceDir As DirectoryInfo
        SourceDir = New DirectoryInfo(FullBookFolder)
        Dim PageNumber As Integer = 1
        Dim MissingPageName As String = ""
        Dim DateWritten As DateTime = DateAdd(DateInterval.Day, -7, Today())
        Dim DateCreated As DateTime = DateAdd(DateInterval.Day, -7, Today())
        Dim FileCount As Int16 = 0

        Dim path As String = "D:\BurdSiPageStore\PageProcess\" & Title & "\Blank_Page.JPG"
        Dim BlankPage As FileInfo = New FileInfo(path)
        Dim Gifpath As String = "D:\BurdSiPageStore\PageProcess\" & Title & "\Blank_Page.GIF"
        Dim GifBlankPage As FileInfo = New FileInfo(path)
        Dim FileExtension As String = ".pdf"


        If Not BlankPage.Exists Or Not GifBlankPage.Exists Then
            Exit Sub
        End If

        FileCount = SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = FileExtension And file.Name.ToLower.StartsWith(Title.ToLower + "_" + Issue + "_")).OrderBy(Function(fi) fi.Name).Count

        If FileCount = 0 Then
            FileExtension = ".bak"
        End If

        For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = FileExtension And file.Name.ToLower.StartsWith(Title.ToLower + "_" + Issue + "_")).OrderBy(Function(fi) fi.Name)




            Select Case CInt(Mid(childFile.Name, Len(childFile.Name) - 6, 3))
                Case PageNumber       ' ignore
                    PageNumber = PageNumber + 1
                Case < PageNumber     ' ignore
                Case > PageNumber     ' write blank pages until pagenumber = childfile.name

                    Do
                        If Not File.Exists(FullBookFolder & "\" & Mid(childFile.Name, 1, Len(childFile.Name) - 7) & Format(PageNumber, "000") & ".bak") Then
                            MissingPageName = FullBookFolder & "\JPG\" & Mid(childFile.Name, 1, Len(childFile.Name) - 7) & Format(PageNumber, "000") & ".JPG"
                            BlankPage.CopyTo(MissingPageName, True)
                            File.SetCreationTime(MissingPageName, DateCreated)
                            File.SetLastWriteTime(MissingPageName, DateWritten)
                            MissingPageName = FullBookFolder & "\JPG\" & Mid(childFile.Name, 1, Len(childFile.Name) - 7) & Format(PageNumber, "000") & ".GIF"
                            GifBlankPage.CopyTo(MissingPageName, True)
                            File.SetCreationTime(MissingPageName, DateCreated)
                            File.SetLastWriteTime(MissingPageName, DateWritten)
                        End If
                        PageNumber = PageNumber + 1
                    Loop While PageNumber < CInt(Mid(childFile.Name, Len(childFile.Name) - 6, 3))
                    PageNumber = PageNumber + 1
            End Select
        Next


        BlankPage = Nothing
        GifBlankPage = Nothing
        SourceDir = Nothing
    End Sub

    Private Sub Button99_Click_1(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Me.Enabled = False
        Try
            Dim PageFolder As String = "NAT"
            Dim FullBookFolder As String
            Dim ThisPDFDocument As PdfDocument
            Dim ThisPDFPage As PdfPage
            Dim oNewPage As PdfPage
            Dim ThisPageSize As iText.Kernel.Geom.Rectangle
            Dim NumOfPDFPages As Integer
            Dim PhysicalPageNum As Integer
            Dim SourceDir As DirectoryInfo
            Dim NewFileName As String

            Dim EvenPage As String
            Dim OddPage As String
            Dim DateWritten As DateTime
            Dim DateCreated As DateTime
            Dim CompareWritten As DateTime
            Dim PageInFolder() As String
            'Dim PageSize() As Long
            'Dim PageMatch() As Boolean
            Dim PagesToProcess As Integer
            Dim P As Integer
            Dim LastFileWritten As String = "0000000"
            Dim InsertPage As String = ""
            Dim InsertPageNumber As Integer

            For Each Edition In {"1", "2"}
                For Each PubRegion In {"Ireland", "National", "Scotland"}
                    If Directory.Exists("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion) Then
                        If PubRegion = "Ireland" Then
                            PageFolder = "IRE (IDM)"
                        End If
                        If PubRegion = "National" Then
                            PageFolder = "NAT "
                        End If
                        If PubRegion = "Scotland" Then
                            PageFolder = "SCT (SCT)"
                        End If
                        FullBookFolder = "D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion & "\" & "FullBook"
                        If Directory.Exists(FullBookFolder) Then
                            'Delete the PDF's already in it
                            'For Each deleteFile In Directory.GetFiles(FullBookFolder, "*.PDF", SearchOption.TopDirectoryOnly)
                            'If File.Exists(Replace(deleteFile.ToLower(), ".pdf", ".bak")) Then
                            'File.Delete(Replace(deleteFile.ToLower(), ".pdf", ".bak"))
                            'End If
                            '   File.Copy(deleteFile, Replace(deleteFile.ToLower(), ".pdf", ".bak"), True)
                            '  File.GetLastWriteTime(deleteFile)
                            ' File.SetCreationTime(Replace(deleteFile.ToLower(), ".pdf", ".bak"), File.GetCreationTime(deleteFile))
                            'File.SetLastWriteTime(Replace(deleteFile.ToLower(), ".pdf", ".bak"), File.GetLastWriteTime(deleteFile))
                            'File.Delete(deleteFile)

                            'Next
                            'Directory.Delete(FullBookFolder, True)
                        Else
                            Directory.CreateDirectory(FullBookFolder)
                        End If

                        'Directory.CreateDirectory(FullBookFolder)
                        For Each LocalDir In My.Computer.FileSystem.GetDirectories("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion, FileIO.SearchOption.SearchTopLevelOnly, "*")
                            '.Where(Function(RDir) RDir.Contains(CStr(Edition) & " " & PageFolder & " " & Format(CDate(ComboBoxIssue.Text), "ddMM")) And Not RDir.Contains("\" & CStr(Edition) & "\" & PubRegion & "\IN"))
                            SourceDir = New DirectoryInfo(LocalDir)

                            PagesToProcess = SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).Count
                            'WriteLogFile(LocalDir & "  " & PagesToProcess.ToString())
                            ReDim PageInFolder(PagesToProcess - 1)
                            'ReDim PageSize(PagesToProcess - 1)
                            'ReDim PageMatch(PagesToProcess - 1)
                            P = -1
                            For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Name)
                                P = P + 1
                                PageInFolder(P) = childFile.Name
                                '   PageSize(P) = childFile.Length
                                '  If P >= 1 Then
                                ' If PageSize(P) = PageSize(P - 1) Then
                                'PageMatch(P) = True
                                'Else
                                'PageMatch(P) = False
                                'End If
                                'End If

                            Next
                            'If P >= 1 Then
                            'If PageSize(0) = PageSize(1) Then
                            'PageMatch(0) = True
                            'Else
                            'PageMatch(0) = False
                            'End If
                            'End If

                            For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.TopDirectoryOnly).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500).OrderBy(Function(fi) fi.Name)

                                InsertPage = ""
                                If Mid(SourceDir.Name, 1, 2) = "IN" Then
                                    InsertPage = "IN"
                                    InsertPageNumber = CInt(Mid(LastFileWritten, Len(LastFileWritten) - 6, 3))
                                End If
                                NewFileName = SourceDir.FullName & "\" & InsertPage & Mid(childFile.Name, Len(childFile.Name) - 6, 3) & " " & Edition & " " & PageFolder & " " & Mid(Replace(ComboBoxIssue.Text, "-", ""), 1, 4) & ".indd"
                                'If File.Exists(NewFileName) Then   'Removing this line prevent the indd exists check therefore now a strict newest file wins senario

                                'matched .pdf to .indd, so got first file in folder
                                'if remaining files in folder are same sizes, then each filename has page number
                                'first file will have as many pages as files in folder
                                'if not matching file size treat as a single or spread page evens left odds right, single page either applies

                                'make sure starting P matches the index of the childFile found
                                For P = 0 To PageInFolder.Count - 1
                                    If PageInFolder(P) = childFile.Name Then
                                        Exit For
                                    End If
                                Next



                                DateCreated = childFile.CreationTime
                                DateCreated = New DateTime(DateCreated.Year, DateCreated.Month, DateCreated.Day, DateCreated.Hour, DateCreated.Minute, DateCreated.Second)
                                DateWritten = childFile.LastWriteTime
                                DateWritten = New DateTime(DateWritten.Year, DateWritten.Month, DateWritten.Day, DateWritten.Hour, DateWritten.Minute, DateWritten.Second)
                                'CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & Replace(childFile.Name.ToLower(), ".pdf", ".bak"))
                                CompareWritten = File.GetLastWriteTime(FullBookFolder & "\" & childFile.Name.ToLower())
                                CompareWritten = New DateTime(CompareWritten.Year, CompareWritten.Month, CompareWritten.Day, CompareWritten.Hour, CompareWritten.Minute, CompareWritten.Second)
                                'is incoming file newer than previously processed
                                'is incoming file newer than previously processed
                                If DateTime.Compare(DateWritten, CompareWritten) > 0 Then
                                    If DateDiff(DateInterval.Second, CompareWritten, DateWritten) > 25 Then
                                        'If DateWritten > File.GetLastWriteTime(FullBookFolder & "\" & childFile.Name) Then

                                        ThisPDFDocument = New PdfDocument(New PdfReader(childFile))
                                        NumOfPDFPages = ThisPDFDocument.GetNumberOfPages()
                                        'get physical page number from file name
                                        PhysicalPageNum = CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3))

                                        For x = 1 To NumOfPDFPages
                                            ThisPDFPage = ThisPDFDocument.GetPage(x)
                                            ThisPageSize = ThisPDFPage.GetPageSize()
                                            'WriteLogFile(PubRegion & "|" & Edition & "|" & childFile.Name & "|" & CStr(NumOfPDFPages) & "|" & CStr(PhysicalPageNum) & "|" & childFile.Length.ToString() & "|" & childFile.LastWriteTime.ToString("dd-MM-yyyy HH:mm:ss"))

                                            'is it already a single page 827.7w x 1105.5h Points
                                            If ThisPageSize.GetHeight() < 1120 And ThisPageSize.GetWidth() < 840 Then
                                                'its a single just copy accross
                                                If NumOfPDFPages = 1 Then
                                                    'NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                                    If InsertPage = "IN" Then
                                                        InsertPageNumber = InsertPageNumber + 1
                                                        NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                    Else

                                                        NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                                    End If
                                                    childFile.CopyTo(NewFileName, True)
                                                    File.SetCreationTime(NewFileName, DateCreated)
                                                    File.SetLastWriteTime(NewFileName, DateWritten)
                                                    LastFileWritten = NewFileName
                                                Else
                                                    'Write 2nd and subsequent ThisPDFPage to new file name
                                                    If InsertPage = "IN" Then
                                                        InsertPageNumber = InsertPageNumber + 1
                                                        NewFileName = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                    Else

                                                        NewFileName = FullBookFolder & "\" & PageInFolder(P)
                                                    End If

                                                    Using pdfNew = New PdfDocument(New PdfWriter(NewFileName))
                                                        oNewPage = ThisPDFPage.CopyTo(pdfNew)
                                                        pdfNew.AddPage(oNewPage)
                                                        pdfNew.Close()
                                                    End Using
                                                    File.SetCreationTime(NewFileName, DateCreated)
                                                    File.SetLastWriteTime(NewFileName, DateWritten)
                                                    LastFileWritten = NewFileName

                                                End If

                                                P = P + 1

                                            Else
                                                If ThisPageSize.GetHeight() < ((1105.5 * 2) - 15) And ThisPageSize.GetWidth() > 840 Then
                                                    'Its a spread even page take the left odd page take the right
                                                    If InsertPage = "IN" Then
                                                        InsertPageNumber = InsertPageNumber + 1
                                                        EvenPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                    Else
                                                        EvenPage = FullBookFolder & "\" & PageInFolder(P)
                                                    End If

                                                    'EvenPage = FullBookFolder & "\" & PageInFolder(P)

                                                    P = P + 1
                                                    If P > (PageInFolder.Count - 1) Then
                                                        'If Theres no other file in array take page number from last page and increase by 1
                                                        P = P - 1
                                                        If InsertPage = "IN" Then
                                                            InsertPageNumber = InsertPageNumber + 1
                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                        Else
                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3) + 1), "000") & ".PDF"
                                                        End If


                                                        'OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(CInt(Mid(PageInFolder(P), Len(PageInFolder(P)) - 6, 3) + 1), "000") & ".PDF"
                                                    Else
                                                        If InsertPage = "IN" Then
                                                            InsertPageNumber = InsertPageNumber + 1
                                                            OddPage = FullBookFolder & "\" & Mid(PageInFolder(P), 1, Len(PageInFolder(P)) - 7) & Format(InsertPageNumber, "000") & ".PDF"
                                                        Else
                                                            OddPage = FullBookFolder & "\" & PageInFolder(P)
                                                        End If
                                                        'OddPage = FullBookFolder & "\" & PageInFolder(P)
                                                    End If

                                                    Dim toMove As iText.Kernel.Geom.Rectangle
                                                    toMove = New iText.Kernel.Geom.Rectangle(getHalfPageSize(ThisPDFPage.GetPageSize()))

                                                    Dim pdfDoc As New PdfDocument(New PdfWriter(EvenPage, New WriterProperties().SetFullCompressionMode(True)))
                                                    pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                                    pdfDoc.AddNewPage()

                                                    'Left Side
                                                    Dim t_canvas1 As New PdfFormXObject(toMove)
                                                    Dim canvas1 As New PdfCanvas(t_canvas1, pdfDoc)
                                                    Dim pageXObject As PdfFormXObject
                                                    pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                                    canvas1.Rectangle(0, 0, toMove.GetWidth(), toMove.GetHeight())
                                                    canvas1.EoClip()
                                                    canvas1.EndPath()
                                                    canvas1.AddXObjectAt(pageXObject, 0, 0)
                                                    Dim canvas As New PdfCanvas(pdfDoc.GetFirstPage())
                                                    canvas.AddXObjectAt(t_canvas1, 0, 0)

                                                    pdfDoc.Close()
                                                    File.SetCreationTime(EvenPage, DateCreated)
                                                    File.SetLastWriteTime(EvenPage, DateWritten)
                                                    'Right Side
                                                    pdfDoc = New PdfDocument(New PdfWriter(OddPage, New WriterProperties().SetFullCompressionMode(True)))
                                                    pdfDoc.SetDefaultPageSize(New iText.Kernel.Geom.PageSize(toMove))
                                                    pdfDoc.AddNewPage()

                                                    toMove.SetX(ThisPageSize.GetWidth() - toMove.GetWidth())

                                                    Dim t_canvas2 As New PdfFormXObject(toMove)
                                                    Dim canvas2 As New PdfCanvas(t_canvas2, pdfDoc)
                                                    pageXObject = ThisPDFPage.CopyAsFormXObject(pdfDoc)
                                                    canvas2.Rectangle(toMove.GetLeft(), toMove.GetBottom(), toMove.GetWidth(), toMove.GetHeight())
                                                    canvas2.Clip()
                                                    canvas2.EndPath()
                                                    canvas2.AddXObjectAt(pageXObject, 0, 0)

                                                    canvas = New PdfCanvas(pdfDoc.GetFirstPage())
                                                    canvas.AddXObjectAt(t_canvas2, 0, 0)

                                                    pdfDoc.Close()
                                                    File.SetCreationTime(OddPage, DateCreated)
                                                    File.SetLastWriteTime(OddPage, DateWritten)
                                                    LastFileWritten = OddPage
                                                    P = P + 1

                                                Else
                                                    'WTF have we got here

                                                End If


                                            End If

                                        Next
                                    End If
                                End If
                                'End If    'This End if matches the indd file exists check

                            Next
                        Next
                        Call FillInTheBlanks(FullBookFolder, ComboBoxTitle.SelectedItem.Value, ComboBoxIssue.Text, LastFileWritten)
                    End If
                Next
            Next
        Catch ex As System.IO.FileNotFoundException
            WriteLogFile("Error: " & ex.Message)
        Catch ex As Exception
            WriteLogFile("Error: " & ex.Message)
        End Try

        Me.Cursor = Cursors.Default
        Me.Enabled = True


    End Sub

    Function GetNumberOfSinglePages(PDF_Filename As FileInfo)

        Dim ThisPDFDocument As PdfDocument
        Dim ThisPDFPage As PdfPage
        Dim NumOfPDFPages As Integer
        Dim PhysicalPages As Integer = 0
        Dim ThisPageSize As iText.Kernel.Geom.Rectangle

        ThisPDFDocument = New PdfDocument(New PdfReader(PDF_Filename))
        NumOfPDFPages = ThisPDFDocument.GetNumberOfPages()
        For x = 1 To NumOfPDFPages
            ThisPDFPage = ThisPDFDocument.GetPage(x)
            ThisPageSize = ThisPDFPage.GetPageSize()
            'is it a single page 827.7w x 1105.5h Points
            If ThisPageSize.GetHeight() < 1120 And ThisPageSize.GetWidth() < 840 Then
                PhysicalPages = PhysicalPages + 1
            ElseIf ThisPageSize.GetHeight() < ((1105.5 * 2) - 15) And ThisPageSize.GetWidth() > 840 Then
                PhysicalPages = PhysicalPages + 2
            End If
        Next

        ThisPDFDocument.Close()
        If ThisPageSize IsNot Nothing Then ThisPageSize = Nothing
        If ThisPDFPage IsNot Nothing Then ThisPDFPage = Nothing
        If ThisPDFDocument IsNot Nothing Then ThisPDFDocument = Nothing

        Return PhysicalPages

    End Function

    Sub CheckForDuplicateInsertPDF(Edition As String, PubRegion As String)

        Dim SourceDir As DirectoryInfo

        Dim FoundDupFile() As String
        Dim FoundLastWrite() As DateTime
        Dim FoundDupDirectory() As String
        Dim FoundCount As Integer = -1

        Dim LastFile As String = ""
        Dim LastWrite As DateTime
        Dim LastDirectory As String = ""
        Dim SameFile As Boolean = False

        SourceDir = New DirectoryInfo("D:\BurdSiPageStore\PageProcess\" & ComboBoxTitle.SelectedItem.Value & "\" & ComboBoxIssue.Text & "\" & CStr(Edition) & "\" & PubRegion)

        For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.AllDirectories).Where(Function(file) file.Extension.ToLower = ".pdf" And file.Name.StartsWith(ComboBoxTitle.SelectedItem.Value + "_" + ComboBoxIssue.Text + "_") And file.Length > 16500 And file.Directory.Name.ToLower.StartsWith("in")).OrderBy(Function(fi) fi.Name).ThenBy(Function(fi) fi.LastWriteTime)

            If LastFile = "" Then
                LastFile = childFile.Name
                LastDirectory = childFile.DirectoryName
                LastWrite = childFile.LastWriteTime
                Continue For
            End If
            If LastFile = childFile.Name Then
                If Not SameFile Then
                    FoundCount = FoundCount + 1
                    ReDim Preserve FoundDupFile(FoundCount), FoundLastWrite(FoundCount), FoundDupDirectory(FoundCount)
                    FoundDupFile(FoundCount) = LastFile
                    FoundLastWrite(FoundCount) = LastWrite
                    FoundDupDirectory(FoundCount) = LastDirectory
                    SameFile = True

                End If
                FoundCount = FoundCount + 1
                ReDim Preserve FoundDupFile(FoundCount), FoundLastWrite(FoundCount), FoundDupDirectory(FoundCount)
                FoundDupFile(FoundCount) = childFile.Name
                FoundLastWrite(FoundCount) = childFile.LastWriteTime
                FoundDupDirectory(FoundCount) = childFile.DirectoryName
            Else
                LastFile = childFile.Name
                LastDirectory = childFile.DirectoryName
                LastWrite = childFile.LastWriteTime
                SameFile = False
            End If


        Next


        'Now process array to remove .indd of duplicate file, array must hold at least 2 files
        Dim LoopCount As Integer = 0
        Dim PageFolder As String
        Dim INDDFileName As String
        If PubRegion = "Ireland" Then
            PageFolder = "IRE (IDM)"
        End If
        If PubRegion = "National" Then
            PageFolder = "NAT "
        End If
        If PubRegion = "Scotland" Then
            PageFolder = "SCT (SCT)"
        End If
        If FoundCount >= 1 Then
            Do
                If FoundDupFile(LoopCount) = FoundDupFile(LoopCount + 1) Then
                    FoundLastWrite(LoopCount) = New DateTime(FoundLastWrite(LoopCount).Year, FoundLastWrite(LoopCount).Month, FoundLastWrite(LoopCount).Day, FoundLastWrite(LoopCount).Hour, FoundLastWrite(LoopCount).Minute, FoundLastWrite(LoopCount).Second)
                    FoundLastWrite(LoopCount + 1) = New DateTime(FoundLastWrite(LoopCount + 1).Year, FoundLastWrite(LoopCount + 1).Month, FoundLastWrite(LoopCount + 1).Day, FoundLastWrite(LoopCount + 1).Hour, FoundLastWrite(LoopCount + 1).Minute, FoundLastWrite(LoopCount + 1).Second)
                    If (DateTime.Compare(FoundLastWrite(LoopCount), FoundLastWrite(LoopCount + 1)) > 0) And (DateDiff(DateInterval.Second, FoundLastWrite(LoopCount), FoundLastWrite(LoopCount + 1)) > 25) Then
                        'Delete loopCount + 1
                        INDDFileName = FoundDupDirectory(LoopCount + 1) & "\IN" & Mid(FoundDupFile(LoopCount + 1), Len(FoundDupFile(LoopCount + 1)) - 6, 3) & " " & Edition & " " & PageFolder & " " & Mid(Replace(ComboBoxIssue.Text, "-", ""), 1, 4) & ".indd"
                        If File.Exists(FoundDupDirectory(LoopCount + 1) & "\" & FoundDupFile(LoopCount + 1)) Then
                            File.Delete(FoundDupDirectory(LoopCount + 1) & "\" & FoundDupFile(LoopCount + 1))
                        End If
                        If File.Exists(INDDFileName) Then
                            File.Delete(INDDFileName)
                        End If

                    Else
                        'Delete Loopcount
                        INDDFileName = FoundDupDirectory(LoopCount) & "\IN" & Mid(FoundDupFile(LoopCount), Len(FoundDupFile(LoopCount)) - 6, 3) & " " & Edition & " " & PageFolder & " " & Mid(Replace(ComboBoxIssue.Text, "-", ""), 1, 4) & ".indd"
                        If File.Exists(FoundDupDirectory(LoopCount) & "\" & FoundDupFile(LoopCount)) Then
                            File.Delete(FoundDupDirectory(LoopCount) & "\" & FoundDupFile(LoopCount))
                        End If
                        If File.Exists(INDDFileName) Then
                            File.Delete(INDDFileName)
                        End If
                    End If
                End If

                LoopCount = LoopCount + 1
            Loop While LoopCount < FoundCount

        End If





    End Sub


End Class


