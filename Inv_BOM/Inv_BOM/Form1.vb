Imports System.IO
Imports System.String
Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.ComponentModel

Public Class Form1
    Public title_counter As Integer
    Public filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Public filepath3 As String = "c:\MyDir"
    Public G1_row_cnt As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.ColumnCount = 30
        DataGridView1.RowCount = 1000
        DataGridView1.ColumnHeadersVisible = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        DataGridView1.Columns(0).HeaderText = "File"
        DataGridView1.Columns(1).HeaderText = "Item "
        DataGridView1.Columns(2).HeaderText = "Qty"
        DataGridView1.Columns(3).HeaderText = "Part"
        DataGridView1.Columns(4).HeaderText = "Desc"
        DataGridView1.Columns(5).HeaderText = "Stock"

        DataGridView1.Columns(6).HeaderText = "DOC_NUMBER"
        DataGridView1.Columns(7).HeaderText = "ITEM_NR"
        DataGridView1.Columns(8).HeaderText = "DOC_STATUS"
        DataGridView1.Columns(9).HeaderText = "DOC_REV"
        DataGridView1.Columns(10).HeaderText = "PART_MATERIAL"
        DataGridView1.Columns(11).HeaderText = "IT_TP"
        DataGridView1.Columns(12).HeaderText = "LENGTH"
        DataGridView1.Columns(13).HeaderText = "Part Icon"

        DataGridView1.Columns(14).HeaderText = "Title"
        DataGridView1.Columns(15).HeaderText = "Subject"
        DataGridView1.Columns(16).HeaderText = "Author"
        DataGridView1.Columns(17).HeaderText = "Comments"

        DataGridView2.ColumnCount = 10
        DataGridView2.RowCount = 1000
        DataGridView2.Columns(0).HeaderText = "File"
        DataGridView2.Columns(1).HeaderText = "Descrp"
        DataGridView2.Columns(2).HeaderText = "D_no"
        DataGridView2.Columns(3).HeaderText = "A_no"
        DataGridView2.Columns(4).HeaderText = "Rev"
        DataGridView2.Columns(5).HeaderText = "Status"

        DataGridView3.ColumnCount = 5
        DataGridView3.RowCount = 20
        DataGridView3.Columns(0).HeaderText = "File"
        DataGridView3.Columns(1).HeaderText = "D_no"
        DataGridView3.Columns(2).HeaderText = "A_no"
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Open_file(1)   'ipt files
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Open_file(2)    'iam files
    End Sub

    Private Sub Open_file(keuze As Integer)
        ' Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog With {
               .InitialDirectory = "c:\Inventor test files\",
               .Filter = "Part File (*.ipt)|*.ipt" _
               & "|Assembly File (*.iam)|*.iam" _
               & "|Presentation File (*.ipn)|*.ipn" _
               & "|Drawing File (*.idw)|*.idw" _
               & "|Drawing File (*.dwg)|*.dwg" _
               & "|Design element File (*.ide)|*.ide",
               .FilterIndex = keuze,                ' *.ipt files
               .RestoreDirectory = True
           }

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                filepath1 = openFileDialog1.FileName
                TextBox1.Text = filepath1.ToString
                Get_dwg_art_nr()
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button3.BackColor = System.Drawing.Color.Green
        DataGridView1.ClearSelection()
        G1_row_cnt = 0
        Qbom(filepath1)
        Button3.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Qbom(ByVal fpath As String)
        Dim information As System.IO.FileInfo
        Dim filen As String

        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        '------- get file info -----------
        information = My.Computer.FileSystem.GetFileInfo(fpath)
        filen = information.Name

        Dim oDoc As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        invApp.SilentOperation = vbTrue
        oDoc = CType(invApp.Documents.Open(fpath, False), Document)

        '--------- determine object type -------
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        If eDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Please Select a IAM file ")
            Exit Sub
        End If

        '-------------READ TITLE BLOCK----------------------------------------
        ' ---- Note: there is no title block in an IAmmodel file -------------

        '---------- Read BOM --------------------------
        'Try
        Dim oBOM As Inventor.BOM
            oBOM = oDoc.ComponentDefinition.BOM
            oBOM.StructuredViewFirstLevelOnly = True
            oBOM.StructuredViewEnabled = True

            Dim oBOMView As Inventor.BOMView
            oBOMView = oBOM.BOMViews.Item("Structured")

            '-------------------------
            Dim oRow As BOMRow
            Dim oCompDef As ComponentDefinition
            Dim oPropSet As PropertySet
            Dim i, j As Integer

            For i = 1 To oBOMView.BOMRows.Count
                G1_row_cnt += 1

                '================= Design Tracking Properties ==========================
                oRow = oBOMView.BOMRows.Item(i)
                oCompDef = oRow.ComponentDefinitions.Item(1)

                oPropSet = oCompDef.Document.PropertySets.Item("Design Tracking Properties")
                DataGridView1.Rows.Add()

                DataGridView1.Rows.Item(G1_row_cnt).Cells(0).Value = filen

                DataGridView1.Rows.Item(G1_row_cnt).Cells(1).Value = oRow.ItemNumber
                DataGridView1.Rows.Item(G1_row_cnt).Cells(2).Value = oRow.ItemQuantity

                Dim design_track() As String =
                {"Part Number",
                "Description",
                "Stock Number",
                "Part Icon"}
                If oPropSet.Count = 0 And Not CheckBox1.Checked Then
                    MessageBox.Show("The are NO 'Design Tracking' properties present in this file")
                Else
                    For j = 0 To design_track.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = "+"
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = oPropSet.Item(design_track(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = "?"
                            If Not CheckBox1.Checked Then MessageBox.Show(design_track(j) & " not found")
                        End Try
                    Next
                End If

                '================== CUSTOM Properties ============================
                Dim custom() As String =
                {"DOC_NUMBER",
                "ITEM_NR",
                "DOC_STATUS",
                "DOC_REV",
                "PART_MATERIAL",
                "IT_TP",
                "LG"}

                oPropSet = oCompDef.Document.PropertySets.Item("Inventor User Defined Properties")
                If oPropSet.Count = 0 And Not CheckBox1.Checked Then
                    MessageBox.Show("The are NO 'Custom' properties present in this file")
                Else
                    For j = 0 To custom.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = "+"
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = oPropSet.Item(custom(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = "?"
                            If Not CheckBox1.Checked Then MessageBox.Show(custom(j) & " not found")
                        End Try
                    Next
                End If

                '========== Inventor Summary Information ===============
                Dim summary() As String =
                {"Title",
                "Subject",
                "Author",
                "Comments"}
                oPropSet = oCompDef.Document.PropertySets.Item("Inventor Summary Information")
                If oPropSet.Count = 0 And Not CheckBox1.Checked Then
                    MessageBox.Show("The are NO 'Inventor Summary Information' present in this file")
                Else
                    For j = 0 To summary.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = "+"
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = oPropSet.Item(summary(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = "?"
                            If Not CheckBox1.Checked Then MessageBox.Show(summary(j) & " not found")
                        End Try
                    Next
                End If

            Next
        'Catch Ex As Exception
        'MessageBox.Show("No BOM in this IAM model")
        'Finally
        'End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.BackColor = System.Drawing.Color.Green
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = filepath3
        SaveFileDialog1.FileName = "_BOM" & "_" & TextBox3.Text & "_" & TextBox4.Text & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView1)
        Button4.BackColor = System.Drawing.Color.Transparent
    End Sub
    Private Sub Write_excel(ByVal dg As DataGridView)
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorksheet As Excel.Worksheet
        Dim fname As String
        Dim str As String

        xlApp = CreateObject("Excel.Application")
        xlWorkBook = xlApp.Workbooks.Add(Type.Missing)
        xlWorksheet = xlWorkBook.Worksheets(1)

        xlApp.Visible = False
        xlApp.DisplayAlerts = False 'Suppress excel messages

        '-------- Header text to excel -------------
        For hor = 0 To dg.Columns.Count - 1
            str = dg.Columns(hor).HeaderText
            xlWorksheet.Cells(1, hor + 1) = str
        Next

        TextBox2.Text &= "Rows...." & dg.Rows.Count & vbCrLf
        TextBox2.Text &= "Columns...." & dg.Columns.Count & vbCrLf

        '-------- Cell_text to excel -------------
        Try
            For vert = 0 To dg.Rows.Count - 1
                For hor = 0 To dg.Columns.Count - 1
                    If Not dg.Rows.Item(vert).Cells(hor).Value Is Nothing Then
                        str = dg.Rows.Item(vert).Cells(hor).Value.ToString
                        TextBox2.Text &= "hor=" & hor.ToString & " vert=" & vert.ToString & " str= " & str & vbCrLf
                        xlWorksheet.Cells(vert + 2, hor + 1) = str
                    End If
                Next
            Next
            fname = SaveFileDialog1.FileName
            xlWorkBook.SaveAs(fname, FileFormat:=XlFileFormat.xlWorkbookNormal)
            xlWorkBook.Close()
            xlApp.Quit()

            Marshal.ReleaseComObject(xlWorksheet)
            Marshal.ReleaseComObject(xlWorkBook)
            Marshal.ReleaseComObject(xlApp)

        Catch ex As Exception
            MessageBox.Show("Problem writing excel " & ex.Message)
        End Try

    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        List_all_properties()
    End Sub
    'see https://forums.autodesk.com/t5/inventor-customization/ilogic-list-all-custom-properties/td-p/6218163
    Private Sub List_all_properties()
        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        '-------- Now list properties--------
        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.Document

        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue
        oDoc = CType(invApp.Documents.Open(filepath1, False), Document)

        Dim Docs As DocumentsEnumerator = oDoc.AllReferencedDocuments
        Dim aDoc As Document
        Dim Pros As New ArrayList
        Dim item As String
        For Each aDoc In Docs
            Dim oPropsets As PropertySets
            oPropsets = oDoc.PropertySets
            Dim oPropSet As PropertySet

            Select Case True
                Case RadioButton1.Checked
                    oPropSet = oPropsets.Item("Inventor User Defined Properties")
                Case RadioButton2.Checked
                    oPropSet = oPropsets.Item(RadioButton2.Text)
                Case RadioButton3.Checked
                    oPropSet = oPropsets.Item(RadioButton3.Text)
                Case Else
                    oPropSet = oPropsets.Item(RadioButton4.Text)
            End Select

            Dim oPro As Inventor.Property
            For Each oPro In oPropSet
                Dim Found As Boolean = False
                For Each item In Pros
                    If oPro.Name = item Then Found = True
                Next
                If Found = False Then
                    Pros.Add(oPro.Name)
                End If
            Next
        Next

        Dim AllPros As String = "List of all used iProperties:"
        For Each item In Pros
            AllPros = AllPros & vbLf & item
        Next
        MsgBox(AllPros)
    End Sub
    'Read IDW Title Block
    Public Sub Read_title_Block(ByVal path As String)
        'http://adndevblog.typepad.com/manufacturing/2012/12/inventor-change-text-items-in-titleblockdefinition.html

        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.Document

        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue
        Try
            oDoc = CType(invApp.Documents.Open(path, False), Document)

            '--------- determine object type -------
            Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
            If eDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                MessageBox.Show("Please Select a IDW file ")
            Else
                '=================================================================================
                'https://forums.autodesk.com/t5/inventor-customization/copy-titleblock-prompted-entries-to-custom-iproperty/td-p/7491136

                Dim oSheet As Sheet
                oSheet = oDoc.ActiveSheet
                Dim oTB1 As TitleBlock
                oTB1 = oSheet.TitleBlock
                Dim titleDef As TitleBlockDefinition
                titleDef = oTB1.Definition
                Dim oPrompt As Inventor.TextBox = Nothing

                ' Find the Prompted Entry called Make in the Title Block
                For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
                    DataGridView2.Rows.Item(title_counter).Cells(0).Value = path
                    If defText.Text = "<DESCRIPTION>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(1).Value = oTB1.GetResultText(oPrompt)
                    End If
                    If defText.Text = "<ITEM_NR>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(2).Value = oTB1.GetResultText(oPrompt)
                    End If
                    If defText.Text = "<DOC_NUMBER>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(3).Value = oTB1.GetResultText(oPrompt)
                    End If
                    If defText.Text = "<DOC_REV>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(4).Value = oTB1.GetResultText(oPrompt)
                    End If
                    If defText.Text = "<DOC_STATUS>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(5).Value = oTB1.GetResultText(oPrompt)
                    End If
                Next
            End If
            oDoc.Close()
        Catch
        End Try
    End Sub

    Private Sub Getresulttext(titleBlock As TitleBlock)
        Throw New NotImplementedException()
    End Sub

    Private Sub Get_dwg_art_nr()
        Dim s, substring As String
        Dim length As Integer
        Dim searchDoc As String = "_D"
        Dim searchArt As String = "_A"
        Dim startindex, endIndex As Integer

        TextBox3.Text = ""
        TextBox4.Text = ""

        s = TextBox1.Text
        length = s.Length
        If length >= 23 Then
            s = s.Substring(length - 22, 18)
            'MessageBox.Show(s)
            startindex = s.IndexOf(searchDoc)
            endIndex = startindex + 7
            If startindex >= 0 Then
                substring = s.Substring(startindex, endIndex + searchDoc.Length - startindex)
                TextBox3.Text = substring.Substring(1, 8)
            End If

            startindex = s.IndexOf(searchArt)
            endIndex = startindex + 7
            If startindex >= 0 Then
                substring = s.Substring(startindex, endIndex + searchArt.Length - startindex)
                TextBox4.Text = substring.Substring(1, 8)
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Button9.BackColor = System.Drawing.Color.Green
        title_counter = -1   'Reset counter

        'Select work directory
        'https://msdn.microsoft.com/en-us/library/07wt70x2(v=vs.110).aspx
        Dim pathfile As String
        pathfile = TextBox6.Text

        If IO.File.Exists(pathfile) Then ' This pathfile is a file.
            ProcessFile(pathfile)
        Else
            If Directory.Exists(pathfile) Then
                ProcessDirectory(pathfile)   ' This path is a directory.
            Else
                MessageBox.Show(pathfile & " is not a valid file or directory.")
            End If
        End If
        Button9.BackColor = System.Drawing.Color.Transparent
    End Sub
    ' Process all files in the directory passed in, recurse on any directories 
    ' that are found, and process the files they contain.

    Private Sub ProcessDirectory(ByVal targetDirectory As String)
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory.
        Dim fileName As String
        For Each fileName In fileEntries
            ProcessFile(fileName)
        Next fileName

        Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
        ' Recurse into subdirectories of this directory.
        Dim subdirectory As String
        For Each subdirectory In subdirectoryEntries
            ProcessDirectory(subdirectory)
        Next subdirectory
    End Sub

    ' Processing found files 
    Private Sub ProcessFile(ByVal file As String)
        'MessageBox.Show("Processed file is " & file)
        Dim extension As String = IO.Path.GetExtension(file)
        If extension = ".idw" Then
            title_counter += 1
            Read_title_Block(file)
        End If
    End Sub
    'Select work directory
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        FolderBrowserDialog1.SelectedPath = TextBox6.Text
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox5.Text = FolderBrowserDialog1.SelectedPath
            TextBox6.Text = FolderBrowserDialog1.SelectedPath
            TextBox7.Text = FolderBrowserDialog1.SelectedPath
            TextBox8.Text = FolderBrowserDialog1.SelectedPath
            TextBox9.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Button7.BackColor = System.Drawing.Color.Green
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = TextBox6.Text
        SaveFileDialog1.FileName = "_Title_Blocks" & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView2)
        Button7.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Button8.BackColor = System.Drawing.Color.Green
        Dim cnt As Integer = 0   'Reset counter
        Dim fext As String = ".dxf"
        Dim extension As String

        Select Case True
            Case RadioButton5.Checked
                fext = ".dxf"
            Case RadioButton6.Checked
                fext = ".iam"
            Case RadioButton7.Checked
                fext = ".ipt"
            Case RadioButton8.Checked
                fext = ".idw"
            Case RadioButton9.Checked
                fext = ".*"
        End Select
        DataGridView3.Rows.Clear()
        DataGridView3.Columns(0).Width = 300


        Dim fileEntries As String() = Directory.GetFiles(TextBox6.Text)
        ' list DXF files found in the directory.
        Dim fileName As String

        For Each fileName In fileEntries
            extension = IO.Path.GetExtension(fileName)
            If String.Equals(extension, fext) Or RadioButton9.Checked Then
                DataGridView3.Rows.Add()
                DataGridView3.Rows.Item(cnt).Cells(0).Value = fileName
                cnt += 1
            End If
        Next fileName
        If cnt = 0 Then MessageBox.Show("NO " & fext & " files in this work directory")
        Button8.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Button11.BackColor = System.Drawing.Color.Green
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = TextBox7.Text
        SaveFileDialog1.FileName = "_DXF_list" & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView3)
        Button11.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Button12.BackColor = System.Drawing.Color.Green
        DataGridView1.ClearSelection()

        Dim fileEntries As String() = Directory.GetFiles(TextBox8.Text)
        ' Process the list of files found in the directory.
        Dim fileName As String
        Dim ext As String
        For Each fileName In fileEntries
            ext = IO.Path.GetExtension(fileName)
            If ext = ".iam" Then
                Qbom(fileName)
            End If
        Next fileName
        Button12.BackColor = System.Drawing.Color.Transparent
    End Sub
    Private Sub PlotDXF()
        'http://modthemachine.typepad.com/my_weblog/2013/02/inventor-api-training-lesson-11.html
        ' Get the DXF translator Add-In.

        Dim oDocument As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        invApp.SilentOperation = vbTrue
        oDocument = CType(invApp.Documents.Open(filepath1, False), Document)

        Dim DXFAddIn As Inventor.TranslatorAddIn
        DXFAddIn = invApp.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")

        Dim oContext As Inventor.TranslationContext
        oContext = invApp.TransientObjects.CreateTranslationContext
        oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As Inventor.NameValueMap
        oOptions = invApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As Inventor.DataMedium
        oDataMedium = invApp.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DXFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then

            Dim strIniFile As String
            strIniFile = "M:\Engineering\PDFprinterVTK\DXF OUTPUTE.ini"

            ' Create the name-value that specifies the ini file to use.
            oOptions.Value("Export_Acad_IniFile") = strIniFile
        End If

        oDataMedium.FileName = "c:\temp\dxf_tst_1234.dxf"

        DXFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ExportSketchDXF()
    End Sub
    Private Sub PlotSTP()
        'https://forums.autodesk.com/t5/inventor-customization/vb-net-export-files-and-then-can-not-change-project/td-p/7404351
        'Dim oDocument As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        Dim oDrawDoc As DrawingDocument
        oDrawDoc = CType(invApp.Documents.Open(filepath1, False), Document)
        Dim oRefDoc As Document

        For Each oRefDoc In oDrawDoc.ReferencedDocuments
            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then

                Dim model As Inventor.PartDocument = invApp.Documents.Open("C:\Inventor_tst\Test_Copy.ipt", False)
                'model.SaveAs("c:\Inventor_tst/Test_Copy.stp", True)
                model.SaveAs("c:\Inventor_tst/Test_Copy.dxf", True)
                invApp.ActiveDocument.Close()
            ElseIf oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim sestava As Inventor.AssemblyDocument = invApp.Documents.Open("C:\Inventor_tst\Test_Copy.iam", False)
                sestava.SaveAs("C:\Inventor_tst\Test_Copy.stp", True)
                invApp.ActiveDocument.Close()
            End If
        Next oRefDoc
    End Sub

    Public Sub ExportSketchDXF()
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        Dim oDrawDoc As Inventor.Document
        oDrawDoc = invApp.Documents.Open(TextBox5.Text, False)

        Dim oSketch As PlanarSketch
        oSketch = invApp.ActiveDocument.ComponentDefinition.Sketches(1)

        Dim oDataIO As DataIO
        oDataIO = oSketch.DataIO

        oDataIO.WriteDataToFile("DXF", "C:\Inventor_tst\dxfout.dxf")
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Open_file(4)   'idw files
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        Dim oDoc As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        'invApp.SilentOperation = vbTrue
        oDoc = CType(invApp.Documents.Open(TextBox1.Text, False), Document)

        '--------- determine object type -------
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        If eDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
            MessageBox.Show("Please Select a IDW file ")
            Exit Sub
        End If

        ' Make sure a parts list is selected.
        Dim partList As Object
        partList = oDoc.SelectSet.Item(1)
        If Not (TypeOf partList Is PartsList) Then
            MessageBox.Show("A parts list must be selected.")
        End If

        'Expand legacy parts list to all levels
        Dim counter As Integer = 1
        Dim k As Integer

        While counter < partList.PartsListRows.Count
            For k = counter To partList.PartsListRows.Count
                Dim orow As PartsListRow
                orow = partList.PartsListRows.Item(k)
                counter = k
                While orow.Expandable And Not (orow.Expanded)
                    orow.Expanded = True
                    counter = counter + 1
                End While
            Next k
        End While

        Dim i, j As Integer

        For i = 1 To partList.PartsListColumns.Count
            'oWorkSheet.Cells(1, i).Value = partList.PartsListColumns.Item(i).Title
            ''MessageBox.Show(partList.PartsListColumns.Item(i).Title)
            For j = 1 To partList.PartsListRows.Count
                'oWorkSheet.Cells(j + 1, i).Value = partList.PartsListRows(j).Item(i).Value
                'MessageBox.Show(partList.PartsListRows(j).Item(i).Value)
            Next j
        Next i
    End Sub
End Class


