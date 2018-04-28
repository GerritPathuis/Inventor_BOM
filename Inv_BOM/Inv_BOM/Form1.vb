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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView2.ColumnCount = 5
        DataGridView2.RowCount = 1000

        DataGridView2.Columns(0).HeaderText = "File"
        DataGridView2.Columns(1).HeaderText = "D_no"
        DataGridView2.Columns(2).HeaderText = "A_no"
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Open_file(1)   'ipt files
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Open_file(2)    'iam files
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Open_file(5)    'dwg files
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
        Qbom()
        Button3.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Qbom()
        Dim information As System.IO.FileInfo
        Dim filen As String

        DataGridView1.ColumnCount = 21
        DataGridView1.ColumnHeadersVisible = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        '------- get file info -----------
        information = My.Computer.FileSystem.GetFileInfo(filepath1)
        filen = information.Name

        Dim oDoc As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        invApp.SilentOperation = vbTrue
        oDoc = CType(invApp.Documents.Open(filepath1, False), Document)

        '--------- determine object type -------
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        If eDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Please Select a IAM file ")
            Exit Sub
        End If

        '-------------READ TITLE BLOCK----------------------------------------
        ' ---- Note: there is no title block in an IAmmodel file -------------

        '---------- Read BOM --------------------------
        Try
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
            Dim i, j, r As Integer

            DataGridView1.Columns(0).HeaderText = "File"
            DataGridView1.Columns(1).HeaderText = "D_no"
            DataGridView1.Columns(2).HeaderText = "A_no"
            DataGridView1.Columns(3).HeaderText = "Item "
            DataGridView1.Columns(4).HeaderText = "Qty"
            DataGridView1.Columns(5).HeaderText = "Part"
            DataGridView1.Columns(6).HeaderText = "Desc"
            DataGridView1.Columns(7).HeaderText = "Stock"

            DataGridView1.Columns(8).HeaderText = "DOC_NUMBER"
            DataGridView1.Columns(9).HeaderText = "ITEM_NR"
            DataGridView1.Columns(10).HeaderText = "DOC_STATUS"
            DataGridView1.Columns(11).HeaderText = "DOC_REV"
            DataGridView1.Columns(12).HeaderText = "PART_MATERIAL"
            DataGridView1.Columns(13).HeaderText = "IT_TP"
            DataGridView1.Columns(14).HeaderText = "LENGTH"
            DataGridView1.Columns(15).HeaderText = "Part Icon"

            DataGridView1.Columns(16).HeaderText = "Title"
            DataGridView1.Columns(17).HeaderText = "Subject"
            DataGridView1.Columns(18).HeaderText = "Author"
            DataGridView1.Columns(19).HeaderText = "Comments"

            For i = 1 To oBOMView.BOMRows.Count
                r = i - 1

                '================= Design Tracking Properties ==========================
                oRow = oBOMView.BOMRows.Item(i)
                oCompDef = oRow.ComponentDefinitions.Item(1)

                oPropSet = oCompDef.Document.PropertySets.Item("Design Tracking Properties")
                DataGridView1.Rows.Add()

                DataGridView1.Rows.Item(r).Cells(0).Value = filen

                DataGridView1.Rows.Item(r).Cells(1).Value = TextBox3.Text
                DataGridView1.Rows.Item(r).Cells(2).Value = TextBox4.Text

                DataGridView1.Rows.Item(r).Cells(3).Value = oRow.ItemNumber
                DataGridView1.Rows.Item(r).Cells(4).Value = oRow.ItemQuantity

                Dim design_track() As String =
                {"Part Number",
                "Description",
                "Stock Number",
                "Part Icon"}
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Design Tracking' properties present in this file")
                Else
                    For j = 0 To design_track.Length - 1
                        Try
                            DataGridView1.Rows.Item(r).Cells(j + 5).Value = oPropSet.Item(design_track(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(r).Cells(j + 5).Value = "?"
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
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Custom' properties present in this file")
                Else
                    For j = 0 To custom.Length - 1
                        Try
                            DataGridView1.Rows.Item(r).Cells(j + 8).Value = oPropSet.Item(custom(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(r).Cells(j + 8).Value = "?"
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
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Inventor Summary Information' present in this file")
                Else
                    For j = 0 To summary.Length - 1
                        Try
                            DataGridView1.Rows.Item(r).Cells(j + 16).Value = oPropSet.Item(summary(j)).Value
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(r).Cells(j + 16).Value = "?"
                            If Not CheckBox1.Checked Then MessageBox.Show(summary(j) & " not found")
                        End Try
                    Next
                End If

            Next
        Catch Ex As Exception
            MessageBox.Show("No BOM in this IAM model")
        Finally
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.BackColor = System.Drawing.Color.Green
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = filepath3
        SaveFileDialog1.FileName = "_BOM" & "_" & TextBox3.Text & "_" & TextBox4.Text & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel()
        Button4.BackColor = System.Drawing.Color.Transparent
    End Sub
    Private Sub Write_excel()
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

        '-------- Header text to excel -------------tr
        For hor = 0 To 19
            str = DataGridView1.Columns(hor).HeaderText
            xlWorksheet.Cells(1, hor + 1) = str
        Next

        'TextBox2.Text &= "Rows...." & DataGridView1.Rows.Count
        'TextBox2.Text &= "Columns...." & DataGridView1.Columns.Count

        '-------- Cell_text to excel -------------
        Try
            For vert = 0 To DataGridView1.Rows.Count - 2
                For hor = 0 To 19
                    str = DataGridView1.Rows.Item(vert).Cells(hor).Value.ToString
                    If str = Nothing Then
                        str = "-"
                    End If
                    'TextBox2.Text &= "hor=" & hor.ToString & " vert=" & vert.ToString & " str= " & str & vbCrLf
                    xlWorksheet.Cells(vert + 2, hor + 1) = str
                Next
            Next
            ' MessageBox.Show("save....")
            fname = SaveFileDialog1.FileName
            xlWorkBook.SaveAs(fname, FileFormat:=XlFileFormat.xlWorkbookNormal)

            xlApp.Quit()

            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorksheet)
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
                    If defText.Text = "<TITLE>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(1).Value = "Title= " & oTB1.GetResultText(oPrompt)
                    End If
                    If defText.Text = "<PART NUMBER>" Then
                        oPrompt = defText
                        DataGridView2.Rows.Item(title_counter).Cells(2).Value = "A_no= " & oTB1.GetResultText(oPrompt)
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
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox6.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

End Class


