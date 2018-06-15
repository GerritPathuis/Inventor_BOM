Imports System.IO
Imports System
Imports System.String
Imports System.Runtime.InteropServices
'Browse to "C:\Programs Files\Autodesk\Inventor XXXX\Bin\Public Assemblies" select “autodesk.inventor.interop.dll”
Imports Inventor
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.ComponentModel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports Microsoft.Vbe.Interop
'==================================
'API samples
'https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2018/ENU/Inventor-API/files/SampleList-htm.html
'==================================
Public Class Form1
    Public filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Public filepath2 As String = "E:\Protmp\Procad"     'Work directory
    Public filepath3 As String = "c:\MyDir"
    Public filepath4 As String = "C:\Inventor_tst\Assembly1.idw"
    Public filepath5 As String = ""                     'Destination directory
    Public G1_row_cnt As Integer
    Public G2_row_cnt As Integer
    Public G5_row_cnt As Integer
    Public Const view_rows = 1000
    Public dxf_file_name(,) As String   '(old name, new name)  
    Dim Pro_user As String

    Public Structure Laserpart
        Public Proj As String
        Public Tmun As String
        Public Artnum As String
        Public Thick As String
        Public Materi As String
        Public Count As String
        Public actie As String
    End Structure
    Public kb As Laserpart

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DataGridView1.ColumnCount = 30
        DataGridView1.RowCount = view_rows
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

        DataGridView2.ColumnCount = 20
        DataGridView2.RowCount = view_rows   'was 20
        DataGridView2.Columns(0).HeaderText = "File"
        DataGridView2.Columns(1).HeaderText = "Assembly"
        DataGridView2.Columns(2).HeaderText = "IDW_Assy"
        DataGridView2.Columns(3).HeaderText = "A_Drwg nr"
        DataGridView2.Columns(4).HeaderText = "Pos"
        DataGridView2.Columns(5).HeaderText = "Qty"
        DataGridView2.Columns(6).HeaderText = "Artikel"
        DataGridView2.Columns(7).HeaderText = "Descrip"
        DataGridView2.Columns(8).HeaderText = "Length"
        DataGridView2.Columns(9).HeaderText = "Std"
        DataGridView2.Columns(10).HeaderText = "Mat"
        DataGridView2.Columns(11).HeaderText = "-"
        DataGridView2.Columns(12).HeaderText = "kg"
        DataGridView2.Columns(13).HeaderText = "Comment"
        DataGridView2.Columns(14).HeaderText = "M/B"

        DataGridView3.ColumnCount = 5
        DataGridView3.RowCount = view_rows
        DataGridView3.Columns(0).HeaderText = "File"
        DataGridView3.Columns(1).HeaderText = "D_no"
        DataGridView3.Columns(2).HeaderText = "A_no"

        DataGridView4.ColumnCount = 20
        DataGridView4.RowCount = view_rows
        DataGridView4.Columns(0).HeaderText = "File"
        DataGridView4.Columns(1).HeaderText = "D_no"
        DataGridView4.Columns(2).HeaderText = "A_no"

        DataGridView5.ColumnCount = 10
        DataGridView5.RowCount = view_rows    'was 20
        DataGridView5.Columns(0).HeaderText = "Artikel"
        DataGridView5.Columns(1).HeaderText = "Old dxf file name"
        DataGridView5.Columns(2).HeaderText = "New dxf file name"
        DataGridView5.Columns(3).HeaderText = "Material"
        DataGridView5.Columns(4).HeaderText = "Thick"
        DataGridView5.Columns(5).HeaderText = "Qty"
        DataGridView5.Columns(0).Width = 100
        DataGridView5.Columns(1).Width = 250
        DataGridView5.Columns(2).Width = 250

        Pro_user = System.Environment.UserName    'User name on the screen

        Dim number As Integer = 8
        Select Case Pro_user
            Case "GP"                               'Home
                filepath2 = "C:\Inventor_tst"
                filepath5 = "c:\Temp"
            Case "GerritP"                          'Work
                filepath2 = "C:\Inventor test files\KarelBakker2"
                filepath5 = "c:\temp"
            Case Else                               'Karel Bakker
                filepath2 = "E:\Protmp\Procad"
                filepath5 = "N:\CAD"
        End Select

        '======== Work directory's ==========
        TextBox5.Text = filepath2
        TextBox6.Text = filepath2
        TextBox7.Text = filepath2
        TextBox8.Text = filepath2
        TextBox9.Text = filepath2
        TextBox35.Text = filepath2

        '======== Destination directory's ==========
        TextBox34.Text = filepath5
        TextBox35.Text = filepath5

        Inventor_running()
    End Sub
    Private Sub Inventor_running()
        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            Label7.Visible = True
            Me.Text = "Inventor NOT running" & " (" & Pro_user & ")"
        Else
            Label7.Visible = False
            Me.Text = "Inventor BOM Extractor" & " (" & Pro_user & ")"
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Inventor_running()
        Open_file(1)   'ipt files
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Inventor_running()
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
               & "|Design element File (*.ide)|*.ide" _
               & "|Sheet matal File (*.dxf)|*.dxf",
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
        Inventor_running()
        Button3.BackColor = System.Drawing.Color.LightGreen
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
        oDoc = invApp.Documents.Open(fpath, False)

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
            If oPropSet.Count = 0 Then
                TextBox2.Text &= "The are NO 'Design Tracking' properties present in this file" & vbCrLf
            Else
                For j = 0 To design_track.Length - 1
                    Try
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = "+"
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = oPropSet.Item(design_track(j)).Value
                    Catch Ex As Exception
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = "?"
                        TextBox2.Text &= design_track(j) & " not found" & vbCrLf
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
                TextBox2.Text &= "The are NO 'Custom' properties present in this file" & vbCrLf
            Else
                For j = 0 To custom.Length - 1
                    Try
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = "+"
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = oPropSet.Item(custom(j)).Value
                    Catch Ex As Exception
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = "?"
                        TextBox2.Text &= "Custom property " & custom(j) & " not found" & vbCrLf
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
                TextBox2.Text &= "The are NO 'Inventor Summary Information' present in this file" & vbCrLf
            Else
                For j = 0 To summary.Length - 1
                    Try
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = "+"
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = oPropSet.Item(summary(j)).Value
                    Catch Ex As Exception
                        DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = "?"
                        TextBox2.Text &= "Inventor Summary " & summary(j) & " not found" & vbCrLf
                    End Try
                Next
            End If

        Next
        'Catch Ex As Exception
        'TextBox2.Text &= "No BOM in this IAM model"
        'Finally
        'End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Inventor_running()
        Button4.BackColor = System.Drawing.Color.LightGreen
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

        '-------- Cell_text to excel -------------
        Try
            For vert = 0 To dg.Rows.Count - 1
                For hor = 0 To dg.Columns.Count - 1
                    If Not dg.Rows.Item(vert).Cells(hor).Value Is Nothing Then
                        str = dg.Rows.Item(vert).Cells(hor).Value.ToString
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
        Inventor_running()
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
        oDoc = invApp.Documents.Open(filepath1, False)

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

    Private Function Isartikel(axx As String) As Boolean
        Dim is_artikel As Boolean

        If axx.Length = 0 Then
            Return False
        End If

        If axx.Length = 8 And axx.Substring(0, 1) = "A" Then
            is_artikel = True
        Else
            is_artikel = False
        End If
        Return is_artikel
    End Function
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
        Button9.BackColor = System.Drawing.Color.LightGreen
        Find_IDW()
        Button9.BackColor = System.Drawing.Color.Transparent
    End Sub
    Private Sub Find_IDW()
        Inventor_running()
        Button9.BackColor = System.Drawing.Color.LightGreen
        G2_row_cnt = -1  'Reset row counter

        'Select work directory
        Dim pathfile As String = TextBox6.Text

        If Directory.Exists(pathfile) Then
            Dim fileEntries As String() = Directory.GetFiles(pathfile)
            For Each fileName In fileEntries
                Increm_progressbar()
                Dim extension As String = IO.Path.GetExtension(fileName)
                If extension = ".idw" Then
                    Read_title_Block_idw(fileName)
                End If
            Next fileName
        Else
            MessageBox.Show(pathfile & " is not a valid file or directory.")
        End If
        Button9.BackColor = System.Drawing.Color.Transparent
    End Sub
    'Read IDW Title Block
    Public Sub Read_title_Block_idw(ByVal path As String)
        'http://adndevblog.typepad.com/manufacturing/2012/12/inventor-change-text-items-in-titleblockdefinition.html

        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.DrawingDocument

        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue

        oDoc = invApp.Documents.Open(path, False)

        'MessageBox.Show("Active document=" & oDoc.DisplayName)
        'MessageBox.Show("Active sheet=" & oDoc.ActiveSheet.Name)

        '=================================================================================
        'https://forums.autodesk.com/t5/inventor-customization/copy-titleblock-prompted-entries-to-custom-iproperty/td-p/7491136

        Dim oSheet As Sheet
        oSheet = oDoc.ActiveSheet
        Dim oTB1 As TitleBlock
        oTB1 = oSheet.TitleBlock
        Dim titleDef As TitleBlockDefinition
        titleDef = oTB1.Definition
        Dim oPrompt As Inventor.TextBox = Nothing
        Dim q_file As String = "-"  'File name
        Dim q_desc As String = "-"  'Description
        Dim q_A00 As String = "-"   'Assembly Artikel nummer
        Dim q_D00 As String = "-"   'Assembly Drawing nummer
        Dim q_mat As String = "-"

        ' Find the Prompted Entry called DESCRIPTION in the Title Block
        For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
            Increm_progressbar()
            q_file = IO.Path.GetFileName(path)          '=File naam (short)

            Select Case defText.Text
                Case "<DESCRIPTION>"        'Description
                    oPrompt = defText
                    q_desc = oTB1.GetResultText(oPrompt)
                Case "<ITEM_NR>"            '=A0000
                    oPrompt = defText
                    q_A00 = oTB1.GetResultText(oPrompt)
                Case "<DOC_NUMBER>"         '=D0000
                    oPrompt = defText
                    q_D00 = oTB1.GetResultText(oPrompt)
            End Select
        Next

        '============== Read The parts List=========================================
        ' Make sure a parts list is selected.
        Dim partList As Object
        '----------- does partlist exist ?------------
        If oDoc.ActiveSheet.PartsLists.Count > 0 Then
            partList = oDoc.ActiveSheet.PartsLists.Item(1)

            If (TypeOf partList Is PartsList) Then
                Dim counter As Integer = 1
                Dim str As String

                For jj = 1 To partList.PartsListRows.Count
                    G2_row_cnt += 1
                    DataGridView2.Rows.Add()
                    DataGridView2.Rows.Item(G2_row_cnt).Cells(0).Value = q_file
                    DataGridView2.Rows.Item(G2_row_cnt).Cells(1).Value = q_desc
                    DataGridView2.Rows.Item(G2_row_cnt).Cells(2).Value = q_A00
                    DataGridView2.Rows.Item(G2_row_cnt).Cells(3).Value = q_D00

                    For ii = 1 To partList.PartsListcolumns.Count 'WAS 4
                        str = partList.PartsListRows(jj).Item(ii).Value.ToString

                        '--------Check is this an artikel number-------
                        If (ii + 3) = 6 Then
                            If Isartikel(str) = False Then TextBox2.Text &= "IDW_drwg " & q_D00 & " BOM problem " & str & " is NOT an artikel number" & vbCrLf
                        End If
                        '-------- update datagrid---------
                        DataGridView2.Rows.Item(G2_row_cnt).Cells(ii + 3).Value = str
                    Next ii
                Next jj
            End If
        End If
        DataGridView2.AutoResizeColumns()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Inventor_running()
        FolderBrowserDialog1.SelectedPath = TextBox6.Text
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox5.Text = FolderBrowserDialog1.SelectedPath
            TextBox6.Text = FolderBrowserDialog1.SelectedPath
            TextBox7.Text = FolderBrowserDialog1.SelectedPath
            TextBox8.Text = FolderBrowserDialog1.SelectedPath
            TextBox9.Text = FolderBrowserDialog1.SelectedPath
            TextBox38.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Inventor_running()
        Button7.BackColor = System.Drawing.Color.LightGreen
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = TextBox6.Text
        SaveFileDialog1.FileName = "_Title_Blocks" & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView2)
        Button7.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Inventor_running()
        Button8.BackColor = System.Drawing.Color.LightGreen
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
                fext = ".ipt"   'sheet metal
            Case RadioButton9.Checked
                fext = ".idw"
            Case RadioButton10.Checked
                fext = ".*"
        End Select
        DataGridView3.Rows.Clear()
        DataGridView3.Columns(0).Width = 300
        DataGridView3.Columns(1).Width = 150

        If Directory.Exists(TextBox6.Text) Then
            Dim fileEntries As String() = Directory.GetFiles(TextBox6.Text)
            ' list files found in the directory.
            Dim fileName As String

            For Each fileName In fileEntries
                Increm_progressbar()
                extension = IO.Path.GetExtension(fileName)
                If String.Equals(extension, fext) Or RadioButton10.Checked Then
                    DataGridView3.Rows.Add()
                    DataGridView3.Rows.Item(cnt).Cells(0).Value = fileName
                    cnt += 1
                End If
                '=============== extra for Sheet metal ==============
                If RadioButton8.Checked And String.Equals(extension, ".ipt") Then
                    Dim invApp As Inventor.Application
                    invApp = Marshal.GetActiveObject("Inventor.Application")
                    invApp.SilentOperation = vbTrue
                    Dim oPartDoc As Inventor.Document
                    oPartDoc = invApp.Documents.Open(fileName, False)

                    Dim oFlatPattern As FlatPattern

                    If oPartDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        If oPartDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                            oFlatPattern = oPartDoc.ComponentDefinition.FlatPattern
                            If oFlatPattern Is Nothing Then
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Value = "NO Flat pattern"
                            Else
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Value = "Contains Flat pattern"
                            End If
                        End If
                    End If
                End If
            Next fileName
            If cnt = 0 Then MessageBox.Show("NO " & fext & " files in this work directory")
        Else
            MessageBox.Show(TextBox6.Text & " is not a valid directory.")
        End If
        DataGridView3.AutoResizeColumns()
        Button8.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Inventor_running()
        Button11.BackColor = System.Drawing.Color.LightGreen
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = TextBox7.Text
        SaveFileDialog1.FileName = "_DXF_list" & ".xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView3)
        Button11.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Inventor_running()
        Button12.BackColor = System.Drawing.Color.LightGreen
        DataGridView1.ClearSelection()

        Dim fileEntries As String() = Directory.GetFiles(TextBox8.Text)
        ' Process the list of files found in the directory.
        Dim fileName As String
        Dim ext As String
        For Each fileName In fileEntries
            Increm_progressbar()
            ext = IO.Path.GetExtension(fileName)
            If ext = ".iam" Then
                Qbom(fileName)
            End If
        Next fileName
        DataGridView1.AutoResizeColumns()
        Button12.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub PlotSTEP()
        'Export STEP of DXF Files
        'https://forums.autodesk.com/t5/inventor-customization/vb-net-export-files-and-then-can-not-change-project/td-p/7404351
        'Dim oDocument As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue

        Dim oDrawDoc As Inventor.DrawingDocument
        oDrawDoc = invApp.Documents.Open(filepath1, False)
        Dim oRefDoc As Document

        For Each oRefDoc In oDrawDoc.ReferencedDocuments
            Increm_progressbar()
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

    Public Sub ExportSketchDXF2(ByVal file_path As String)
        'https://forums.autodesk.com/t5/inventor-customization/flat-pattern-to-dxf/m-p/7033961#M71803
        'https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2018/ENU/Inventor-API/files/WriteFlatPatternAsDXF-Sample-htm.html
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue

        If IO.File.Exists(file_path) Then ' This pathfile is a file.
            Dim oPartDoc As Inventor.Document
            oPartDoc = invApp.Documents.Open(file_path, False)

            Dim oFlatPattern As FlatPattern

            'Pre-processing check: The Active document must be a Sheet metal Part with a flat pattern
            If oPartDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                TextBox2.Text &= "The Active document must be a 'Part'" & vbCrLf
                Exit Sub
            Else
                If oPartDoc.SubType <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                    'TextBox2.Text &= "The Active document must be a 'Sheet Metal Part'" & vbCrLf
                    Exit Sub
                Else
                    oFlatPattern = oPartDoc.ComponentDefinition.FlatPattern
                    If oFlatPattern Is Nothing Then
                        TextBox2.Text &= "IPT " & file_path & " does NOT contain a flat pattern" & vbCrLf
                        Exit Sub
                    End If
                End If
            End If

            'Processing:
            Dim oDataIO As DataIO
            oDataIO = oPartDoc.ComponentDefinition.DataIO

            'Dim strPartNum As String
            'strPartNum = oPartDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value
            'Dim strRev As String
            'strRev = oPartDoc.PropertySets("Inventor Summary Information").Item("Revision Number").Value


            '============= Check to see if the specified property exists.
            'http://modthemachine.typepad.com/my_weblog/2010/02/custom-iproperties.html
            'https://forums.windowssecrets.com/showthread.php/13785-Existing-CustomDocumentProperties-(VBA-Word)
            'https://www.office-forums.com/threads/how-can-i-check-to-see-if-a-customdocumentproperties-exists.1865599/

            Dim artikel As String = ""
            Dim customPropSet As PropertySet
            customPropSet = oPartDoc.PropertySets.Item("Inventor User Defined Properties")

            For Each prop In customPropSet
                If prop.Name = "ITEM_NR" Then
                    If prop.ToString.Length > 0 Then
                        artikel = oPartDoc.PropertySets("Inventor User Defined Properties").Item("ITEM_NR").Value
                    Else
                        artikel = "Axxx"
                    End If
                End If
            Next prop

            Dim oDXFfileNAME As String
            Dim strPath As String
            Dim sOut As String
            strPath = TextBox34.Text & "\"  'Must end with a "\"
            oDXFfileNAME = strPath & TextBox31.Text & "_" & TextBox33.Text & "_" & artikel & ".dxf"

            sOut = "FLAT PATTERN DXF?AcadVersion=R12"
            oDataIO.WriteDataToFile(sOut, oDXFfileNAME)
            DataGridView5.Rows.Item(G5_row_cnt).Cells(0).Value = artikel
            DataGridView5.Rows.Item(G5_row_cnt).Cells(1).Value = oDXFfileNAME

            'Plate thickness
            'Material sort
            'Part Count


            G5_row_cnt += 1
            TextBox2.Text &= "Dxf file " & oDXFfileNAME & " is written to work directory " & vbCrLf
        Else
            MessageBox.Show("File does noet exist")
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Inventor_running()
        Open_file(4)   'idw files
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Inventor_running()
        Read_idw_parts(TextBox1.Text)
    End Sub
    Private Sub Read_idw_parts(ByVal fpath As String)
        Dim oDoc As Inventor.DrawingDocument
        Dim invApp As Inventor.Application
        invApp = Marshal.GetActiveObject("Inventor.Application")

        invApp.SilentOperation = vbTrue
        oDoc = invApp.Documents.Open(fpath, False)  'Not visible

        '--------- determine object type -------
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        If eDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
            MessageBox.Show("Please Select a IDW file ")
            Exit Sub
        End If

        'http://beinginventive.typepad.com/files/ExportPartslistToExcel/ExportPartslistToExcel.txt
        ' Make sure a parts list is selected.
        Dim partList As Object
        If oDoc.ActiveSheet.PartsLists.Count > 0 Then
            partList = oDoc.ActiveSheet.PartsLists.Item(1)

            If (TypeOf partList Is PartsList) Then  'Parts-list exists ?
                'Expand legacy parts list to all levels
                Dim counter As Integer = 1
                Dim i, j As Integer

                '------ Column names ------------- 
                DataGridView4.Rows.Add()
                For i = 1 To partList.PartsListColumns.Count
                    DataGridView4.Columns(i - 1).HeaderText = partList.PartsListColumns.Item(i).Title.ToString
                Next

                '------ Column content ------------- 
                For j = 1 To partList.PartsListRows.Count
                    For i = 1 To partList.PartsListColumns.Count
                        DataGridView4.Rows.Item(j - 1).Cells(i - 1).Value = partList.PartsListRows(j).Item(i).Value.ToString
                    Next
                Next
            End If
        End If
        DataGridView4.AutoResizeColumns()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Inventor_running()
        DataGridView2.Rows.Clear()
        DataGridView2.RowCount = view_rows    'was 20
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Extract_dxf_from_IDW()
    End Sub
    Private Sub Extract_dxf_from_IDW()
        'Extract DXF file from the IDW
        Inventor_running()
        Button16.BackColor = System.Drawing.Color.LightGreen
        G5_row_cnt = 0

        If IO.Directory.Exists(TextBox5.Text) Then ' This pathfile is a file.
            Dim fileEntries As String() = Directory.GetFiles(TextBox5.Text)
            ' Process the list of files found in the directory.
            DataGridView1.ClearSelection()
            Dim fileName As String
            Dim ext As String
            For Each fileName In fileEntries
                Increm_progressbar()
                ext = IO.Path.GetExtension(fileName)
                If ext = ".ipt" Then
                    ExportSketchDXF2(fileName)
                End If
            Next fileName
            DataGridView1.AutoResizeColumns()
        Else
            MessageBox.Show("Directory does not exist")
        End If
        Button16.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Find_artikel()
        'Find the the artikel on the assembly drawing (IDW)
        'Print result in DataGridView5
        'DataGridView5 contains old and new file name

        Dim art As String
        Dim ask_once As Boolean = False

        For Each row In DataGridView5.Rows
            If row.Cells(0).Value <> Nothing Then
                art = row.Cells(0).Value.ToString
                Find_dwg_pos(DataGridView2, art)
                row.Cells(2).Value = kb.actie
                row.Cells(3).Value = kb.Materi
                row.Cells(4).Value = kb.Thick
                row.Cells(5).Value = kb.Count
            End If
        Next
        DataGridView5.AutoResizeColumns()
    End Sub
    Private Sub Rename_dxf()
        'Find the the artikel on the assembly drawing (IDW)
        'Print result in DataGridView5
        'DataGridView5 contains old and new file name

        Dim old_f, new_f, new_ff As String
        Dim delete_file As Boolean
        Dim ask_once As Boolean = False

        For Each row In DataGridView5.Rows
            Increm_progressbar()
            If row.Cells(0).Value <> Nothing Then
                'art = row.Cells(0).Value.ToString
                'row.Cells(2).Value = Find_dwg_pos(DataGridView2, art)
                old_f = row.Cells(1).Value.ToString
                new_f = row.Cells(2).Value.ToString

                new_ff = TextBox34.Text & "\" & new_f   'Full path required
                If IO.File.Exists(new_ff) And ask_once = False Then
                    delete_file = Question_replace_dxf_files()
                    ask_once = True
                End If

                If new_f.Length > 1 Then    'Make sure file name exist
                    If delete_file = True Then
                        IO.File.Delete(new_ff)
                        TextBox2.Text &= "Dxf file " & old_f & " deleted " & vbCrLf
                    End If

                    If Not IO.File.Exists(new_ff) Then
                        My.Computer.FileSystem.RenameFile(old_f, new_f)
                        TextBox2.Text &= "Dxf file " & new_f & " renamed " & vbCrLf
                    End If
                Else
                    TextBox2.Text &= "Dxf file " & old_f & "Failed NO new name !" & vbCrLf
                End If
            End If
        Next
        DataGridView5.AutoResizeColumns()
    End Sub

    Private Function Question_replace_dxf_files() As Boolean
        Dim result As DialogResult = DialogResult.No
        result = MessageBox.Show("Replace file with new one (YES for ALL)", "This dxf file already exist !!", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Return True
        Else
            Return False
        End If
    End Function
    'Find drawing name en postion of the artikel
    Private Sub Find_dwg_pos(ByVal dtg As DataGridView, ByVal Axxxxx As String) 'As String
        Dim actie As String = " "
        Dim found As Boolean = False
        Dim pos As Integer

        TextBox2.Text &= "Lookup drwg + pos for Artikel " & Axxxxx & " "

        For Each row As DataGridViewRow In dtg.Rows
            Increm_progressbar()
            If row.Cells.Item(6).Value = Axxxxx Then
                found = True
                actie = TextBox31.Text & "_"                    'Project
                actie &= TextBox33.Text & "_"                   'Tnumber
                actie &= row.Cells(3).Value.ToString() & "_"    'Drwg= 
                pos = row.Cells(4).Value                        'Pos= 
                actie &= pos.ToString("D2")
                If Not CheckBox3.Checked Then
                    actie &= "_" & row.Cells(6).Value.ToString() 'Artikel=  
                End If
                actie &= ".dxf"

                kb.Count = row.Cells(5).Value.ToString()        'Quantity
                kb.Thick = Isolate_thickness(row.Cells(7).Value.ToString())
                kb.Materi = row.Cells(10).Value.ToString()
                kb.actie = actie
                Exit For
            End If
        Next

        If found = True Then
            TextBox2.Text &= " found" & vbCrLf
        Else
            TextBox2.Text &= " NOT found !" & vbCrLf
        End If
    End Sub
    Private Function Isolate_thickness(str As String) As Integer
        Dim delta As Int16
        str = str.Substring(5, 3)

        Int16.TryParse(str, delta)

        Return delta.ToString
    End Function

    Private Sub Increm_progressbar()
        ProgressBar1.Value += 1
        If ProgressBar1.Value = 99 Then ProgressBar1.Value = 0
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        '==========WVB button======

        TextBox2.Clear()
        ProgressBar1.Visible = True
        Button19.BackColor = System.Drawing.Color.LightGreen
        DataGridView5.Rows.Clear()
        DataGridView5.RowCount = view_rows    'was 20

        TextBox2.Text &= "============= Find the IDW's =======================" & vbCrLf
        Button19.Text = "Find the IDW's..."
        Find_IDW()
        TextBox2.Text &= "============= Extract dxf from IDW ==================" & vbCrLf
        Button19.Text = "Extract dxf from idw's..."
        Extract_dxf_from_IDW()
        TextBox2.Text &= "============= Find artikel drwg + pos and rename ====" & vbCrLf
        Button19.Text = "Lookup artikel dwg and pos..."
        Find_artikel()
        TextBox2.Text &= "============= Rename dxf file ======================" & vbCrLf
        Button19.Text = "Rename dxg file..."
        Rename_dxf()
        Button19.Text = "WVB"
        ProgressBar1.Visible = False
        Button19.BackColor = System.Drawing.Color.Aqua
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Add Note then print IDW
        Inventor_running()
        Button6.BackColor = System.Drawing.Color.LightGreen

        If IO.Directory.Exists(TextBox5.Text) Then ' This pathfile is a file.
            Dim fileEntries As String() = Directory.GetFiles(TextBox5.Text)
            ' Process the list of files found in the directory.

            Dim fileName As String
            Dim ext As String
            For Each fileName In fileEntries
                Increm_progressbar()
                ext = IO.Path.GetExtension(fileName)
                If ext = ".idw" Then
                    Print_w_note(fileName)
                End If
            Next fileName
        Else
            MessageBox.Show("Directory does not exist")
        End If
        Button6.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Print_w_note(fileName As String)
        'http://adndevblog.typepad.com/manufacturing/2013/01/adding-general-note-idw-and-obtain-its-height.html
        'http://modthemachine.typepad.com/my_weblog/wayne/page/2/
        'https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2018/ENU/Inventor-API/files/GeneralNote-Sample-htm.html

        Dim sNote As String = "<StyleOverride Font='Arial' FontSize='0.5' Bold='True'>" + "Sample note" + "</StyleOverride>"
        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.Document
        Dim x, y As Integer
        Dim f_size As Double
        Dim dest As String
        Dim q_file As String = "-"  'File name

        invApp = Marshal.GetActiveObject("Inventor.Application")
        invApp.SilentOperation = vbTrue
        oDoc = invApp.Documents.Open(fileName, False)

        ' Set a reference to the active sheet.
        Dim oActiveSheet As Sheet
        oActiveSheet = oDoc.ActiveSheet

        ' Set a reference to the GeneralNotes object
        Dim oGeneralNotes As GeneralNotes
        oGeneralNotes = oActiveSheet.DrawingNotes.GeneralNotes

        Dim oTG As TransientGeometry
        oTG = invApp.TransientGeometry

        ' Create text with simple string as input. Since this doesn't use
        ' any text overrides, it will default to the active text style.
        Dim sText As String = TextBox36.Text

        Dim oGeneralNote As GeneralNote
        x = NumericUpDown1.Value
        y = NumericUpDown2.Value
        f_size = NumericUpDown3.Value
        oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x, y), "-")
        oGeneralNote.FormattedText = "<StyleOverride FontSize = '" & f_size.ToString & "'>" & TextBox36.Text & "</StyleOverride>"

        'Save the document
        dest = TextBox35.Text & "\" & IO.Path.GetFileName(fileName)
        oDoc.SaveAs(dest, True)                 'WORKS
        'oDoc.Save()                            'Works
        TextBox2.Text &= "Drawing Note added to " & dest & vbCrLf
    End Sub

End Class


