Imports System.IO
Imports System.Runtime.InteropServices
'Browse to "C:\Programs Files\Autodesk\Inventor XXXX\Bin\Public Assemblies" select “autodesk.inventor.interop.dll”
Imports Inventor
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

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
    Dim invApp As Inventor.Application = Nothing
    'Public Property vbColor As Object

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
        DataGridView3.Columns(3).HeaderText = "Flat pattern"

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
                filepath5 = "c:\Temp"
            Case Else                               'Karel Bakker
                filepath2 = "E:\Protmp\Procad"
                filepath5 = "c:\Temp"
        End Select

        Try 'Create directory 
            If (Not System.IO.Directory.Exists("c:\Temp")) Then System.IO.Directory.CreateDirectory("c:\Temp")
        Catch ex As Exception
        End Try

        '======== Work directory's ==========
        TextBox5.Text = filepath2
        TextBox6.Text = filepath2
        TextBox7.Text = filepath2
        TextBox8.Text = filepath2
        TextBox9.Text = filepath2
        TextBox35.Text = filepath2

        TextBox31.Text = "P" & DateTime.Now.ToString("yy") & ".10"
        '======== Destination directory's ==========
        TextBox34.Text = filepath5
        TextBox35.Text = filepath5

        TextBox39.Text = "All idw drawings in the work-directory are copied as dwg or dxf format" & vbCrLf & vbCrLf
        TextBox39.Text &= "Save a Inventor drawing (idw format) to dwg format using" & vbCrLf
        TextBox39.Text &= "additional options in 'c:\Temp\dwgout2.ini' file" & vbCrLf & vbCrLf
        TextBox39.Text &= "You can export an Inventor drawing to AutoCAD dwg or dxf format" & vbCrLf
        TextBox39.Text &= "using the DWG and DXF Translator AddIns" & vbCrLf
        TextBox39.Text &= "The translators use ini file (configuration file) " & vbCrLf
        TextBox39.Text &= "to set additional options for the AutoCAD dwg Or dxf that will be created." & vbCrLf
        TextBox39.Text &= "You can create the ini file Using the Options dialog which can be reached " & vbCrLf
        TextBox39.Text &= "From the 'SaveCopyAs' dialog, when the *.dwg file format are selected."
        Inventor_running()
    End Sub
    Private Sub Inventor_running()
        '-------- inventor must be running----
        Me.Text = "Inventor BOM Extractor" & " (" & Pro_user & ") 05-04-2019"

        Try
            invApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        Catch 'Inventor not started
            System.Windows.Forms.MessageBox.Show("Start an Inventor session")
            Exit Sub
        End Try
        Label7.Visible = False
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
               .InitialDirectory = "c\Inventor test files\",
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
                MessageBox.Show("Cannot read file from disk. Original error " & Ex.Message)
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
        DataGridView1.Sort(DataGridView1.Columns(10), System.ComponentModel.ListSortDirection.Descending)
    End Sub
    'Read assembly make part summary from IAM (Autodesk Inventor assembly file) 
    Private Sub Qbom(ByVal fpath As String)
        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.AssemblyDocument
        Dim oBOM As BOM
        Dim oBOMView As BOMView
        Dim oBOMRow As BOMRow
        Dim oCompDef As ComponentDefinition
        Dim eDocumentType As DocumentTypeEnum
        'Dim odef As AssemblyComponentDefinition
        Dim oPropSets As PropertySets
        Dim oPropSet As PropertySet
        Dim information As System.IO.FileInfo
        Dim filen As String
        Dim doc_status As String
        Dim i, j As Integer


        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor Is Not running")
            Exit Sub
        End If


        ProgressBar3.Visible = True
        '------- get file info -----------
        information = My.Computer.FileSystem.GetFileInfo(fpath)
        filen = information.Name
        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)

        invApp.SilentOperation = True
        oDoc = CType(invApp.Documents.Open(fpath, False), AssemblyDocument)


        '--------- determine object type ---------
        '------ jump out when not a assembly !!-----
        eDocumentType = oDoc.DocumentType

        If eDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Please Select a IAM file ")
            Exit Sub
        End If

        '--------- test section -------
        'Dim objDrawDoc As DrawingDocument = CType(oDoc.ActiveDocument, AssemblyDocument)
        'Dim colTitleBlkDefs As TitleBlockDefinitions = objDrawDoc.TitleBlockDefinitions

        'If colTitleBlkDefs.Count = Nothing Then
        '    MessageBox.Show("NO Titleblock resent ")
        '    Exit Sub
        'End If
        '-------- end test section

        '-------------READ TITLE BLOCK----------------------------------------
        '---- Note: there is no title block in an IAM model file -------------
        '---------- Read BOM, in IAM model file --------------------------
        'Loopt vast als er een base model is !!!!!!!!!!!!
        Try
            oBOM = oDoc.ComponentDefinition.BOM
            oBOM.StructuredViewFirstLevelOnly = True
            oBOM.PartsOnlyViewEnabled = True
            oBOMView = oBOM.BOMViews.Item("Parts Only")
            '-------------------------

            For i = 1 To oBOMView.BOMRows.Count
                G1_row_cnt += 1
                Increm_progressbar()

                '================= Design Tracking Properties ==========================
                oBOMRow = oBOMView.BOMRows(i)
                oCompDef = oBOMRow.ComponentDefinitions(1)

                oPropSets = oDoc.PropertySets
                oPropSet = oPropSets.Item("Design Tracking Properties")

                DataGridView1.Rows.Add()
                DataGridView1.Rows.Item(G1_row_cnt).Cells(0).Value = filen
                DataGridView1.Rows.Item(G1_row_cnt).Cells(1).Value = oBOMRow.ItemNumber
                DataGridView1.Rows.Item(G1_row_cnt).Cells(2).Value = oBOMRow.ItemQuantity

                Dim design_track() As String =
                {"Part Number",
                "Description",
                "Stock Number",
                "Part Icon"}
                If oPropSet.Count < 1 Then
                    TextBox2.Text &= "The are NO 'Design Tracking' properties present in this file" & vbCrLf
                Else
                    For j = 0 To design_track.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = oPropSet.Item(design_track(j)).Value.ToString
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 3).Value = "?"
                            TextBox2.Text &= fpath & ", " & design_track(j) & " not found" & vbCrLf
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

                oPropSet = oPropSets.Item("Inventor User Defined Properties")
                If oPropSet.Count = 0 Then
                    TextBox2.Text &= "The are NO 'Custom' properties present in this file" & vbCrLf
                Else
                    For j = 0 To custom.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = oPropSet.Item(custom(j)).Value.ToString

                            '--- check PDM status of document ---
                            doc_status = oPropSet.Item(custom(j)).Value.ToString
                            If CBool(CInt(j = 2)) And String.Equals(doc_status, "In work") Then
                                TextBox2.Text &= fpath & " document is NOT Released !!" & vbCrLf
                                DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Style.BackColor = System.Drawing.Color.Red
                                DataGridView1.Rows.Item(G1_row_cnt).Cells(0).Style.BackColor = System.Drawing.Color.Red
                            End If
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 6).Value = "?"
                            'TextBox2.Text &= fpath & ", " & "Custom property " & custom(j) & " not found" & vbCrLf
                        End Try
                    Next
                End If

                '========== Inventor Summary Information ===============
                Dim summary() As String =
                {"Title",
                "Subject",
                "Author",
                "Comments"}
                oPropSet = oPropSets.Item("Inventor Summary Information")
                If oPropSet.Count = 0 Then
                    TextBox2.Text &= "The are NO 'Inventor Summary Information' present in this file" & vbCrLf
                Else
                    For j = 0 To summary.Length - 1
                        Try
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = oPropSet.Item(summary(j)).Value.ToString
                        Catch Ex As Exception
                            DataGridView1.Rows.Item(G1_row_cnt).Cells(j + 14).Value = "?"
                            'TextBox2.Text &= fpath & ", " & "Inventor Summary " & summary(j) & " not found" & vbCrLf
                        End Try
                    Next
                End If
            Next
        Catch Ex As Exception
            Form2.Show()
            'TextBox2.Text &= fpath & ", " & "No BOM in this IAM model " & vbCrLf
            Form2.TextBox1.Text &= fpath & ", " & "No BOM in this IAM model " & vbCrLf
            Return
        Finally
        End Try
        ProgressBar3.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Inventor_running()
        Button4.BackColor = System.Drawing.Color.LightGreen
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = filepath3
        SaveFileDialog1.FileName = "_IAM_BOM_List.xls"
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

        xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
        xlWorkBook = xlApp.Workbooks.Add(Type.Missing)
        xlWorksheet = CType(xlWorkBook.Worksheets(1), Worksheet)

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

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)
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
        Dim cnt As String

        Button9.BackColor = System.Drawing.Color.LightGreen
        ProgressBar2.Visible = True
        cnt = CType(Find_IDW(), String)
        If CInt(cnt) = 0 Then
            MessageBox.Show("WARNING NO IDW files found in the Work directory !!")
        Else
            TextBox2.Text &= cnt & " IDW files found in the Work directory"
        End If

        Button9.BackColor = System.Drawing.Color.Transparent
        ProgressBar2.Visible = False
    End Sub
    Private Function Find_IDW() As Integer
        Dim idw_counter As Integer = 0
        Inventor_running()
        Button9.BackColor = System.Drawing.Color.LightGreen
        G2_row_cnt = -1  'Reset row counter

        'Select work directory
        Dim pathfile As String = TextBox6.Text

        If Directory.Exists(pathfile) Then
            Dim fileEntries As String() = Directory.GetFiles(pathfile)
            For Each fileName In fileEntries
                TextBox42.Text = fileName
                Increm_progressbar()
                Dim extension As String = IO.Path.GetExtension(fileName)
                If extension = ".idw" Then
                    Read_title_Block_idw(fileName)
                    idw_counter += 1
                    Label27.Text = "IDW " & idw_counter.ToString
                    Label24.Text = "IDW " & idw_counter.ToString
                End If
            Next fileName
        Else
            MessageBox.Show(pathfile & " is not a valid file or directory.")
        End If
        Button9.BackColor = System.Drawing.Color.Transparent
        TextBox42.Text = " "
        Return (idw_counter)
    End Function
    'Read IDW Title Block
    Public Sub Read_title_Block_idw(ByVal path As String)
        'http://adndevblog.typepad.com/manufacturing/2012/12/inventor-change-text-items-in-titleblockdefinition.html

        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.DrawingDocument
        Dim partList As Inventor.PartsList
        Dim oSheet As Sheet
        Dim oTB1 As TitleBlock
        Dim titleDef As TitleBlockDefinition

        Dim oPrompt As Inventor.TextBox = Nothing
        Dim q_file As String = "-"  'File name
        Dim q_desc As String = "-"  'Description
        Dim q_A00 As String = "-"   'Assembly Artikel nummer
        Dim q_D00 As String = "-"   'Assembly Drawing nummer
        Dim q_mat As String = "-"

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)
        oDoc = CType(invApp.Documents.Open(path, False), DrawingDocument)

        'MessageBox.Show("Active document=" & oDoc.DisplayName)
        'MessageBox.Show("Active sheet=" & oDoc.ActiveSheet.Name)
        'MessageBox.Show("Active document type= " & oDoc.DocumentType.ToString)

        '=================================================================================
        'https://forums.autodesk.com/t5/inventor-customization/copy-titleblock-prompted-entries-to-custom-iproperty/td-p/7491136
        oSheet = oDoc.ActiveSheet
        oTB1 = oSheet.TitleBlock
        titleDef = oTB1.Definition

        ' Find the Prompted Entry called DESCRIPTION in the Title Block
        For Each defText As Inventor.TextBox In titleDef.Sketch.TextBoxes
            If defText Is Nothing Then Exit Sub
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
        '----------- does partlist exist ?------------

        If oDoc.ActiveSheet.PartsLists.Count > 0 Then
                partList = oDoc.ActiveSheet.PartsLists.Item(1)

                If (TypeOf partList Is PartsList) Then
                    Dim counter As Integer = 1
                    Dim str As String
                Try
                    For jj = 1 To partList.PartsListRows.Count
                        G2_row_cnt += 1
                        DataGridView2.Rows.Add()
                        DataGridView2.Rows.Item(G2_row_cnt).Cells(0).Value = q_file
                        DataGridView2.Rows.Item(G2_row_cnt).Cells(1).Value = q_desc
                        DataGridView2.Rows.Item(G2_row_cnt).Cells(2).Value = q_A00
                        DataGridView2.Rows.Item(G2_row_cnt).Cells(3).Value = q_D00

                        For ii = 1 To partList.PartsListColumns.Count
                            str = partList.PartsListRows(jj).Item(ii).Value.ToString
                            '--------Check is this an artikel number-------
                            If ((ii + 3) = 6) Then
                                If Isartikel(str) = False Then TextBox2.Text &= "IDW_drwg " & q_D00 & " BOM problem " & str & " is NOT an artikel number" & vbCrLf
                            End If
                            '-------- update datagrid---------
                            DataGridView2.Rows.Item(G2_row_cnt).Cells(ii + 3).Value = str
                        Next ii
                    Next jj
                Catch ex As Exception
                    MessageBox.Show("Problem Read_title_Block_idw, " & ex.Message)
                End Try
            End If
            End If

        Remove_empty_rows(DataGridView2)
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
        SaveFileDialog1.FileName = "IDW_BOM_List.xls"
        SaveFileDialog1.ShowDialog()
        Write_excel(DataGridView2)
        Button7.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        List_files()
    End Sub
    Private Sub List_files()
        Dim invApp As Inventor.Application
        Dim oPartDoc As Inventor.PartDocument
        Dim oFlatPattern As FlatPattern
        Dim fileEntries As String() = Directory.GetFiles(TextBox6.Text)

        Dim cnt As Integer = 0   'Reset counter
        Dim fext As String = ".dxf"
        Dim extension As String
        Dim fileName As String

        Inventor_running()
        Button8.BackColor = System.Drawing.Color.LightGreen
        DataGridView3.Rows.Clear()
        DataGridView3.Columns(0).Width = 450
        DataGridView3.Columns(1).Width = 60
        DataGridView3.Columns(2).Width = 60

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


        If Directory.Exists(TextBox6.Text) Then
            ' list files found in the directory.
            invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
            invApp.SilentOperation = CBool(vbTrue)

            For Each fileName In fileEntries
                Increm_progressbar()
                extension = IO.Path.GetExtension(fileName)
                If String.Equals(extension, fext) Or RadioButton10.Checked Then
                    DataGridView3.Rows.Add()
                    DataGridView3.Rows.Item(cnt).Cells(0).Value = fileName
                    cnt += 1
                End If
                '=============== extra for Sheet metal parts ==============
                If RadioButton8.Checked And String.Equals(extension, ".ipt") Then
                    oPartDoc = CType(invApp.Documents.Open(fileName, False), PartDocument)

                    If oPartDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        If oPartDoc.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                            oFlatPattern = oPartDoc.ComponentDefinition.FlatPattern

                            If oFlatPattern Is Nothing Then
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Value = "Sheet metal, NO Flat pattern"
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Style.BackColor = System.Drawing.Color.Red
                                DataGridView3.Rows.Item(cnt - 1).Cells(0).Style.BackColor = System.Drawing.Color.Red
                            Else
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Value = "Contains Flat pattern"
                                DataGridView3.Rows.Item(cnt - 1).Cells(3).Style.BackColor = System.Drawing.Color.White
                                DataGridView3.Rows.Item(cnt - 1).Cells(0).Style.BackColor = System.Drawing.Color.White
                            End If
                        End If
                    Else
                        DataGridView3.Rows.Item(cnt - 1).Cells(3).Value = "Part"
                    End If
                End If
            Next fileName
            If cnt = 0 Then MessageBox.Show("NO " & fext & " files in this work directory")
        Else
            MessageBox.Show(TextBox6.Text & " is not a valid directory.")
        End If
        Remove_empty_rows(DataGridView3)
        DataGridView3.Sort(DataGridView3.Columns(3), System.ComponentModel.ListSortDirection.Descending)
        DataGridView3.AutoResizeColumns()
        Button8.BackColor = System.Drawing.Color.Transparent
    End Sub
    Private Sub Remove_empty_rows(grid As DataGridView)
        For r As Integer = grid.Rows.Count - 2 To 10 Step -1
            Dim empty As Boolean = True
            For Each cell As DataGridViewCell In grid.Rows(r).Cells
                If Not IsNothing(cell.Value) Then
                    empty = False
                    Exit For
                End If
            Next
            If empty Then grid.Rows.RemoveAt(r)
        Next
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
        Dim iam_cnt As Integer = 0
        Inventor_running()
        Button12.BackColor = System.Drawing.Color.LightGreen
        DataGridView1.ClearSelection()

        Dim fileEntries As String() = Directory.GetFiles(TextBox8.Text)
        ' Process the list of files found in the directory.
        Dim fileName As String
        Dim ext As String
        For Each fileName In fileEntries
            Increm_progressbar()
            TextBox43.Text = fileName
            ext = IO.Path.GetExtension(fileName)
            If ext = ".iam" Then
                iam_cnt += 1
                Label12.Text = "IAM " & iam_cnt.ToString
                Qbom(fileName)
            End If
        Next fileName

        TextBox43.Text = " "
        Remove_empty_rows(DataGridView1)
        DataGridView1.AutoResizeColumns()
        Button12.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Plot_STEPorDXF()
        'Export STEP or DXF Files
        'https://forums.autodesk.com/t5/inventor-customization/vb-net-export-files-and-then-can-not-change-project/td-p/7404351
        'Dim oDocument As Inventor.Document
        Dim invApp As Inventor.Application
        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)

        Dim oDrawDoc As Inventor.DrawingDocument
        oDrawDoc = CType(invApp.Documents.Open(filepath1, False), DrawingDocument)
        Dim oRefDoc As Document

        For Each oRefDoc In oDrawDoc.ReferencedDocuments
            Increm_progressbar()
            If oRefDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                Dim model As Inventor.PartDocument = CType(invApp.Documents.Open("C:\Inventor_tst\Test_Copy.ipt", False), PartDocument)
                'model.SaveAs("c:\Inventor_tst/Test_Copy.stp", True)
                model.SaveAs("c:\Inventor_tst/Test_Copy.dxf", True)
                invApp.ActiveDocument.Close()
            ElseIf oRefDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim sestava As Inventor.AssemblyDocument = CType(invApp.Documents.Open("C:\Inventor_tst\Test_Copy.iam", False), AssemblyDocument)
                sestava.SaveAs("C:\Inventor_tst\Test_Copy.stp", True)
                invApp.ActiveDocument.Close()
            End If
        Next oRefDoc
    End Sub

    Public Sub ExportSketchDXF2(ByVal file_path As String)
        'https://forums.autodesk.com/t5/inventor-customization/flat-pattern-to-dxf/m-p/7033961#M71803
        'https://knowledge.autodesk.com/search-result/caas/CloudHelp/cloudhelp/2018/ENU/Inventor-API/files/WriteFlatPatternAsDXF-Sample-htm.html
        Dim invApp As Inventor.Application
        Dim oPartDoc As Inventor.partDocument
        Dim oFlatPattern As FlatPattern
        Dim oDataIO As DataIO
        Dim customPropSet As PropertySet
        'Dim prop As PropertySet
        Dim oDXF_fileNAME, oDWG_FfileNAME As String
        Dim strPath As String
        Dim sOut As String
        Dim artikel As String = ""

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)
        If IO.File.Exists(file_path) Then ' This pathfile is a file.
            oPartDoc = invApp.Documents.Open(file_path, False)

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
                    'oFlatPattern = CType(oPartDoc.ComponentDefinition, FlatPattern)
                    If oFlatPattern Is Nothing Then
                        If Not CheckBox1.Checked Then TextBox2.Text &= "IPT sheet metal part " & file_path & " does NOT contain a flat pattern" & vbCrLf
                        Exit Sub
                    End If
                End If
            End If

            'Processing:
            oDataIO = oPartDoc.ComponentDefinition.DataIO

            'Dim strPartNum As String
            'strPartNum = oPartDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value
            'Dim strRev As String
            'strRev = oPartDoc.PropertySets("Inventor Summary Information").Item("Revision Number").Value

            '============= Check to see if the specified property exists.
            'http://modthemachine.typepad.com/my_weblog/2010/02/custom-iproperties.html
            'https://forums.windowssecrets.com/showthread.php/13785-Existing-CustomDocumentProperties-(VBA-Word)
            'https://www.office-forums.com/threads/how-can-i-check-to-see-if-a-customdocumentproperties-exists.1865599/


            customPropSet = oPartDoc.PropertySets.Item("Inventor User Defined Properties")
            Try
                For Each prop In customPropSet
                    If prop.Name = "ITEM_NR" Then
                        If prop.ToString.Length > 0 Then
                            artikel = CType(oPartDoc.PropertySets("Inventor User Defined Properties").Item("ITEM_NR").Value, String)
                        Else
                            artikel = "Axxx"
                        End If
                    End If
                Next prop
            Catch ex As Exception
                MessageBox.Show("Problem with ITEM_NR, " & ex.Message)
            End Try

            strPath = TextBox34.Text & "\"  'Must end with a "\"
                oDXF_fileNAME = strPath & TextBox31.Text & "_" & TextBox33.Text & "_" & artikel & ".dxf"
                oDWG_FfileNAME = strPath & TextBox31.Text & "_" & TextBox33.Text & "_" & artikel & ".dwg"

                'Write dxf file
                sOut = "FLAT PATTERN DXF?AcadVersion=R12"
                oDataIO.WriteDataToFile(sOut, oDXF_fileNAME)

                'Write dwg file
                If CheckBox2.Checked Then
                    sOut = "FLAT PATTERN DWG?AcadVersion=2000"
                    oDataIO.WriteDataToFile(sOut, oDWG_FfileNAME) 'Write dwg
                End If

                DataGridView5.Rows.Item(G5_row_cnt).Cells(0).Value = artikel
                DataGridView5.Rows.Item(G5_row_cnt).Cells(1).Value = oDXF_fileNAME

                'Plate thickness
                'Material sort
                'Part Count

                G5_row_cnt += 1
                If Not CheckBox1.Checked Then TextBox2.Text &= "Dxf file " & oDXF_fileNAME & " is written to work directory " & vbCrLf
            Else
                MessageBox.Show("DXF File does noet exist")
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
        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)

        invApp.SilentOperation = CBool(vbTrue)
        oDoc = CType(invApp.Documents.Open(fpath, False), DrawingDocument)  'Not visible

        '--------- determine object type -------
        Dim eDocumentType As DocumentTypeEnum = oDoc.DocumentType
        If eDocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
            MessageBox.Show("Please Select a IDW file ")
            Exit Sub
        End If

        'http://beinginventive.typepad.com/files/ExportPartslistToExcel/ExportPartslistToExcel.txt
        ' Make sure a parts list is selected.
        Dim partList As PartsList
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
        Dim cnt As Integer
        cnt = Extract_dxf_from_IPT()
        If cnt = 0 Then MessageBox.Show("WARNING NO Dxf's extracted from IDW files")
    End Sub
    Private Function Extract_dxf_from_IPT() As Integer
        'Extract DXF file from the IDW is name contains Plate
        Dim ipt_counter As Integer = 0
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
                TextBox42.Text = fileName
                Increm_progressbar()
                ext = IO.Path.GetExtension(fileName)
                'If ext = ".ipt" And fileName.ToUpper.Contains("PLATE") Then
                If ext = ".ipt" Then
                    ExportSketchDXF2(fileName)
                    ipt_counter += 1
                    Label25.Text = "IPT " & ipt_counter.ToString
                End If
            Next fileName
            DataGridView1.AutoResizeColumns()
        Else
            MessageBox.Show("Directory does not exist")
        End If
        TextBox2.Text &= "Extract_dxf encountered " & ipt_counter.ToString & " ipt files" & vbCrLf
        Button16.BackColor = System.Drawing.Color.Transparent
        TextBox42.Text = " "
        Return (ipt_counter)
    End Function

    Private Sub Find_artikel()
        'Find the the artikel on the assembly drawing (IDW)
        'Print result in DataGridView5
        'DataGridView5 contains old and new file name

        Dim art As String
        Dim ask_once As Boolean = False
        DataGridView5.AllowUserToAddRows = False
        For Each row As System.Windows.Forms.DataGridViewRow In DataGridView5.Rows
            If row.Cells(0).Value IsNot Nothing Then
                art = row.Cells(0).Value.ToString   'Artikel Axxxxxx
                Find_dwg_pos(DataGridView2, art)
                row.Cells(2).Value = kb.actie
                row.Cells(3).Value = kb.Materi
                row.Cells(4).Value = kb.Thick
                row.Cells(5).Value = kb.Count
            End If
        Next
        Remove_empty_rows(DataGridView5)
        DataGridView5.AutoResizeColumns()
    End Sub
    Private Function Rename_dxf() As Integer
        'Find the the artikel on the assembly drawing (IDW)
        'Print result in DataGridView5
        'DataGridView5 contains old and new file name

        Dim old_f, new_f, new_ff As String
        Dim delete_file As Boolean
        Dim ask_once As Boolean = False
        Dim dxf_cnt As Integer = 0

        For Each row As System.Windows.Forms.DataGridViewRow In DataGridView5.Rows
            Increm_progressbar()
            If row.Cells(0).Value IsNot Nothing Then    'Artikel
                If row.Cells(1).Value IsNot Nothing Then   'Preventing exceptions
                    old_f = row.Cells(1).Value.ToString
                Else
                    old_f = "-"
                End If

                If row.Cells(2).Value IsNot Nothing Then   'Preventing exceptions
                    new_f = row.Cells(2).Value.ToString
                Else
                    new_f = "-"
                End If

                new_ff = TextBox34.Text & "\" & new_f   'Full path required
                TextBox42.Text = new_ff
                If IO.File.Exists(new_ff) And ask_once = False Then
                    delete_file = Question_replace_dxf_files()
                    ask_once = True
                End If

                If new_f.Length > 1 Then    'Make sure file name exist
                    If delete_file = True Then
                        IO.File.Delete(new_ff)
                        If Not CheckBox1.Checked Then TextBox2.Text &= "Dxf file " & old_f & " deleted " & vbCrLf
                    End If

                    If Not IO.File.Exists(new_ff) Then
                        My.Computer.FileSystem.RenameFile(old_f, new_f)
                        dxf_cnt += 1
                        Label26.Text = "DXF " & dxf_cnt.ToString
                        If Not CheckBox1.Checked Then TextBox2.Text &= "Dxf file " & new_f & " renamed " & vbCrLf
                    End If
                Else
                    TextBox2.Text &= "Dxf file " & old_f & "Failed NO new name !" & vbCrLf
                End If
            End If
        Next
        DataGridView5.AutoResizeColumns()
        TextBox42.Text = " "
        Return (dxf_cnt)
    End Function

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
        Dim plate_thick As Double

        For Each row As DataGridViewRow In dtg.Rows
            Increm_progressbar()
            If String.Equals(row.Cells.Item(6).Value, Axxxxx) Then           '????? was =
                found = True
                actie = TextBox31.Text & "_"                    'Project
                actie &= TextBox33.Text & "_"                   'Tnumber
                actie &= row.Cells(3).Value.ToString() & "_"    'Drwg= 
                pos = CInt(row.Cells(4).Value)                  'Pos= 
                actie &= pos.ToString("D2")
                If Not CheckBox3.Checked Then
                    actie &= "_" & row.Cells(6).Value.ToString() 'Artikel=  
                End If
                actie &= ".dxf"

                kb.Count = row.Cells(5).Value.ToString()        'Quantity
                kb.Thick = CType(Isolate_thickness(row.Cells(7).Value.ToString()), String)

                '======== thickness plate =======
                Double.TryParse(kb.Thick, plate_thick)
                If plate_thick < 0.1 Then TextBox2.Text &= "Plate thickness " & Axxxxx & " < 0.1 mm, " & row.Cells(7).Value.ToString & vbCrLf

                '======== Plate in name ============
                Check_for_plate(row.Cells(7).Value.ToString())
                If Check_for_plate(row.Cells(7).Value.ToString()) Then
                    TextBox2.Text &= row.Cells(7).Value.ToString & " =PLATE= missing in Artikel " & Axxxxx & vbCrLf
                End If

                kb.Materi = row.Cells(10).Value.ToString()
                kb.actie = actie
                Exit For
            End If
        Next

        If found = True Then
            If Not CheckBox1.Checked Then
                TextBox2.Text &= "IDW BOM list, lookup drwg + pos for Artikel " & Axxxxx & " found" & vbCrLf
            End If
        Else
            TextBox2.Text &= "IDW BOM list, lookup drwg + pos for Artikel " & Axxxxx & " NOT found" & vbCrLf
        End If
    End Sub
    Private Function Isolate_thickness(str As String) As Integer
        Dim delta As Int16
        Dim str2, str3 As String

        str2 = str.Substring(5, 4)
        str3 = System.Text.RegularExpressions.Regex.Replace(str2, "[^\d]", " ")
        Int16.TryParse(str3, delta)

        Return CInt(delta.ToString)
    End Function
    Private Function Check_for_plate(str As String) As Boolean
        Dim exi As Boolean
        str = str.Substring(0, 5)
        exi = CBool(String.Compare(str.ToUpper, "PLATE"))
        'MessageBox.Show(str & "-" & exi.ToString)
        Return (exi)
    End Function

    Private Sub Increm_progressbar()
        ProgressBar1.Value += 1
        ProgressBar2.Value += 1
        ProgressBar3.Value += 1
        If ProgressBar1.Value = 99 Then
            ProgressBar1.Value = 0  'Extract dxf from ipt's
            ProgressBar2.Value = 0  'Find IDW
            ProgressBar3.Value = 0  'IAM BOM
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim cnt As Integer

        '==========WVB button======
        TextBox2.Clear()
        ProgressBar1.Visible = True
        Button19.BackColor = System.Drawing.Color.LightGreen
        DataGridView5.Rows.Clear()
        DataGridView5.RowCount = view_rows    'was 20

        TextBox2.Text &= "============= Find the IDW's =======================" & vbCrLf
        Button19.Text = "Find the IDW's..."
        cnt = Find_IDW()
        If cnt = 0 Then MessageBox.Show("WARNING NO IDW files found in the Work directory !!")
        TextBox2.Text &= "============= Extract dxf from IDW ==================" & vbCrLf
        Button19.Text = "Extract dxf from idw's..."
        cnt = Extract_dxf_from_IPT()
        If cnt = 0 Then MessageBox.Show("WARNING NO DXF files Extraxted from idw's !!")
        TextBox2.Text &= "============= Find artikel drwg + pos and rename ====" & vbCrLf
        Button19.Text = "Lookup artikel dwg and pos..."
        Find_artikel()
        TextBox2.Text &= "============= Rename dxf file ======================" & vbCrLf
        Button19.Text = "Rename dxf file..."
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
        Dim oDoc As Inventor.DrawingDocument
        Dim oActiveSheet As Sheet
        Dim oGeneralNotes As GeneralNotes
        Dim oTG As TransientGeometry
        Dim oGeneralNote As GeneralNote

        Dim x, y As Integer
        Dim f_size As Double
        Dim dest As String
        Dim q_file As String = "-"  'File name

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)
        oDoc = CType(invApp.Documents.Open(fileName, False), DrawingDocument)

        ' Set a reference to the active sheet.
        oActiveSheet = oDoc.ActiveSheet
        ' Set a reference to the GeneralNotes object
        oGeneralNotes = oActiveSheet.DrawingNotes.GeneralNotes

        oTG = invApp.TransientGeometry

        ' Create text with simple string as input. Since this doesn't use
        ' any text overrides, it will default to the active text style.
        Dim sText As String = TextBox36.Text


        x = CInt(NumericUpDown1.Value)
        y = CInt(NumericUpDown2.Value)
        f_size = NumericUpDown3.Value
        oGeneralNote = oGeneralNotes.AddFitted(oTG.CreatePoint2d(x, y), "-")
        oGeneralNote.FormattedText = "<StyleOverride FontSize = '" & f_size.ToString & "'>" & TextBox36.Text & "</StyleOverride>"

        'Save the document
        dest = TextBox35.Text & "\" & IO.Path.GetFileName(fileName)
        oDoc.SaveAs(dest, True)                 'WORKS
        'oDoc.Save()                            'Works
        TextBox2.Text &= "Drawing Note added to " & dest & vbCrLf
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        'http://adndevblog.typepad.com/manufacturing/2012/08/save-a-drawing-to-dwg-dxf-format-using-additional-options-in-an-ini-file.html
        sts()
    End Sub

    Private Sub Sts()
        '=============
        Inventor_running()
        Button17.BackColor = System.Drawing.Color.LightGreen

        'Select work directory
        Dim pathfile As String = TextBox6.Text

        If Directory.Exists(pathfile) Then
            Dim fileEntries As String() = Directory.GetFiles(pathfile)
            For Each fileName In fileEntries
                Increm_progressbar()
                Dim extension As String = IO.Path.GetExtension(fileName)
                If extension = ".idw" Then
                    'Read_title_Block_idw(fileName)
                    'DWGOutUsingTranslatorAddIn(fileName)
                    DWGOutUsingTranslatorAddIn2(fileName)
                End If
            Next fileName
        Else
            MessageBox.Show(pathfile & " is not a valid file or directory.")
        End If
        Button17.BackColor = System.Drawing.Color.Transparent
    End Sub

    Public Sub DWGOutUsingTranslatorAddIn(ByVal path As String)
        ' Set a reference to the DWG translator add-in.
        Dim oDWGAddIn As TranslatorAddIn = Nothing
        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.DrawingDocument
        Dim i As Long

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)

        oDoc = CType(invApp.Documents.Open(path, False), DrawingDocument)

        TextBox40.Text = oDoc.DisplayName
        TextBox41.Text = oDoc.ActiveSheet.Name

        ' ==== saveAs as pdf ============
        If CheckBox4.Checked Then oDoc.SaveAs(path.Substring(0, path.Length - 5) & ".pdf", CBool(vbTrue))

        ' ==== SaveCopyAs as dwg =======
        If CheckBox5.Checked Then
            For i = 1 To invApp.ApplicationAddIns.Count
                If invApp.ApplicationAddIns.Item(CInt(i)).ClassIdString = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}" Then
                    oDWGAddIn = CType(invApp.ApplicationAddIns.Item(CInt(i)), TranslatorAddIn)
                    Exit For
                End If
            Next

            If oDWGAddIn Is Nothing Then
                MessageBox.Show("DWG add-in not found.")
                Exit Sub
            End If

            ' Check to make sure the add-in is activated.
            If Not oDWGAddIn.Activated Then
                oDWGAddIn.Activate()
            End If

            ' Create a name-value map to supply information to the translator.
            Dim oNameValueMap As NameValueMap
            oNameValueMap = invApp.TransientObjects.CreateNameValueMap()

            Dim strIniFile As String
            strIniFile = "C:\Temp\dwgout2.ini"

            ' Create the name-value that specifies the ini file to use
            If IO.File.Exists(strIniFile) Then
                oNameValueMap.Add("Export_Acad_IniFile", strIniFile)
            Else
                MessageBox.Show(strIniFile & " Ini file does not exist")
            End If

            ' Create a translation context and define that we want to output to a file.
            Dim oContext As TranslationContext
            oContext = invApp.TransientObjects.CreateTranslationContext
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

            ' Define the type of output by  specifying the filename.
            Dim oOutputFile As DataMedium
            oOutputFile = invApp.TransientObjects.CreateDataMedium
            oOutputFile.FileName = path.Substring(0, path.Length - 5) & ".dwg"

            oDWGAddIn.SaveCopyAs(oDoc, oContext, oNameValueMap, oOutputFile)
        End If
        TextBox40.Text = ""
        TextBox41.Text = ""
    End Sub

    Public Sub DWGOutUsingTranslatorAddIn2(ByVal path As String)
        ' Set a reference to the DWG translator add-in.
        Dim oDWGAddIn As TranslatorAddIn = Nothing
        Dim invApp As Inventor.Application
        Dim oDoc As Inventor.DrawingDocument

        invApp = CType(Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        invApp.SilentOperation = CBool(vbTrue)

        oDoc = CType(invApp.Documents.Open(path, False), DrawingDocument)
        TextBox40.Text = oDoc.DisplayName
        TextBox41.Text = oDoc.ActiveSheet.Name
        ' ==== saveAs as pdf ============
        If CheckBox4.Checked Then oDoc.SaveAs(path.Substring(0, path.Length - 5) & ".pdf", CBool(vbTrue))

        ' ==== SaveCopyAs as dwg =======
        If CheckBox5.Checked Then

            ' ==== SaveCopyAs as dwg =======
            Dim DWGAddIn As TranslatorAddIn
            DWGAddIn = CType(invApp.ApplicationAddIns.ItemById("{C24E3AC2-122E-11D5-8E91-0010B541CD80}"), TranslatorAddIn)

            If DWGAddIn Is Nothing Then
                MessageBox.Show("DWG add-in not found.")
                Exit Sub
            End If

            Dim oContext As TranslationContext
        oContext = invApp.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = invApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = invApp.TransientObjects.CreateDataMedium

            ' Check whether the translator has 'SaveCopyAs' options
            If DWGAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then
                oOptions.Value("Export_Acad_IniFile") = "C:\Temp\dwgout2.ini"
                oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
                'oOptions.Value("Custom_Begin_Sheet") = 3
                'oOptions.Value("Custom_End_Sheet") = 3
            Else
                MessageBox.Show("The translator has NO 'SaveCopyAs' options")
            End If

            'Set the destination file name
            oDataMedium.FileName = path.Substring(0, path.Length - 5) & ".dwg"

            'Publish document.
            DWGAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)
        End If
        TextBox40.Text = "Done"
        TextBox41.Text = ""
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Button18.BackColor = System.Drawing.Color.LightGreen
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = filepath3
        SaveFileDialog1.FileName = "_Error_log.txt"
        SaveFileDialog1.ShowDialog()
        MessageBox.Show(SaveFileDialog1.FileName)
        My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, TextBox2.Text, False)
        Button18.BackColor = System.Drawing.Color.Transparent
    End Sub
End Class


