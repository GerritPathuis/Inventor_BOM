Imports System.IO
Imports System.String
Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.ComponentModel

Public Class Form1
    Public filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Public filepath3 As String = "c:\MyDir"

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Open_file(1)   'ipt files
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Open_file(2)    'iam files
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Open_file(4)    'idw files
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
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button3.BackColor = System.Drawing.Color.Yellow
        DataGridView1.ClearSelection()
        Qbom()
        Button3.BackColor = System.Drawing.Color.Transparent
    End Sub

    Private Sub Qbom()
        Dim information As System.IO.FileInfo
        Dim filen As String

        DataGridView1.ColumnCount = 19
        DataGridView1.ColumnHeadersVisible = True
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        '-------- inventor must be running----
        Dim p() As Process
        p = Process.GetProcessesByName("Inventor")
        If p.Count = 0 Then
            MessageBox.Show("Inventor is not running")
            Exit Sub
        End If

        '------- get file inf0-----------
        information = My.Computer.FileSystem.GetFileInfo(filepath1)
        filen = information.Name

        Dim oDoc As Inventor.Document
        Dim invApp As Inventor.Application

        invApp = Marshal.GetActiveObject("Inventor.Application")

        invApp.SilentOperation = vbTrue
        oDoc = CType(invApp.Documents.Open(filepath1, False), Document)
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
            Dim i, r As Integer

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

            For i = 1 To oBOMView.BOMRows.Count
                r = i - 1

                '--------- Design Tracking Properties ------------------------
                oRow = oBOMView.BOMRows.Item(i)
                oCompDef = oRow.ComponentDefinitions.Item(1)

                oPropSet = oCompDef.Document.PropertySets.Item("Design Tracking Properties")
                DataGridView1.Rows.Add()

                DataGridView1.Rows.Item(r).Cells(0).Value = filen

                DataGridView1.Rows.Item(r).Cells(1).Value = oRow.ItemNumber
                DataGridView1.Rows.Item(r).Cells(2).Value = oRow.ItemQuantity

                DataGridView1.Rows.Item(r).Cells(3).Value = oPropSet.Item("Part Number").Value
                DataGridView1.Rows.Item(r).Cells(4).Value = oPropSet.Item("Description").Value
                DataGridView1.Rows.Item(r).Cells(5).Value = oPropSet.Item("Stock Number").Value
                Try
                    DataGridView1.Rows.Item(r).Cells(13).Value = oPropSet.Item(31).Name & " " & oPropSet.Item(32).Name
                Catch Ex As Exception
                    If Not CheckBox1.Checked Then MessageBox.Show("Part Icon not found")
                End Try

                '--------- CUSTOM Properties ------------------------
                oPropSet = oCompDef.Document.PropertySets.Item("Inventor User Defined Properties")
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Custom' properties present in this file")
                Else
                    Try
                        DataGridView1.Rows.Item(r).Cells(6).Value = oPropSet.Item("DOC_NUMBER").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("DOC_NUMBER not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(7).Value = oPropSet.Item("ITEM_NR").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("ITEM_NR not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(8).Value = oPropSet.Item("DOC_STATUS").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("DOC_STATUS not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(9).Value = oPropSet.Item("DOC_REV").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("DOC_REV not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(10).Value = oPropSet.Item("PART_MATERIAL").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("PART_MATERIAL not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(11).Value = oPropSet.Item("IT_TP").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("IT_TP not found")
                    End Try
                    Try
                        DataGridView1.Rows.Item(r).Cells(12).Value = oPropSet.Item("LG").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("LENGTH not found")
                    End Try

                End If

                '--------- Inventor Summary Information ---------------- 
                oPropSet = oCompDef.Document.PropertySets.Item("Inventor Summary Information")
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Inventor Summary Information' present in this file")
                Else
                    Try
                        DataGridView1.Rows.Item(r).Cells(14).Value = oPropSet.Item("Title").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("Title not found")
                    End Try

                    Try
                        DataGridView1.Rows.Item(r).Cells(15).Value = oPropSet.Item("Subject").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("Subject not found")
                    End Try

                    Try
                        DataGridView1.Rows.Item(r).Cells(16).Value = oPropSet.Item("Author").Value
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("Author not found")
                    End Try

                    Try
                        DataGridView1.Rows.Item(r).Cells(17).Value = oPropSet.Item(17).GetType.ToString
                    Catch Ex As Exception
                        If Not CheckBox1.Checked Then MessageBox.Show("Comments not found")
                    End Try
                End If
                '-----------------------------------------------------

            Next
        Catch Ex As Exception
            MessageBox.Show("No BOM in this drawing ")
        Finally
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SaveFileDialog1.Title = "Please Select a File"
        SaveFileDialog1.InitialDirectory = filepath3
        SaveFileDialog1.FileName = "Inventor_BOM.xls"
        SaveFileDialog1.ShowDialog()
        Button4.BackColor = System.Drawing.Color.Yellow
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

        '-------- Header text to excel -------------
        For hor = 1 To DataGridView1.Columns.Count - 1
            xlWorksheet.Cells(1, hor) = DataGridView1.Columns(hor - 1).HeaderText
        Next

        '-------- Cell_text to excel -------------
        Try
            For vert = 1 To DataGridView1.Rows.Count - 1
                For hor = 1 To DataGridView1.Columns.Count - 1
                    str = DataGridView1.Rows.Item(vert - 1).Cells(hor).Value
                    xlWorksheet.Cells(vert + 1, hor) = str
                Next
            Next
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
        Dim oDoc As Inventor.Document
        Dim invApp As Inventor.Application

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

            'oPropSet = oPropsets.Item("Inventor User Defined Properties")

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

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ChgTitleBlkDef()
    End Sub
    Private Sub ChgTitleBlkDef()
        'http://adndevblog.typepad.com/manufacturing/2012/12/inventor-change-text-items-in-titleblockdefinition.html
        TextBox2.Clear()

        Dim oApp As Inventor.Application
        oApp = CType(GetObject(, "Inventor.Application"), Inventor.Application)
        oApp.SilentOperation = vbTrue

        Dim objDrawDoc As DrawingDocument = CType(oApp.ActiveDocument, DrawingDocument)
        objDrawDoc = CType(oApp.Documents.Open(filepath1, False), Document)
        TextBox2.Text &= "objDrawDoc is " & objDrawDoc.ToString & vbCrLf

        Dim colTitleBlkDefs As TitleBlockDefinitions = objDrawDoc.TitleBlockDefinitions

        Dim objTitleBlkDef As TitleBlockDefinition = Nothing
        For Each objTitleBlkDef In colTitleBlkDefs
            TextBox2.Text &= "objTitleBlkDef name = " & objTitleBlkDef.Name & vbCrLf
            If objTitleBlkDef.Name = "DIN" Then
                TextBox2.Text &= "Found Title Block DIN !" & vbCrLf
                Exit For
            End If
        Next

        TextBox2.Text &= "----------------" & vbCrLf

        ' If we are here we have the title block of interest.
        ' Get the title block sketch and set it active

        Dim objDrwSketch As DrawingSketch = Nothing
        objTitleBlkDef.Edit(objDrwSketch)

        Dim colTextBoxes As Inventor.TextBoxes = objDrwSketch.TextBoxes

        For Each objTextBox As Inventor.TextBox In colTextBoxes
            TextBox2.Text &= "objTextBox.Text = " & objTextBox.Text & vbCrLf
            If objTextBox.Text = "<TITLE>" Then
                TextBox2.Text &= "TITLE is !" & objTextBox.Text & vbCrLf
                'objTextBox.Text = "Captain CAD Engineering"
                Exit For
            End If
        Next
        objTitleBlkDef.ExitEdit(False)

        Beep()
    End Sub


End Class


