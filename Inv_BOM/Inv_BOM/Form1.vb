﻿Imports System.Runtime.InteropServices
Imports Inventor
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop

Public Class Form1
    Public filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Public filepath3 As String = "c:\MyDir"

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Find_directory(1)   'ipt files
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Find_directory(2)    'iam files
    End Sub

    Private Sub Find_directory(keuze As Integer)

        ' Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog With {
               .InitialDirectory = "c:\Inventor test files\",
               .Filter = "Part File (*.ipt)|*.ipt" _
               & "|Assembly File (*.iam)|*.iam" _
               & "|Presentation File (*.ipn)|*.ipn" _
               & "|Drawing File (*.idw)|*.idw" _
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
        Qbom()

    End Sub

    Private Sub Qbom()
        Dim information As System.IO.FileInfo
        Dim filen As String

        DataGridView1.ColumnCount = 15
        DataGridView1.ColumnHeadersVisible = True

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
            DataGridView1.Columns(13).HeaderText = "IT_CL"

            For i = 1 To oBOMView.BOMRows.Count
                r = i - 1

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


                '--------- CUSTOM Properties ------------------------
                oPropSet = oCompDef.Document.PropertySets.Item("Inventor User Defined Properties")
                If oPropSet.Count = 0 Then
                    MessageBox.Show("The are NO 'Custom' properties present in this file")
                Else
                    DataGridView1.Rows.Item(r).Cells(6).Value = oPropSet.Item("DOC_NUMBER").Value
                    DataGridView1.Rows.Item(r).Cells(7).Value = oPropSet.Item("ITEM_NR").Value
                    DataGridView1.Rows.Item(r).Cells(8).Value = oPropSet.Item("DOC_STATUS").Value
                    DataGridView1.Rows.Item(r).Cells(9).Value = oPropSet.Item("DOC_REV").Value
                    DataGridView1.Rows.Item(r).Cells(10).Value = oPropSet.Item("PART_MATERIAL").Value
                    DataGridView1.Rows.Item(r).Cells(11).Value = oPropSet.Item("IT_TP").Value
                    'DataGridView1.Rows.Item(r).Cells(12).Value = oPropSet.Item("LG").Value
                    'DataGridView1.Rows.Item(r).Cells(13).Value = oPropSet.Item("IT_TP").Value
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
        Write_excel()
    End Sub
    Private Sub Write_excel()
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorksheet As Excel.Worksheet
        Dim fname As String
        Dim str As String

        Button4.BackColor = Color.Red

        xlApp = CreateObject("Excel.Application")
        xlWorkBook = xlApp.Workbooks.Add(Type.Missing)
        xlWorksheet = xlWorkBook.Worksheets(1)

        xlApp.Visible = False
        xlApp.DisplayAlerts = False 'Suppress excel messages

        Try
            For vert = 1 To DataGridView1.Rows.Count - 1
                For hor = 1 To DataGridView1.Columns.Count - 1
                    str = DataGridView1.Rows.Item(vert - 1).Cells(hor).Value
                    xlWorksheet.Cells(vert, hor) = str
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

End Class


