Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form1

    Dim node1 As New List(Of TreeNode)
    Dim pdfFiles As New List(Of String)()
    Dim answer As String
    Dim folderpath As New List(Of String)()
    Dim FileName As New List(Of String)
    Dim newpath As String

    'AFTER THE CUSTOMER INPUT BUTTON:
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim customer_input As String = TextBox1.Text
        'Dim raw_string As String = $"\\fileserver1\ENGG_PRODUCTION\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        Dim raw_string As String = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        newpath = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input"

        'Dim raw_string As String = $"D:\TESTING\A1\Input\Customer Input\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        'newpath = $"D:\TESTING\A1\Input\Customer Input\{customer_input}\INPUTS\Customer Input"


        TextBox3.Text = raw_string


        Dim topnode As TreeNode = TreeView1.Nodes.Add(TextBox2.Text)

        ' Add child nodes to the parent node
        'PopulateTreeView(TextBox3.Text, TreeView1.Nodes)
        CreateFolder()
        PopulateTreeView(TextBox3.Text, topnode.Nodes)

    End Sub


    'ADDING THE DIRECTORIES TO THE TREE VIEW:
    Private Sub PopulateTreeView(ByVal directory1 As String, ByVal parentNode As TreeNodeCollection)
        Try

            'parentNode.Add(TextBox2.Text)
            ' Get all subdirectories in the current directory
            Dim subDirectories() As String = Directory.GetDirectories(directory1)

            ' Loop through each subdirectory and add it to the TreeView
            For Each subDirectory As String In subDirectories
                Dim directoryNode As New TreeNode(Path.GetFileName(subDirectory))

                ' Recursively call the PopulateTreeView method for subdirectories
                PopulateTreeView(subDirectory, directoryNode.Nodes)

                ' Add the directory node to the parent node
                parentNode.Add(directoryNode)
            Next
        Catch ex As Exception
            ' Handle any exceptions here
            Console.WriteLine(ex.Message)
        End Try
    End Sub


    'MULTI SELECT FUNCTIONALITY IN TREE VIEW:
    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        ' Check if the node is already selected
        If node1.Contains(e.Node) Then
            ' Node is already selected, so deselect it
            Application.DoEvents()
            TreeView1.SelectedNode = Nothing
            node1.Remove(e.Node)
            Application.DoEvents()
            e.Node.BackColor = TreeView1.BackColor
            Application.DoEvents()
            e.Node.ForeColor = TreeView1.ForeColor
            Application.DoEvents()


        Else
            ' Node is not selected, so select it
            node1.Add(e.Node)
            e.Node.BackColor = SystemColors.Highlight
            e.Node.ForeColor = SystemColors.HighlightText
        End If

    End Sub

    'EXCEL UPDATE BUTTON:
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '---------------------

        ' Specify the source and destination paths
        Dim sourcePath As String = $"{TextBox3.Text}\DSM-Template\XXXXXXBasic Design Data_R0.xlsx"
        Dim destinationPath As String = $"{TextBox3.Text}\{TextBox2.Text}\XXXXXXBasic Design Data_R0.xlsx"

        Try
            ' Copy the file from the source to the destination
            File.Copy(sourcePath, destinationPath, True)

            ' Optionally, you can delete the original file after copying
            ' File.Delete(sourcePath)

            'MessageBox.Show("Excel file saved to new path successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        '---------------------
        Application.DoEvents()

        For Each selectedNode As TreeNode In node1

            Dim answer As String = ""
            While selectedNode IsNot Nothing
                answer = answer + selectedNode.Text + "*"
                selectedNode = selectedNode.Parent
            End While



            Dim result As String = answer.Substring(0, answer.Length - 1)
            'MsgBox(result)
            Dim array1 As String() = result.Split("*")
            Array.Reverse(array1)
            Dim final As String = String.Join("\", array1)
            folderpath.Add($"{newpath}\{final}")
            'MsgBox($"{newpath}\{final}")
            answer = ""
        Next

        MsgBox("WAIT! UNTILL THE EXCEL POPUPS")

        ' -------------------------------------- DIRECTORY LOCATION HARD CODING: ---------------------------------------------'

        ' -------------------------------------- CHANGING THE DIRECTORY LOCATION ---------------------------------------------'

        Application.DoEvents()

#Region "EXCEL OPERATION"

        Dim oxl As Excel.Application
        Dim owb As Excel.Workbook
        Dim osheet As Excel.Worksheet
        oxl = CreateObject("Excel.Application")
        oxl.Visible = False

        'EXCEL APPLICATION:
        'owb = oxl.Workbooks.Open("C:\Users\19433\Desktop\PROJECT AUTOMATES\XXXXXXBasic Design Data_R0.xlsx")
        owb = oxl.Workbooks.Open(destinationPath)
        Application.DoEvents()

        osheet = CType(owb.Sheets(6), Excel.Worksheet)
        'Dim currentRow As Integer = 1
        Dim currentRow As Integer = 7
        Dim skipRow As Integer

        Do

            Dim cellValue As String = osheet.Cells(currentRow, 2).Value 'osheet.Cells(ROW,COL).value
            currentRow += 1

        Loop While Not String.IsNullOrEmpty(osheet.Cells(currentRow, 2).Value)
        skipRow = (currentRow - 1) + 1

        'MsgBox(skipRow)

        For Each z As String In folderpath


            Dim mainFolderPath As String = z
            Dim mainiee As String = mainFolderPath.Replace("/", "\")
            If Directory.Exists(mainiee) Then
                pdfFiles.AddRange(Directory.GetFiles(mainiee, "*.pdf", SearchOption.AllDirectories))
            End If


            ' ---------------------------------------EXCEL VALIDATION FOR NULL VALUES:-----------------------------------------------


            'MsgBox(skipRow)

            Application.DoEvents()

            ' ---------------------------------------EXCEL UPDATION WITH RESPECTIVE COLUMNS:-----------------------------------------------

            For i As Integer = 0 To pdfFiles.Count - 1

                Application.DoEvents()
                'SerialNumber:
                'osheet.Range($"A{i + skipRow }").Value = i + 1
                osheet.Range($"B{i + skipRow }").Value = (i + (skipRow - 1)) - 6
                Application.DoEvents()

                'FileName:
                Dim FileNameWithExtension As String = System.IO.Path.GetFileName(pdfFiles(i))
                Dim FileName As String = FileNameWithExtension.Substring(0, FileNameWithExtension.Length - 4)
                osheet.Range($"D{i + skipRow }").Value = FileName
                Application.DoEvents()

                'ModifiedDate:
                'Dim Raw_DateTime As String = IO.File.GetLastWriteTime(pdfFiles(i)).ToString("MM-DD-YYYY")
                'Dim Raw_DateTime As String = IO.File.GetLastWriteTime(pdfFiles(i)).ToString("dd-MM-yyyy")
                'osheet.Range($"C{i + skipRow }").Value = Raw_DateTime
                'Application.DoEvents()


                'MODIFIED DATE:
                If File.Exists(pdfFiles(i)) Then
                    ' Create a FileInfo object for the PDF file
                    Dim fileInfo As New FileInfo(pdfFiles(i))

                    ' Get the last modified date of the file
                    Dim lastModifiedDate As String = fileInfo.LastWriteTime.ToString()
                    Dim splited As String() = lastModifiedDate.Split(" ")


                    ' Display the last modified date
                    osheet.Range($"C{i + skipRow }").Value = splited(0).Trim
                End If




                'Description:
                'Dim File_Description As String = Description(pdfFiles(i)) ' JUMP INTO DESCRIPTION FUNCTION:
                'osheet.Range($"D{i + skipRow}").Value = File_Description


            Next i
        Next z

        owb.Save()

#End Region

        oxl.Visible = True
        Me.WindowState = FormWindowState.Maximized

#Region "RELOADS"
        'Kill_Process()
        TreeView1.Nodes.Clear()
        pdfFiles.Clear()
        node1.Clear()
        answer = ""
        folderpath.Clear()
        FileName.Clear()
        newpath = ""
#End Region


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

    End Sub

    'KILL THE EXCEL APPLICATION:
    Sub Kill_Process()
        '------------------temp cmnt- for QC---------------------------
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each ProcessX As Process In xlp
            ProcessX.Kill()
            If Process.GetProcessesByName("EXCEL").Count = 0 Then
                Exit For
            End If
        Next
    End Sub



    Sub CreateFolder()
        ' Specify the path of the parent folder
        Dim parentFolderPath As String = TextBox3.Text

        ' Specify the name of the new folder to be created
        Dim newFolderName As String = TextBox2.Text

        ' Combine the parent folder path and new folder name to get the full path
        Dim newFolderPath As String = Path.Combine(parentFolderPath, newFolderName)

        ' Check if the folder already exists
        If Not Directory.Exists(newFolderPath) Then
            ' If it doesn't exist, create the folder
            Directory.CreateDirectory(newFolderPath)
            MsgBox("Folder created successfully.")
        Else
            ' If it already exists, display a message
            MsgBox("Folder already exists.")
        End If

    End Sub
End Class

