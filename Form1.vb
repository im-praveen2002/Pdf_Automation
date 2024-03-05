Imports System.IO
Imports iText.Layout.Element
Imports Microsoft.Office.Interop

Public Class Form1

    Dim Selected_Folder_Path As New List(Of String)()
    Dim dateforfolder As DateTime
    Dim node1 As New List(Of TreeNode)
    Dim pdfFiles As New List(Of String)()
    Dim answer As String
    Dim folderpath As New List(Of String)()
    Dim FileName As New List(Of String)
    Dim newpath As String

    Dim lines As String() = {}
    Dim filePath As String



    Public Function Checknode(value As String)


        For Each i As TreeNode In node1

            If i.ToString.EndsWith(value) Then
                Return True
            End If

        Next

        Return False


    End Function
    'AFTER THE CUSTOMER INPUT BUTTON:
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim customer_input As String = TextBox1.Text
        'Dim raw_string As String = $"\\fileserver1\ENGG_PRODUCTION\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        'newpath = $"\\fileserver1\ENGG_PRODUCTION\Current Project\{customer_input}\INPUTS\Customer Input"

        ' -- FILE SERVER: --
        Dim raw_string As String = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        newpath = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input"


        ' -- LOCAL DISK: --
        'Dim raw_string As String = $"D:\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        'newpath = $"D:\Current Project\{customer_input}\INPUTS\Customer Input"


        TextBox3.Text = raw_string


        Dim topnode As TreeNode = TreeView1.Nodes.Add(TextBox2.Text)

        ' Add child nodes to the parent node
        'PopulateTreeView(TextBox3.Text, TreeView1.Nodes)


        Application.DoEvents()
        PopulateTreeView(TextBox3.Text, topnode.Nodes)

    End Sub


    'ADDING THE DIRECTORIES TO THE TREE VIEW:
    Private Sub PopulateTreeView(ByVal directory1 As String, ByVal parentNode As TreeNodeCollection)
        Try


            Dim subDirectories() As String = Directory.GetDirectories(directory1)

            For Each subDirectory As String In subDirectories
                Dim directoryNode As New TreeNode(Path.GetFileName(subDirectory))

                PopulateTreeView(subDirectory, directoryNode.Nodes)
                parentNode.Add(directoryNode)
            Next
        Catch ex As Exception
            ' Handle any exceptions here
            Console.WriteLine(ex.Message)
        End Try
    End Sub


    'MULTI SELECT FUNCTIONALITY IN TREE VIEW:
    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

        'Check If the node Is already selected
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


            ' SELECT THE NEWNODE:
        Else
            node1.Add(e.Node)
            e.Node.BackColor = SystemColors.Highlight
            e.Node.ForeColor = SystemColors.HighlightText
        End If




    End Sub
    Sub sortedarray()

    End Sub

    'EXCEL UPDATE BUTTON:
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CreateFolder()
        sortedarray()



#Region "XL COPY"
        '---------------------
        Dim currentDateTime As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
        Dim sourcePath As String = $"{TextBox3.Text}\DSM-Template\XXXXXXBasic Design Data_R0.xlsx"
        Dim destinationPath As String = $"{TextBox3.Text}\Step-1-Output\XXXXXXBasic Design Data_R0.xlsx"
        Dim EXCEL_WRITING As String

        '---------------------

        If File.Exists(destinationPath) Then

            EXCEL_WRITING = destinationPath
        Else

            File.Copy(sourcePath, destinationPath, True)
            EXCEL_WRITING = destinationPath

        End If



#End Region

        '---------------------
        Application.DoEvents()


        Dim inputString As String = TextBox3.Text

        Dim parts As String() = inputString.Split("\")
        Dim value As String = parts(parts.Length - 1)



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
            answer = ""
        Next


        MsgBox("WAIT! UNTILL THE EXCEL POPUPS")

        Application.DoEvents()

#Region "EXCEL OPERATION"

        Dim oxl As Excel.Application
        Dim owb As Excel.Workbook
        Dim osheet As Excel.Worksheet
        oxl = CreateObject("Excel.Application")
        oxl.Visible = False

        'EXCEL APPLICATION:
        owb = oxl.Workbooks.Open(EXCEL_WRITING)
        Application.DoEvents()

        osheet = CType(owb.Sheets(6), Excel.Worksheet)
        Dim currentRow As Integer = 7
        Dim skipRow As Integer

        Do

            Dim cellValue As String = osheet.Cells(currentRow, 2).Value 'osheet.Cells(ROW,COL).value
            currentRow += 1

        Loop While Not String.IsNullOrEmpty(osheet.Cells(currentRow, 2).Value)
        skipRow = (currentRow - 1) + 1

        'MsgBox(skipRow)
        'myList.IndexOf(elementToCheck) <> -1 

        If Checknode(value) And node1.Count = 1 Then

            folderpath = Directory.GetDirectories(TextBox3.Text).ToList

        ElseIf Checknode(value) And node1.Count > 1 Then

            MsgBox("SELECTED ALL AND CHILD FOLDER TOO!!")
            folderpath = Directory.GetDirectories(TextBox3.Text).ToList


        End If




        For Each z As String In folderpath


            Dim folderInfo1 As New DirectoryInfo(z)

            ' Check if the folder exists
            If folderInfo1.Exists Then
                ' Get the creation time of the folder
                dateforfolder = folderInfo1.CreationTime

            End If

            Dim mainiee As String = z

            If lines.Contains(z) Then

                'MsgBox($"{mainiee} --> REPEATED RECORDS FOUNDED!! ")
                MessageBox.Show($"{mainiee} --> REPEATED RECORDS FOUNDED!! ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Application.DoEvents()
                'Exit For


            Else
                If Directory.Exists(mainiee) Then
                    Selected_Folder_Path.Add(mainiee)
                    pdfFiles.AddRange(Directory.GetFiles(mainiee, "*.pdf", SearchOption.AllDirectories))
                End If
            End If





            ' ---------------------------------------EXCEL VALIDATION FOR NULL VALUES:-----------------------------------------------


            'MsgBox(skipRow)

            Application.DoEvents()

            ' ---------------------------------------EXCEL UPDATION WITH RESPECTIVE COLUMNS:-----------------------------------------------

            For i As Integer = 0 To pdfFiles.Count - 1
                Application.DoEvents()

                'SNO::
                osheet.Range($"B{i + skipRow }").Value = (i + (skipRow - 1)) - 6
                Application.DoEvents()


                'FILENAME:
                Dim FileNameWithExtension As String = System.IO.Path.GetFileName(pdfFiles(i))
                Dim FileName As String = FileNameWithExtension.Substring(0, FileNameWithExtension.Length - 4)
                osheet.Range($"D{i + skipRow }").Value = FileName.ToString
                Application.DoEvents()


                'DATE:
                osheet.Range($"C{i + skipRow }").Value = dateforfolder.ToString("dd-MM-yyyy")
                Application.DoEvents()
            Next i


#Region "Log Writing"

            Using writer As New StreamWriter(filePath, True)

                If Not (z.Contains("DSM") OrElse z.Contains("Step")) And Not (lines.Contains(z)) Then

                    writer.WriteLine(z)
                End If
            End Using


#End Region
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
        Selected_Folder_Path.Clear()
        'dateforfolder = ""
#End Region


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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

    End Sub

    'FOLDER CREATION WITH THE TEXTBOX2:
    Sub CreateFolder()

        Dim parentFolderPath As String = TextBox3.Text
        Dim newFolderName As String = "Step-1-Output"

        'COMBINING THE PATH:
        Dim newFolderPath As String = Path.Combine(parentFolderPath, newFolderName)

        If Not Directory.Exists(newFolderPath) Then
            Directory.CreateDirectory(newFolderPath)
            'MsgBox("Folder created successfully.")
        Else
            'MsgBox("Folder already exists.")
        End If


        'FILE NOT PRESENT CREATE || READ THE FILE CONTENTS:
        filePath = $"{parentFolderPath}/{TextBox2.Text}"
        If Not (File.Exists(filePath)) Then

            Using writer As New StreamWriter(filePath, True)
                writer.Close()
            End Using

        Else

            lines = File.ReadAllLines(filePath)

        End If

    End Sub

End Class

