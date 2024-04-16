Imports System.IO
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
    Dim ShowExcel As Boolean = True

    Dim lines As String() = {}
    Dim filePath As String


    Dim oxl As Excel.Application
    Dim owb As Excel.Workbook
    Dim osheet As Excel.Worksheet
    Dim YESORNOT As Boolean




    'BUTTON : OK  --> TREE VIEW:
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim customer_input As String = TextBox1.Text

        ' -- FILE SERVER: --
        Dim raw_string As String = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input\{TextBox2.Text}"
        newpath = $"\\fileserver1\Temp\Current Project\{customer_input}\INPUTS\Customer Input"

        TextBox3.Text = raw_string
        Application.DoEvents()

        Dim topnode As TreeNode = TreeView1.Nodes.Add(TextBox2.Text)

        Application.DoEvents()
        PopulateTreeView(TextBox3.Text, topnode.Nodes)

    End Sub


    'SHOW TREE VIEW:
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


    'MULTI - TREE VIEW:
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


    'BUTTON : EXCEL UPDATE
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ProgressBar1.Visible = True
        Application.DoEvents()
        Label6.Visible = True
        Application.DoEvents()
        ProgressBar1.Value = 15
        CreateFolder()
        YESORNOT = True

        oxl = CreateObject("Excel.Application")
        oxl.Visible = False

#Region "XL COPY"
        '---------------------
        'Dim currentDateTime As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")

        Dim Folder As String = $"{TextBox3.Text}\DSM-Template"
        Dim excelFiles As String() = Directory.GetFiles(Folder, "*.xlsx")
        Dim excelfilename As String() = excelFiles(0).Split("\")


        'Dim sourcePath As String = $"{TextBox3.Text}\DSM-Template\XXXXXXBasic Design Data_R0.xlsx"
        Dim sourcePath As String = $"{TextBox3.Text}\DSM-Template\{excelfilename(excelfilename.Length - 1)}"
        Dim destinationPath As String = $"{TextBox3.Text}\Step-1-Output\{excelfilename(excelfilename.Length - 1)}"

        Dim EXCEL_WRITING As String



        If File.Exists(destinationPath) Then

            EXCEL_WRITING = destinationPath
        Else

            File.Copy(sourcePath, destinationPath, True)
            EXCEL_WRITING = destinationPath

        End If

        ProgressBar1.Value = 30

#End Region

        '---------------------
        Application.DoEvents()


        Dim inputString As String = TextBox3.Text

        Dim parts As String() = inputString.Split("\") 'FOLDERS
        Dim value As String = parts(parts.Length - 1) ' TASK FOLDER


        'PATH SELECTED NODE --> FOLDER PATH
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


        'MsgBox("WAIT! UNTILL THE EXCEL POPUPS")
        Label4.Text = "VALIDATING THE SELECTED FOLDER IN THE VIEWS"

        Application.DoEvents()

        ProgressBar1.Value = 50
#Region "EXCEL OPERATION"


        'EXCEL APPLICATION:

        owb = oxl.Workbooks.Open(EXCEL_WRITING, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Editable:=True)
        Application.DoEvents()

        If owb IsNot Nothing Then

            If Checknode(value) And node1.Count = 1 Then

                folderpath = Directory.GetDirectories(TextBox3.Text).ToList

            ElseIf Checknode(value) And node1.Count > 1 Then

                ShowYesNoMessage()



            End If




            For Each z As String In folderpath


                ProgressBar1.Value = 70
                osheet = CType(owb.Sheets(6), Excel.Worksheet)
                Dim currentRow As Integer = 7
                Dim skipRow As Integer
                Dim sno As Integer

                Do

                    Dim cellValue As String = osheet.Cells(currentRow, 2).Value 'osheet.Cells(ROW,COL).value
                    currentRow += 1

                Loop While Not (String.IsNullOrEmpty(osheet.Cells(currentRow, 2).Value) And String.IsNullOrEmpty(osheet.Cells(currentRow + 1, 2).Value))

                If currentRow = 8 Then
                    skipRow = (currentRow - 1) + 1
                    sno = 1
                Else
                    skipRow = (currentRow - 1) + 2
                    sno = CInt(osheet.Cells((currentRow - 1), 2).Value.ToString)
                    sno += 1
                End If


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


                Label4.Text = "WAIT! Untill The XL Pop Ups"


                ' ---------------------------------------EXCEL VALIDATION FOR NULL VALUES:-----------------------------------------------


                'MsgBox(skipRow)

                Application.DoEvents()

                ' ---------------------------------------EXCEL UPDATION WITH RESPECTIVE COLUMNS:-----------------------------------------------
                ProgressBar1.Value = 80

                For i As Integer = 0 To pdfFiles.Count - 1


                    Application.DoEvents()


                    'SNO::sno
                    osheet.Range($"B{i + skipRow }").Value = sno + i
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

                    If (Not (z.Contains("DSM") Or z.Contains("Step"))) And Not (lines.Contains(z)) Then

                        writer.WriteLine(z)
                    End If
                End Using


#End Region
                pdfFiles.Clear()
            Next z

            ProgressBar1.Value = 100

            If YESORNOT Then
                owb.Save()
            End If

            Application.DoEvents()


#End Region

            If ShowExcel Then
                oxl.Visible = True
            End If


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
            ShowExcel = True
            ProgressBar1.Visible = False
#End Region

        End If

    End Sub


    'ALL OR RESET:
    Private Sub ShowYesNoMessage()
        ' Display a MessageBox with Yes and No buttons
        Dim result As DialogResult = MessageBox.Show("SELECT ALL - YES, RESET - NO ?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' Check the user's choice
        If result = DialogResult.Yes Then
            folderpath = Directory.GetDirectories(TextBox3.Text).ToList
            YESORNOT = True
        Else
            ShowExcel = False
            TreeView1.Nodes.Clear()
            pdfFiles.Clear()
            node1.Clear()
            answer = ""
            folderpath.Clear()
            FileName.Clear()
            newpath = ""
            Selected_Folder_Path.Clear()
            owb.Close()
            Button5.PerformClick()
            Kill_Process()
            YESORNOT = False
            Label6.Visible = False


        End If
    End Sub

    Public Function Checknode(value As String)


        For Each i As TreeNode In node1

            If i.ToString.EndsWith(value) Then
                Return True
            End If

        Next

        Return False


    End Function


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
        Label4.Text = ""

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
        filePath = $"{parentFolderPath}/{TextBox2.Text}-LOG.txt"
        If Not (File.Exists(filePath)) Then

            Using writer As New StreamWriter(filePath, True)
                writer.Close()
            End Using

        Else

            lines = File.ReadAllLines(filePath)

        End If

    End Sub
End Class

