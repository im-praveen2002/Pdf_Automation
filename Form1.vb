Imports System.IO
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser
Imports iText.Kernel.Pdf.Canvas.Parser.Listener
Imports Microsoft.Office.Interop

Public Class Form1


    Dim node1 As New List(Of TreeNode)
    Dim pdfFiles As New List(Of String)()
    Dim answer As String
    Dim folderpath As New List(Of String)()

    Private Function GetPathToRoot(node As TreeNode, path As List(Of TreeNode))
        If node Is Nothing Then
            Return vbEmpty
        Else
            path.Add(node)
            Return GetPathToRoot(node.Parent, path)
        End If
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For Each selectedNode As TreeNode In node1
            'Dim parentname As String = selectedNode.Parent.Text
            'MessageBox.Show(GetPathToRoot(selectedNode, node1))
            'MessageBox.Show($"{selectedNode.Parent.Text}/{selectedNode.Text}")
            Dim answer As String
            While selectedNode IsNot Nothing
                answer = answer + selectedNode.Text + "*"
                selectedNode = selectedNode.Parent
            End While



            Dim result As String = answer.Substring(0, answer.Length - 1)
            'MsgBox(result)
            Dim array1 As String() = result.Split("*")
            Array.Reverse(array1)
            Dim final As String = String.Join("\", array1)
            folderpath.Add($"{TB1.Text}\{final}")
            'MsgBox($"{TB1.Text}\{final}")
            answer = ""
        Next

        MsgBox("COMPLETED!")

    End Sub

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

    Private Sub PopulateTreeView(ByVal directory1 As String, ByVal parentNode As TreeNodeCollection)
        Try
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
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' -------------------------------------- DIRECTORY LOCATION HARD CODING: ---------------------------------------------'

        '-------------------------------------------------------------------------------------------------------------------------

        ' -------------------------------------- CHANGING THE DIRECTORY LOCATION ---------------------------------------------'

        Dim oxl As Excel.Application
        Dim owb As Excel.Workbook
        Dim osheet As Excel.Worksheet
        oxl = CreateObject("Excel.Application")
        oxl.Visible = True

        'EXCEL APPLICATION:
        owb = oxl.Workbooks.Open("C:\Users\19433\Desktop\PROJECT AUTOMATES\Output.xlsx")
        Application.DoEvents()



        For Each z As String In folderpath


            Dim mainFolderPath As String = z
            Dim mainiee As String = mainFolderPath.Replace("/", "\")
            If Directory.Exists(mainiee) Then
                pdfFiles.AddRange(Directory.GetFiles(mainiee, "*.pdf", SearchOption.AllDirectories))
            End If


            ' ---------------------------------------EXCEL VALIDATION FOR NULL VALUES:-----------------------------------------------

            osheet = CType(owb.Sheets(1), Excel.Worksheet)
            Dim currentRow As Integer = 1
            Dim skipRow As Integer

            Do

                Dim cellValue As String = osheet.Cells(currentRow, 1).Value 'osheet.Cells(ROW,COL).value
                currentRow += 1

            Loop While Not String.IsNullOrEmpty(osheet.Cells(currentRow, 1).Value)
            skipRow = (currentRow - 1) + 1
            'MsgBox(skipRow)

            Application.DoEvents()

            ' ---------------------------------------EXCEL UPDATION WITH RESPECTIVE COLUMNS:-----------------------------------------------

            For i As Integer = 0 To pdfFiles.Count - 1

                Application.DoEvents()
                'SerialNumber:
                osheet.Range($"A{i + skipRow }").Value = i + 1
                Application.DoEvents()

                'FileName:
                Dim FileNameWithExtension As String = System.IO.Path.GetFileName(pdfFiles(i))
                Dim FileName As String = FileNameWithExtension.Substring(0, FileNameWithExtension.Length - 4)
                osheet.Range($"B{i + skipRow }").Value = FileName
                Application.DoEvents()

                'ModifiedDate:
                Dim Raw_DateTime As String = IO.File.GetLastWriteTime(pdfFiles(i)).ToString("MM-dd-yyyy")
                osheet.Range($"C{i + skipRow }").Value = Raw_DateTime
                Application.DoEvents()

                'Description:
                'Dim File_Description As String = Description(pdfFiles(i)) ' JUMP INTO DESCRIPTION FUNCTION:
                'osheet.Range($"D{i + skipRow}").Value = File_Description


            Next i
            oxl.Visible = True

        Next z

        pdfFiles.Clear()

        MsgBox("COMPLETED!")
        Kill_Process()


    End Sub

    'DESCRIPTION FOR DIFFERENT FILES:
    Public Function Description(document As String)

        Dim pdfFilePath As String = document

        'LINE BY LINE IN THE PDF WORDS:
        Dim pdfLines As New List(Of String)()

        Using pdfReader As New PdfReader(pdfFilePath)
            Using pdfDocument As New PdfDocument(pdfReader)


                Dim strategy As New SimpleTextExtractionStrategy()
                Dim pageContent As String = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(1), strategy)
                Dim lines As String() = pageContent.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                pdfLines.AddRange(lines)

            End Using
        End Using

        Dim splited As String() = pdfFilePath.Split("\")
        Dim extracted As String = splited(splited.Length - 1)
        Dim filename As String = extracted.Substring(0, extracted.Length - 4)



        'IF THE PDF FILE IS UNREADABLE OR EMPTY:
        If pdfLines.Count = 0 AndAlso Not filename.StartsWith("SHAHEEN") Then
            answer = ""
            Return answer
        End If

        'TR3032_519 FILES:
        If filename.Contains("_") AndAlso filename.StartsWith("TR") Then

            Dim pattern As String() = filename.Split("_")

            Dim count As Integer = 0
            For Each line As String In pdfLines

                If line.Contains(pattern(0)) = True Then
                    answer = line

                    Exit For

                ElseIf count >= 6 Then
                    answer = ""
                    Exit For

                End If
                count += 1

            Next



            'TR2023 PACKAGE FILES:
        ElseIf filename.StartsWith("TR") AndAlso filename.Contains(" ") Then

            Dim pattern As String() = filename.Split(" ")

            Dim count As Integer

            For Each line As String In pdfLines

                If line.Contains(pattern(0)) AndAlso count <= 6 Then
                    answer = line

                    Exit For

                ElseIf count >= 6 Then
                    answer = ""
                    Exit For


                End If
                count += 1
            Next

            'C030 FILES:
        ElseIf filename.StartsWith("C030") And filename.Contains("-xx") Then

            Dim count As Integer = 0
            Dim result As String = ""

            For Each line As String In pdfLines

                If count < 2 Then
                    result += line + " "

                Else
                    Exit For

                End If

                count += 1
            Next

            answer = result


            'SDRL FILES:
        ElseIf filename.StartsWith("SDRL") Then

            Dim count As Boolean = False

            For Each line As String In pdfLines

                If line.Contains("Documentation description") Then
                    count = True
                ElseIf count = True Then
                    answer = line
                    Exit For

                End If
            Next

            'PACKAGE FILES:
        ElseIf filename.StartsWith("Package Specification") Then

            Dim count As Integer = 0
            Dim result As String = ""

            For Each line As String In pdfLines

                If count < 2 Then
                    result += line + " "

                Else
                    Exit For

                End If

                count += 1
            Next

            answer = result

            'P&ID FILES:
        ElseIf filename.StartsWith("P&ID") Then

            Dim count As Integer = 1
            Dim result As String = ""

            For Each line As String In pdfLines

                If line.Contains("DRAWING TITLE") Then
                    count += 1
                ElseIf count = 2 Then
                    result += line + " "
                    count += 1
                ElseIf count = 3 Then
                    result += line
                    Exit For
                End If
            Next

            answer = result


            'NORSOK M
        ElseIf filename.StartsWith("NORSOK M_") Then


            Dim count As Integer = 0
            Dim result1 As String = ""
            Dim result2 As String = ""

            For Each line As String In pdfLines
                If line.Contains("NORSOK M") Then

                    If count = 0 Then
                        count += 1
                    ElseIf count = 1 Then
                        result1 = line
                        Exit For
                    End If
                End If
            Next

            Dim count1 As Boolean = False

            For Each line As String In pdfLines
                If line.Contains("NORSOK M") Then
                    count1 = True
                ElseIf count1 = True Then
                    result2 = line
                    Exit For
                End If
            Next

            answer = result1 + " " + result2


            'NORSOK N
        ElseIf filename.StartsWith("NORSOK N_") Then


            Dim count As Integer = 0
            Dim result1 As String = ""
            Dim result2 As String = ""

            For Each line As String In pdfLines
                If line.Contains("NORSOK N") Then

                    If count = 0 Then
                        count += 1
                    ElseIf count = 1 Then
                        result1 = line
                        Exit For
                    End If
                End If
            Next

            Dim count1 As Boolean = False

            For Each line As String In pdfLines
                If line.Contains("NORSOK N") Then
                    count1 = True
                ElseIf count1 = True Then
                    result2 = line
                    Exit For
                End If
            Next

            answer = result1 + " " + result2

        ElseIf filename.StartsWith("SHAHEEN") And filename.contains("_") Then

            answer = ""



            'SHAHEEN FILES WITH SPACES:
        ElseIf filename.StartsWith("SHAHEEN") And filename.contains(" ") Then

            Dim array As String() = filename.Split(" ")
            Dim result As String = String.Join(" ", array.Skip(1).ToArray())
            answer = result




            'SHAHEEN FILES WITHOUT SPACES:
        ElseIf filename.StartsWith("SHAHEEN") AndAlso pdfLines.IndexOf("UNIT") < pdfLines.Count() Then


            Dim count As Boolean = False
            For Each line As String In pdfLines
                If line.Contains("UNIT") Then
                    count = True
                ElseIf count Then
                    answer = line
                    Exit For
                End If

            Next


        Else
            answer = "nothing"

        End If

        Return answer


    End Function

    'PDF FILEEXTRACTION:
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim pdfFilePath As String = "Z:\1 Praveen Kumar\Task 6\Input\ATTACHMENT C1 Data Sheet & Gas Composition Primary Buffer Composition Table\SHAHEEN-2112-ME-EDS-K-211201.pdf"


        Dim pdfLines As New List(Of String)()

        Using pdfReader As New PdfReader(pdfFilePath)
            Using pdfDocument As New PdfDocument(pdfReader)


                Dim strategy As New SimpleTextExtractionStrategy()
                Dim pageContent As String = PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(1), strategy)
                Dim lines As String() = pageContent.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                pdfLines.AddRange(lines)

            End Using
        End Using

        Dim outputFilename As String = $"{pdfFilePath.Substring(0, pdfFilePath.Length - 4)}.txt"

        Using writer As New StreamWriter(outputFilename)

            For Each i As String In pdfLines
                writer.WriteLine(i)
            Next

        End Using

        MsgBox("Genreated!")

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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim customer_input As String = TextBox1.Text
        Dim raw_string As String = $"E:\TESTING\{customer_input}\Input\Customer Input"


        Dim folderBrowserDialog1 As New FolderBrowserDialog()
        folderBrowserDialog1.SelectedPath = raw_string
        folderBrowserDialog1.Description = "Select a Folder"

        If folderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Dim selectedFolderPath As String = folderBrowserDialog1.SelectedPath
            TB1.Text = selectedFolderPath
            'MessageBox.Show("Selected folder: " & selectedFolderPath)

        End If

        ' Call the PopulateTreeView method to populate the TreeView
        PopulateTreeView(TB1.Text, TreeView1.Nodes)
    End Sub




End Class

