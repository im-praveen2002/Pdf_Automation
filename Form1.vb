Imports System.IO
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser
Imports iText.Kernel.Pdf.Canvas.Parser.Listener
Imports Microsoft.Office.Interop

Public Class Form1

    Dim pdfFiles As New List(Of String)()
    Dim answer As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Dim customer_input As String = TextBox1.Text
        Dim raw_string As String = $"E:\TESTING\{customer_input}\Input\Customer Input"


        Dim folderBrowserDialog1 As New FolderBrowserDialog()

        folderBrowserDialog1.SelectedPath = raw_string

        folderBrowserDialog1.Description = "Select a Folder"

        If folderBrowserDialog1.ShowDialog() = DialogResult.OK Then

            Dim selectedFolderPath As String = folderBrowserDialog1.SelectedPath
            TB1.Text = selectedFolderPath

            MessageBox.Show("Selected folder: " & selectedFolderPath)

        End If


        ' ------------------------------------------------------------ CHANGING THE DIRECTORY LOCATION ---------------------------------------------'

        Application.DoEvents()
        'Dim mainfolderpath1 As String = TB1.Text

        ' Specify the path of the main folder
        Dim mainFolderPath As String = TB1.Text

        ' Check if the main folder exists
        If Directory.Exists(mainFolderPath) Then
            ' Get all PDF file names in the main folder and its subfolders
            pdfFiles.AddRange(Directory.GetFiles(mainFolderPath, "*.pdf", SearchOption.AllDirectories))
        End If

        Dim oxl As Excel.Application
        Dim owb As Excel.Workbook
        Dim osheet As Excel.Worksheet


        oxl = CreateObject("Excel.Application")
        oxl.Visible = True


        owb = oxl.Workbooks.Open("C:\Users\19433\Desktop\PROJECT AUTOMATES\Output.xlsx")
        Application.DoEvents()

        ' Reference the first sheet (you might want to change this based on your specific sheet)

        osheet = CType(owb.Sheets(1), Excel.Worksheet)

        ' Add a value to cell A1 (you can change the cell reference as needed)

        ' ----------------------------------------------------------------EXCEL CREATION WITH RESPECTIVE COLUMNS:------------------------------------------------

        For i As Integer = 0 To pdfFiles.Count - 1


            'SerialNumber:
            osheet.Range($"A{i + 2}").Value = i + 1

            'FileName:
            Dim FileNameWithExtension As String = System.IO.Path.GetFileName(pdfFiles(i))
            Dim FileName As String = FileNameWithExtension.Substring(0, FileNameWithExtension.Length - 4)
            osheet.Range($"B{i + 2}").Value = FileName

            'ModifiedDate:
            Dim Raw_DateTime As String = IO.File.GetLastWriteTime(pdfFiles(i)).ToString("MM-dd-yyyy")
            osheet.Range($"C{i + 2}").Value = Raw_DateTime

            'Description:
            Dim File_Description As String = Description(pdfFiles(i)) ' JUMP INTO DESCRIPTION FUNCTION:
            osheet.Range($"D{i + 2}").Value = File_Description

        Next i

        MsgBox("COMPLETED!")

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
        If pdfLines.Count = 0 Then
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

        ElseIf filename.StartsWith("SHAHEEN") Then

            Dim result As String = ""
            For i As Integer = 0 To pdfLines.Count

                If pdfLines(i + 1).Contains("SHAHEEN-COM") Then
                    result = pdfLines(i + 1) + " " + pdfLines(i)
                    Exit For
                End If

            Next i
            answer = result


        Else
            answer = ""

        End If

        Return answer

    End Function

    'PDF FILEEXTRACTION:
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim pdfFilePath As String = "Z:\1 Praveen Kumar\Task 6\Input\ATTACHMENT A Applied Specifications (Customer Specifications Summary)\01_General\SHAHEEN-COM-DM-PRO-0004-1 PROJECT NUMBERING PROCEDURE.pdf"


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

End Class

