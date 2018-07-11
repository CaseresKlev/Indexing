Imports Microsoft.Office.Interop.Word
Imports MySql.Data.MySqlClient
Imports System.IO


Public Class Form1

    Dim con As MySqlConnection
    Dim db_con = New db_config
    Dim query As String
    Dim cmd As MySqlCommand
    Dim wordApplication As Microsoft.Office.Interop.Word.Application = Nothing
    Dim wordDocument As Microsoft.Office.Interop.Word.Document = Nothing
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            OpenFileDialog1 = New OpenFileDialog()
            'Make name
            OpenFileDialog1.Title = "Load DataBase"
            'Filter
            OpenFileDialog1.Filter = "PDF File (*.docx)|*.docx"
            'set the root to the z drive
            OpenFileDialog1.InitialDirectory = "Z:\"
            'make sure the root goes back to where the user started
            'openFileDialog1.RestoreDirectory = True
            'show the dialog
            OpenFileDialog1.ShowDialog()
            'MessageBox.Show(System.IO.Path.GetFullPath(OpenFileDialog1.FileName).ToString)
            TextBox1.Text = System.IO.Path.GetFullPath(OpenFileDialog1.FileName).ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con = db_con.getCon()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Create application instance.
        ' app As Application = New Application

        'check if Document is open
        If wordApplication Is Nothing Then
            wordApplication = New Microsoft.Office.Interop.Word.Application
            If wordDocument Is Nothing Then
                wordDocument = wordApplication.Documents.Open(OpenFileDialog1.FileName.ToString)
            Else
                wordDocument.Close()
            End If



            Dim Statistic As WdStatistic = WdStatistic.wdStatisticWords
            Dim missing = System.Reflection.Missing.Value

            Dim num = wordDocument.ComputeStatistics(Statistic, missing)
            Dim range As Range = wordDocument.Words(1)
            Dim count As Integer = wordDocument.Words.Count
            MessageBox.Show("Total words: " + count.ToString)
            Dim TextRange As Microsoft.Office.Interop.Word.Range = Nothing
            For Each sentence As Microsoft.Office.Interop.Word.Paragraph In wordDocument.Paragraphs
                TextRange = sentence.Range
                TextRange.Find.ClearFormatting()
                Dim par = sentence.Range.Text
                'RichTextBox1.Text = par
                Dim words() As String = par.Split(" ")

                For i As Integer = 0 To words.Length - 1
                    If words.Length >= 15 Then
                        Dim temp = RemoveWhitespace(words(i))
                        Dim wlength = temp.Length
                        Dim final As String = ""

                        For j As Integer = 0 To temp.Length - 1
                            If Not temp(j) = Nothing Then


                                '' -1 if white space
                                '' 0 DB Error handling like \ or '
                                '' 1 if Alpa Numeric
                                '' 2 if end of statement
                                '' 3 if division symbol only + / - *
                                If Not Char.IsWhiteSpace(temp(j)) Then
                                    If isValidChar(temp(j)) = 0 Then
                                        final += "\" + temp(j)
                                    ElseIf isValidChar(temp(j)) = 1 Then
                                        final += temp(j)
                                    End If
                                End If
                                

                            End If
                    'loadToDB(final)

                        Next

                        loadToDB(final)
                        ProgressBar1.Maximum = num
                        

                    End If

                Next

            Next
            
        Else
            MessageBox.Show("Document is already open!")
        End If
        ' Open specified file.



        wordDocument.Close()
        wordApplication.Quit()
        con.Close()

    End Sub

    Function RemoveWhitespace(fullString As String) As String
        Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
    End Function

    Function isValidChar(ByVal s As String)
        ' Dim b As Integer

        ''Database error handling
        If s.Contains("'") Or s.Contains("\") Then
            Return 0

            ''Sentences terminator
        ElseIf s.Contains(".") Or s.Contains("!") Or s.Contains("?") Or s.Contains(",") Then
            Return 2
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(s, "^[a-zA-Z0-9]+$") Then
            Return 1
        Else
            Return -1
        End If

    End Function

    Private Function isAlpha(ByVal letterChar As String) As Boolean
        Dim b As Integer

        '' -1 if white space
        '' 0 DB Error handling like \ or '
        '' 1 if Alpa Numeric
        '' 2 if end of statement
        '' 3 if division symbol. its for DB error handling

        If System.Text.RegularExpressions.Regex.IsMatch(letterChar, "^[a-zA-Z0-9]{1}$") Then
            b = 1
        ElseIf letterChar = "\" Or letterChar = "'" Then
            b = 0
        ElseIf letterChar = "." Or letterChar = "!" Or letterChar = "?" Or letterChar = "-" Then
            b = 2
        ElseIf letterChar = "+" Or letterChar = "-" Or letterChar = "*" Or letterChar = "=" Or letterChar = "/" Then
            b = 3
        Else
            b = -1
        End If
        Return b
    End Function

    Private Sub loadToDB(ByVal s As String)

        Try
            con.Open()
            RichTextBox1.Text = "Indexing " + s
            query = "SELECT COUNT(id) from `dictionary` WHERE word = '" & s & "'"
            cmd = New MySqlCommand(query, con)
            Dim wordCount As Integer = cmd.ExecuteScalar
            'MessageBox.Show("Count: " + isExist.ToString)
            If wordCount < 1 Then
                'Load to detabase if word not found
                query = "INSERT INTO `dictionary` (`id`, `word`) VALUES (NULL, '" & s & "')"
                cmd = New MySqlCommand(query, con)
                cmd.ExecuteNonQuery()
            Else
                'Index it
            End If
            ProgressBar1.Value = ProgressBar1.Value + 1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'wordDocument.Close()
            'wordApplication.Quit()
            'con.Close()
        Finally
            ' Quit the application.
            'wordDocument.Close()
            'wordApplication.Quit()
            con.Close()
        End Try
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim wordApplication As New Microsoft.Office.Interop.Word.Application
        Dim wordDocument As Microsoft.Office.Interop.Word.Document = Nothing
        'Dim outputFilename As String

        Try
            wordDocument = wordApplication.Documents.Open(OpenFileDialog1.FileName)
            'outputFilename = System.IO.Path.ChangeExtension(OpenFileDialog1.FileName, "pdf")

            'If Not wordDocument Is Nothing Then
            'wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, True, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
            'End If
        Catch ex As Exception
            'TODO: handle exception
        Finally
            If Not wordDocument Is Nothing Then
                wordDocument.Close(False)
                wordDocument = Nothing
            End If

            If Not wordApplication Is Nothing Then
                wordApplication.Quit()
                wordApplication = Nothing
            End If
        End Try
        MessageBox.Show("Done")

    End Sub
End Class
