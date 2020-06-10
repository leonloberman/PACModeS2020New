Imports System.Data.OleDb


Module SQLforAccess
    Public Sub AccessSQL(SQL As String, DBName As String)
        Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
        }
            Dim cmd2 As New OleDbCommand(SQL, con)

            Try
                'Write logfile
                'Using writer As StreamWriter = New StreamWriter(logfile, True)
                '    writer.WriteLine(Label1.Text & "-" & Date.Now)
                'End Using
                'SQL = "DELETE allhex.* FROM allhex;"
                con.Open()
                cmd2.ExecuteNonQuery()
                con.Dispose()
                con.Close()

            Catch ex As System.Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
        End Using
    End Sub
End Module

