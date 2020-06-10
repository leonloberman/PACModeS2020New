Module DBCompact

    Public Sub CompactDB()

        Dim acc As New ADOX.Catalog
        Dim datestring As String
        Dim DBName As String = "c:\ModeS\logged.mdb"
        Dim CompactedDBName As String = DBName & "_compacted"

        Dim jro = New JRO.JetEngine()

        Dim FileToDelete As String


        FileToDelete = My.Settings.CompactedDBName

        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)

        End If

        datestring = Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second
        CompactedDBName = CompactedDBName & "_" & datestring
        My.Settings.CompactedDBName = CompactedDBName


        jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "",
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CompactedDBName)

        FileToDelete = DBName

        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)
            My.Computer.FileSystem.RenameFile(CompactedDBName, "logged.mdb")

        End If


    End Sub

End Module
