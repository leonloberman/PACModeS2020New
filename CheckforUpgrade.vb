Imports System.Data.OleDb
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.Remoting

Module CheckforUpgrade

    Friend CurrentICAOversion As Integer
    Friend NewestICAOversion As String
    Friend NewestICAOversionsplit As String()
    Dim Lines1 As String()

    Friend NewestICAOversionvalue As Integer
    Public Property UpgradeText As String
    Public Property UpgradeURL As String

    Public Sub UpgradeCheck(DBName As String)
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim Request As HttpWebRequest = System.Net.HttpWebRequest.Create("https://www.gfiapac.org/ModeSVersions/PACModeS2020version.txt")
        Dim Response As HttpWebResponse = Request.GetResponse
        Dim Stream As Stream = Response.GetResponseStream()
        Dim username As String = "pacmodes2020"
        Dim password As String = "FkNrELRx"
        Dim localPath As String = "C:\ModeS\"
        Dim client As New WebClient
        Dim myCache As CredentialCache = New CredentialCache()

        Using SR As StreamReader = New StreamReader(Response.GetResponseStream)

            Dim StreamString As String = SR.ReadToEnd



            If StreamString.Length > 0 Then
                Lines1 = StreamString.Split(vbLf)

                'ICAOCodes.mdb version check
                Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                }
                Try
                    con.Open()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                End Try

                Dim cmdObj As New OleDbCommand("Select Version from VersionNo", con)

                Using con
                    Try
                        If con.State = ConnectionState.Closed Then con.Open()
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                    End Try
                    Using ICAOVersionRdr As OleDbDataReader = cmdObj.ExecuteReader
                        While ICAOVersionRdr.Read
                            CurrentICAOversion = ICAOVersionRdr("Version")
                        End While
                    End Using
                    con.Close()
                End Using

                Dim ICAOmdbresult As String() = Array.FindAll(Lines1, Function(s) s.Contains("ICAOCodes.mdb"))
                NewestICAOversion = ICAOmdbresult(0)
                NewestICAOversionsplit = NewestICAOversion.Split(":")
                NewestICAOversionvalue = Integer.Parse(NewestICAOversionsplit(1))

                'CurrentICAOversion = 3

                'Messaging section
                If NewestICAOversionvalue <> CurrentICAOversion Then
                    UpgradeText = "A new version of the ICAOCodes.mdb file is available - do you want to download it now?"
                    UpgradeURL = "https://www.gfiapac.org/members/ModeS/ICAOCodes.mdb"
                    Select Case MsgBox(UpgradeText, MsgBoxStyle.YesNo, "PACModeS2020 ICAOCodes Upgrade check")
                        Case MsgBoxResult.Yes


                            myCache.Add(
                            New System.Uri("https://www.gfiapac.org/"), "Basic",
                            New System.Net.NetworkCredential(username, password))

                            Stream = Response.GetResponseStream()
                            client.Credentials = myCache
                            PacModeS2020.Timer2.Start()
                            client.DownloadFile(UpgradeURL, localPath & "ICAOCodes.mdb")

                            Response.Close()

                            'Dim StartUpdate As New ProcessStartInfo(localPath & "ICAOCodes_Install.exe")
                            'Process.Start(StartUpdate)


                            Dim con2 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                }
                            Try
                                con2.Open()
                            Catch ex As Exception
                                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                            End Try

                            Dim cmdObj2 As New OleDbCommand("Select Version from VersionNo", con2)

                            Using con
                                Try
                                    If con2.State = ConnectionState.Closed Then con2.Open()
                                Catch ex As Exception
                                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                                End Try
                                Using ICAOVersionRdr2 As OleDbDataReader = cmdObj2.ExecuteReader
                                    While ICAOVersionRdr2.Read
                                        PacModeS2020.CurrentICAOversion = ICAOVersionRdr2("Version")
                                    End While
                                End Using
                                con2.Close()
                            End Using

                            'End
                        Case MsgBoxResult.No
                            'Carry on
                    End Select


                End If
            End If
        End Using

    End Sub
End Module
