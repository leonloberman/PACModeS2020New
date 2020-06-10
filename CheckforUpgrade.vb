Imports System.Data.OleDb
Imports System.IO
Imports System.Net

Module CheckforUpgrade

    Friend Currentloggedversion As Integer
    Friend CurrentICAOversion As Integer
    Friend Newestloggedversion As String
    Friend Newestloggedversionsplit As String()
    Friend NewestICAOversion As String
    Friend NewestICAOversionsplit As String()
    Dim Versions As String()
    Friend Newestappversion As String
    Friend Currentappversion As String = Application.ProductVersion
    Dim Lines1 As String()
    Dim Product As String

    Friend Newestloggedversionvalue As Integer
    Friend NewestICAOversionvalue As Integer





    Public Sub UpgradeCheck(DBName As String)
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim Request As HttpWebRequest = System.Net.HttpWebRequest.Create("https://www.gfiapac.org/ModeSVersions/PACModeS2020version.txt")
        Dim Response As HttpWebResponse = Request.GetResponse
        Dim Stream As Stream = Response.GetResponseStream()

        Using SR As StreamReader = New StreamReader(Response.GetResponseStream)

            Dim StreamString As String = SR.ReadToEnd



            If StreamString.Length > 0 Then
                Lines1 = StreamString.Split(vbLf)



                'App version check
                Dim result As String() = Array.FindAll(Lines1, Function(s) s.Contains(Application.ProductName))
                Product = result(0)
                Versions = Product.Split(":")
                Newestappversion = Versions(1)

                Dim UpgradeText As String = "Version " & Newestappversion & " is available - do you want to upgrade now?"
                Dim UpgradeFile As String = Replace(Newestappversion, ".", "")
                Dim UpgradeURL As String = "https://www.gfiapac.org/members/ModeS/PACModeS2020_v" & UpgradeFile & "_install.exe"

                'logged.mdb version check
                Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                }
                'Try
                '    Con.Open()
                'Catch ex As Exception
                '    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                'End Try

                Dim cmdObj As New OleDbCommand("Select Version from LoggedmdbVersionNo", con)
                Try
                    If con.State = ConnectionState.Closed Then con.Open()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                End Try
                Using VersionRdr As OleDbDataReader = cmdObj.ExecuteReader
                    While VersionRdr.Read()
                        Currentloggedversion = VersionRdr("Version")
                    End While
                    con.Close()
                End Using

                Dim loggedmdbresult As String() = Array.FindAll(Lines1, Function(s) s.Contains("logged.mdb"))
                Newestloggedversion = loggedmdbresult(0)
                Newestloggedversionsplit = Newestloggedversion.Split(":")
                Newestloggedversionvalue = Integer.Parse(Newestloggedversionsplit(1))


                'ICAOCodes.mdb version check
                con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                }
                Try
                    con.Open()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
                End Try

                cmdObj = New OleDbCommand("Select Version from ICAOCodesVersionNo", con)

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


                'Messaging section
                If Newestappversion = Currentappversion And Newestloggedversionvalue = Currentloggedversion And NewestICAOversionvalue = CurrentICAOversion Then
                    'Carry on
                ElseIf (Newestappversion) <> (Currentappversion) Or (Newestloggedversionvalue) <> (Currentloggedversion) Then
                    'MsgBox(UpgradeText, vbYesNo, "Upgrade check")
                    Select Case MsgBox(UpgradeText, MsgBoxStyle.YesNo, "PACModeS2020 Upgrade check")
                        Case MsgBoxResult.Yes
                            Dim sInfo As New ProcessStartInfo(UpgradeURL)
                            Process.Start(sInfo)
                            End
                        Case MsgBoxResult.No
                            'Carry on
                    End Select
                ElseIf NewestICAOversionvalue <> CurrentICAOversion Then
                    UpgradeText = "A new version of the ICAOCodes.mdb file is available - do you want to download it now?"
                    UpgradeURL = "https://www.gfiapac.org/members/ModeS/ICAOCodes_v" & NewestICAOversionvalue & "_Install.exe"
                    Select Case MsgBox(UpgradeText, MsgBoxStyle.YesNo, "PACModeS2020 ICAOCodes Upgrade check")
                        Case MsgBoxResult.Yes
                            Dim sInfo As New ProcessStartInfo(UpgradeURL)
                            Process.Start(sInfo)
                            End
                        Case MsgBoxResult.No
                            'Carry on
                    End Select


                End If
            End If
        End Using

    End Sub
End Module
