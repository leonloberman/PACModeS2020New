Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.IO
Imports System.Data.SQLite
Imports AutoUpdaterDotNET
Imports System.Net

Public Class PacModeS2020
    Private Const ConnectString As String = "Provider=System.Data.SQLite;DataSource=:memory:;New=True;"
    Public Con As OleDbConnection
    Public ReadOnly DBName As String = "C:\Modes\logged.mdb"
    Public ReadOnly ICAOName As String = "C:\Modes\ICAOCodes.mdb"
    Private ReadOnly CompactedDBName As String
    Private BSLoc As String
    Private BSBackupLoc As String
    Private ReadOnly BSTempLoc As String
    Private PACBSLoc As String
    Private ReadOnly EDITMODE As Boolean = False
    Private ReadOnly NEWMODE As Boolean = False
    Private ReadOnly SQL As String
    Private WithEvents BGW As New BackgroundWorker

    Private ReadOnly fd As OpenFileDialog = New OpenFileDialog()
    Private ReadOnly strFileName As String
    Private ReadOnly backupfilename As String = "basestation.sqb.copy." + String.Format("{0:yyyyMMdd_HHmmss}", Date.Now)
    Private ReadOnly BSTempfilename As String = "basestation.sqb.tmp." + String.Format("{0:yyyyMMdd_HHmmss}", Date.Now)
    Private ReadOnly BSFileName As String
    Private ReadOnly logfileseq As String
    Private ReadOnly BS_Con = New SQLiteConnection
    Private ReadOnly BS_Con_mem = New SQLiteConnection(ConnectString1)
    Private ReadOnly Logged_con As OleDbConnection
    Private ReadOnly da1 As New SQLiteDataAdapter
    Private ReadOnly da4 As New SQLiteDataAdapter
    Private ReadOnly ds1 As New DataSet("ACLocal")
    Private ReadOnly ds2 As New DataSet("AllHexLocal")
    Private ReadOnly ds3 As New DataSet
    Private ReadOnly cmdBuilder As SQLiteCommandBuilder = New SQLiteCommandBuilder(da1)
    Private ReadOnly dt1 As DataTable
    Private ReadOnly logged_cmd As OleDbCommand
    Private ReadOnly dtb As New DataTable
#Disable Warning IDE0044 ' Add readonly modifier
    Private dr2 As DataRow
#Enable Warning IDE0044 ' Add readonly modifier
    Private dtc As New DataTable
    Private SQLTrans As SQLiteTransaction
    Private ReadOnly dtopflags As New DataTable
    Private ReadOnly OperatorFlags As String

    Private FileToDelete As String
    Private FileToRename As String

    Private bRet As Boolean = False

    Private ReadOnly fileReader As String
    Private ReadOnly DBHistory_Text As String
    Public Property BS_SQL1 As String

    Public Currentloggedversion As Integer
    Public CurrentICAOversion As Integer
    Public AutoUpdaterFile As String = "https://www.gfiapac.org/ModeSVersions/PACModeS2020Version.xml"
    Private LatestGFIAUpdate As Integer


    Public Shared ReadOnly Property ConnectString1 As String
        Get
            Return ConnectString2
        End Get
    End Property

    Public Shared ReadOnly Property ConnectString2 As String
        Get
            Return ConnectString
        End Get
    End Property


    Private Sub PACModes2020_Load(sender As Object, e As EventArgs) Handles MyBase.Load
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
        If My.Settings.Default.UpgradeRequired Then
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.Upgrade()
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.UpgradeRequired = False
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
#Disable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
            My.Settings.Default.Save()
#Enable Warning BC42025 ' Access of shared member, constant member, enum member or nested type through an instance
        End If

        'get logged.mdb version
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                }

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

        'get ICAOCodes.mdb version
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

        'get GFIA latest update version
        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "") With {
                            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
                        }
        Try
            con.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
        End Try

        cmdObj = New OleDbCommand("Select MsSysUpdate from MsSysBuilder", con)

        Using con
            Try
                If con.State = ConnectionState.Closed Then con.Open()
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
            End Try
            Using GFIAUpdateVersionRdr As OleDbDataReader = cmdObj.ExecuteReader
                While GFIAUpdateVersionRdr.Read
                    LatestGFIAUpdate = GFIAUpdateVersionRdr("MsSysUpdate")
                End While
            End Using
            con.Close()
        End Using

        If My.Computer.Network.IsAvailable Then
            If My.Computer.Network.Ping("www.google.com") Then
                'Application Upgrade Check
                Dim BasicAuthentication As BasicAuthentication = New BasicAuthentication("pad", "Blackmrs99")
                AutoUpdater.BasicAuthXML = BasicAuthentication
                AutoUpdater.ReportErrors = True
                AutoUpdater.ShowSkipButton = False
                'AutoUpdater.Mandatory = True
                'AutoUpdater.Synchronous = True
                AutoUpdater.UpdateFormSize = New System.Drawing.Size(800, 600)
                AutoUpdater.Start(AutoUpdaterFile)


                AddHandler AutoUpdater.CheckForUpdateEvent, AddressOf AutoUpdaterOnCheckForUpdateEvent
            End If
        Else
            'MsgBox("Computer is not connected to the internet.")
        End If

        Label7.Text += My.Application.Info.Version.ToString
        Label8.Text += Currentloggedversion.ToString
        Label3.Text += CurrentICAOversion.ToString
        Label10.Text += LatestGFIAUpdate.ToString
        TextBox1.Text = My.Settings.BSloc
        TextBox2.Text = My.Settings.BSBackupLoc
        BSLoc = My.Settings.BSloc
        BSBackupLoc = My.Settings.BSBackupLoc
        If My.Settings.InterestedButton = True Then
            RadioButton2.Checked = True
        ElseIf My.Settings.RQPsButton = True Then
            RadioButton1.Checked = True
        ElseIf My.Settings.RQPsandIButton = True Then
            RadioButton10.Checked = True
        End If
        If My.Settings.PPSymbols = True Then
            CheckBox1.Checked = True
            If My.Settings.PPSymbolsType = "v1" Then
                RadioButton7.Checked = True
            ElseIf My.Settings.PPSymbolsType = "v3" Then
                RadioButton8.Checked = True
            End If
        Else CheckBox1.Checked = False
        End If
        If My.Settings.OperatorFlags = "Kinetic" Then
            RadioButton5.Checked = True
        ElseIf My.Settings.OperatorFlags = "GFIA" Then
            RadioButton6.Checked = True
        Else
            RadioButton9.Checked = True
        End If
        If My.Settings.NullOpFlags = True Then
            CheckBox3.Checked = True
        Else
            CheckBox3.Checked = False
        End If


    End Sub

    Private Sub AutoUpdaterOnCheckForUpdateEvent(ByVal args As UpdateInfoEventArgs)
        If args IsNot Nothing Then

            If args.IsUpdateAvailable Then

                AutoUpdater.ShowUpdateForm(args)

            Else
                'MessageBox.Show("There is no update available please try again later.", "No update available", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'If My.Computer.Network.Ping("www.google.com") Then
                UpgradeCheck("C:\ModeS\ICAOCodes.mdb")
                'Else
                '    MsgBox("Computer is not connected to the internet.")
                'End If
            End If
        Else
            MessageBox.Show("There is a problem reaching update server please check your internet connection and try again later.", "Update check failed", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End If
    End Sub


    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        ProgressBar2.Visible = True
        ProgressBar2.Value += 10
        If ProgressBar2.Value = 100 Then
            Timer2.Stop()
            ProgressBar2.Visible = False
            Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ICAOName & "") With {
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ICAOName & ""
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
            Label3.Text = "ICAOCodes Version: " + CurrentICAOversion.ToString
            Label9.Visible = True
            'End If
        End If

    End Sub

    Public Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        BGW.WorkerSupportsCancellation = True
        BGW.WorkerReportsProgress = True

        If TextBox1.TextLength = 0 Then
            MsgBox("You must enter the location of your BaseStation.sqb file", vbExclamation, "BS Location Check")
            Button3.Enabled = True
            Button4.Enabled = True
            Exit Sub
        Else

            BSLoc = My.Settings.BSloc & "\BaseStation.sqb"
            PACBSLoc = My.Settings.BSloc

        End If

        If TextBox2.TextLength = 0 Then
            MsgBox("You must enter a location for your pre-update copy of your BaseStation.sqb file", vbExclamation, "BS Backup Location Check")
            Button3.Enabled = True
            Button4.Enabled = True
            Exit Sub
        Else

            BSLoc = My.Settings.BSloc & "\BaseStation.sqb"
            PACBSLoc = My.Settings.BSloc

        End If

        If Not BGW.IsBusy = True Then
            ' Disable the Start button
            Button3.Enabled = False
            ' Enable to Cancel button
            Button4.Enabled = True
            ' Start the Background Worker working
            BGW.RunWorkerAsync()

        End If

        ProgressBar1.Refresh()
        ProgressBar1.Style = ProgressBarStyle.Marquee

    End Sub
    Public Sub BGW_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles BGW.DoWork

        Static start_time As DateTime
        Static stop_time As DateTime
        Dim elapsed_time As TimeSpan
        Dim elapsed As String

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        AddHandler BGW.DoWork, AddressOf BGW_DoWork

        BS_Con.ConnectionString = "Provider=System.Data.SQLite;Data Source=" & BSLoc & "" & ";PRAGMA cache_size = -10000;"

        Con = New OleDbConnection With {
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & ""
        }
        Try
            'Error handling keeps our software from crashing
            'when an error occurs
            Con.Open()
        Catch ex As Exception
            'You could opt to do something here
            'This code is called when an error occurs
            'I usually do a MsgBox
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Connection Error")
        End Try

        System.Threading.Thread.Sleep(500)

        Dim daBSCommand = New SQLiteCommand

        start_time = Now
        SetLabelText_ThreadSafe(Label6, vbCrLf + "Start time" & Chr(32) & start_time, Color.Blue, 0)

BSBackupStep:

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Backing up BaseStation.sqb", Color.Yellow, 0)

        BSBackupLoc = Path.Combine(My.Settings.BSBackupLoc, backupfilename)
        BSLoc = Path.Combine(My.Settings.BSloc, "basestation.sqb")

        File.Copy(BSLoc, BSBackupLoc)

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton4.Checked = True Then

            GoTo QuickUpdate

        End If

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing tables", Color.Yellow, 0)
        AccessSQL("DELETE OperatorFlags.* From OperatorFlags", DBName)
        AccessSQL("DELETE Kinetic_Operator_Flags.* From Kinetic_Operator_Flags", DBName)
        AccessSQL("Delete Allhex.* from Allhex;", DBName)
        AccessSQL("DELETE loggedhex.* FROM loggedhex;", DBName)
        AccessSQL("DELETE Duplicate_Hex_from_allhex.* FROM Duplicate_Hex_from_allhex;", DBName)
        AccessSQL("DELETE Ps_reset.* FROM Ps_reset;", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton6.Checked = True Then
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Getting hex data from GFIA and setting GFIA flags", Color.Yellow, 0)
            'Set GFIA flags
            AccessSQL("INSERT INTO allhex ( AircraftID, ModeS, ModeSCountry, Registration, ICAOTypecode, SerialNo, UserInt1, RegisteredOwners, OperatorFlagCode )" &
                    "SELECT tbldataset.ID, tbldataset.Hex, tblCountry.CountryName, tbldataset.Registration, tblSeries.code, tbldataset.CN, tbldataset.FKCMXO, PRO_tbloperator.Operator,  Str([PRO_tbloperator].[FKoperator]) AS OperatorFlagCode" &
                    " FROM PRO_tbloperator RIGHT JOIN (tblSeries INNER JOIN (tblCountry RIGHT JOIN tbldataset ON tblCountry.FKcountry = tbldataset.FKcountry) On tblSeries.FKseries = tbldataset.FKseries) On PRO_tbloperator.FKoperator = tbldataset.FKoperator" &
                    " where tbldataset.hex Is Not NULL Or tbldataset.hex <> " & """"";", DBName)

            AccessSQL("UPDATE allhex Set OperatorFlagCode = 'x'+TRIM([OperatorFlagCode]);", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        ElseIf RadioButton5.Checked = True Then
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Getting hex data from GFIA", Color.Yellow, 0)
            'Set Kinetic flags
            AccessSQL("INSERT INTO allhex ( AircraftID, ModeS, ModeSCountry, Registration, ICAOTypecode, SerialNo, UserInt1, RegisteredOwners)" &
                            "SELECT tbldataset.ID, tbldataset.Hex, tblCountry.CountryName, tbldataset.Registration, tblSeries.code, tbldataset.CN, tbldataset.FKCMXO, PRO_tbloperator.Operator" &
                            " FROM PRO_tbloperator RIGHT JOIN (tblSeries INNER JOIN (tblCountry RIGHT JOIN tbldataset ON tblCountry.FKcountry = tbldataset.FKcountry) ON tblSeries.FKseries = tbldataset.FKseries) ON PRO_tbloperator.FKoperator = tbldataset.FKoperator" &
                            " where tbldataset.hex is not NULL or tbldataset.hex <> " & """"";", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        ElseIf RadioButton9.Checked = True Then
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Getting hex data from GFIA - personal flags", Color.Yellow, 0)
            'Set Kinetic flags
            AccessSQL("INSERT INTO allhex ( AircraftID, ModeS, ModeSCountry, Registration, ICAOTypecode, SerialNo, UserInt1, RegisteredOwners)" &
                            "SELECT tbldataset.ID, tbldataset.Hex, tblCountry.CountryName, tbldataset.Registration, tblSeries.code, tbldataset.CN, tbldataset.FKCMXO, PRO_tbloperator.Operator" &
                            " FROM PRO_tbloperator RIGHT JOIN (tblSeries INNER JOIN (tblCountry RIGHT JOIN tbldataset ON tblCountry.FKcountry = tbldataset.FKcountry) ON tblSeries.FKseries = tbldataset.FKseries) ON PRO_tbloperator.FKoperator = tbldataset.FKoperator" &
                            " where tbldataset.hex is not NULL or tbldataset.hex <> " & """"";", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        End If

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Deleting records with no hex codes", Color.Yellow, 0)
        AccessSQL("DELETE allhex.*, allhex.Modes FROM allhex WHERE (((allhex.ModeS) Is Null Or (allhex.ModeS)=""""" & "));", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Finding duplicate hex codes", Color.Yellow, 0)
        AccessSQL("INSERT INTO Duplicate_Hex_from_allhex SELECT allhex.AircraftID AS ID, allhex.ModeS AS Hex, allhex.ModeSCountry AS CountryName, allhex.Registration AS Registration, allhex.ICAOTypecode AS code, allhex.SerialNo AS CN FROM allhex" &
        " WHERE (((allhex.ModeS) In (SELECT [ModeS] FROM [allhex] As Tmp GROUP  BY [ModeS] HAVING Count(*)>1 ))) ORDER BY allhex.ModeS;", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Delete duplicate hex codes", Color.Yellow, 0)
        AccessSQL("DELETE * FROM allhex WHERE EXISTS " &
        "(select * from Duplicate_Hex_from_allhex where allhex.ModeS = Duplicate_Hex_from_allhex.Hex);", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Getting hex data for logged records", Color.Yellow, 0)
        AccessSQL("INSERT INTO loggedhex ( ID, Registration, Hex ) SELECT logllp.ID, logllp.Registration, tbldataset.Hex" &
                  " FROM (logllp INNER JOIN (SELECT ID, Max([when]) as LastDate" &
                  " FROM logllp GROUP BY ID)  AS B ON (logllp.ID = B.ID) AND (logllp.[when] = B.LastDate)) INNER JOIN tbldataset ON B.ID = tbldataset.ID;", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Deleting logged records with no hex data", Color.Yellow, 0)
        AccessSQL("DELETE loggedhex.* FROM loggedhex WHERE (((loggedhex.Hex) Is Null Or (loggedhex.Hex)=""""" & "));", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing Interested fields", Color.Yellow, 0)
        AccessSQL("Update allhex SET Interested = FALSE WHERE Interested = True;", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton2.Checked = True Then

            If CheckBox2.Checked = True Then
                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field", Color.Yellow, 0)
                AccessSQL("UPDATE Allhex LEFT JOIN logLLp ON Allhex.[AircraftID] = logLLp.[ID]" &
                             " SET Allhex.Interested = True, Allhex.LastModified = Now()" &
                             " WHERE (((logLLp.ID)=[Allhex].[AircraftID]) AND ((logLLp.Registration)=[Allhex].[Registration]));", DBName)
                'SetLabelText_ThreadSafe(Me.Label2, vbCrLf + "Completed", Color.Green, 0)

                'Removing Ps if previously logged with this registration (i.e. repeated lease) Part 1
                AccessSQL("INSERT INTO Ps_Reset ( Registration, AircraftID, Hex )" &
                          " SELECT DISTINCT logLLp.Registration,  logLLp.ID AS AircraftID, tbldataset.Hex" &
                          " FROM ((logLLp INNER JOIN tblOperatorHistory ON logLLp.ID = tblOperatorHistory.ID) INNER JOIN tbldataset ON logLLp.ID = tbldataset.ID)" &
                          " INNER JOIN Allhex ON logLLp.ID = Allhex.AircraftID" &
                          " WHERE (((logLLp.Registration) In (select tblOperatorHistory.previous from tblOperatorHistory)));", DBName)

                'Removing Ps if previously logged with this registration (i.e. repeated lease) Part 2
                AccessSQL("UPDATE Allhex INNER JOIN Ps_Reset ON Allhex.AircraftID = Ps_Reset.AircraftID" &
                          " SET allhex.Interested = True, allhex.LastModified = Now()" &
                          " Where allhex.AircraftID in (select AircraftID from Ps_Reset);", DBName)
                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
            Else

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field", Color.Yellow, 0)
                AccessSQL("UPDATE Allhex LEFT JOIN logLLp ON Allhex.[AircraftID] = logLLp.[ID]" &
                             " SET Allhex.Interested = True, Allhex.LastModified = Now()" &
                             " WHERE (((logLLp.ID)=[Allhex].[AircraftID]) AND ((logLLp.Registration)=[Allhex].[Registration]));", DBName)
                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field Part 2", Color.Yellow, 0)
                AccessSQL("UPDATE Allhex INNER JOIN loggedhex ON (loggedhex.ID = Allhex.AircraftID) SET Allhex.Interested = FALSE, Allhex.LastModified = Now()" &
                            " WHERE (((Allhex.AircraftID)=[loggedhex].[ID]) AND ((Allhex.Registration)<>[loggedhex].[Registration]));", DBName)
                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field Part 3", Color.Yellow, 0)
                AccessSQL("INSERT INTO Ps_Reset ( AircraftID, Registration, Hex )" &
                      " SELECT DISTINCT logLLp.ID AS AircraftID, logLLp.Registration, tbldataset.Hex" &
                      " FROM (logLLp INNER JOIN tblOperatorHistory ON (logLLp.ID = tblOperatorHistory.ID) AND (logLLp.Registration = tblOperatorHistory.previous))" &
                      " INNER JOIN tbldataset ON logLLp.ID = tbldataset.ID" &
                      " WHERE (((logLLp.Registration)<>([tbldataset].[registration]) And (logLLp.Registration) In (select tblOperatorHistory.previous from tblOperatorHistory)));", DBName)
                AccessSQL("UPDATE Allhex INNER JOIN Ps_Reset ON Allhex.AircraftID = Ps_Reset.AircraftID" &
                      " SET allhex.Interested = False, allhex.LastModified = Now()" &
                      " Where allhex.Registration in (select registration from Ps_Reset);", DBName)
                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

            End If

        End If

        If RadioButton1.Checked = True Or RadioButton10.Checked = True Then

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing User Tag fields", Color.Yellow, 0)
            AccessSQL("Update Allhex SET UserTag = " & """RQ""" & " WHERE (UserTag = " & """Ps""" & ")" &
                    " OR (UserTag IS NULL);", DBName)
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Updating UserTag field only", Color.Yellow, 0)
            AccessSQL("UPDATE Allhex LEFT JOIN logLLp ON Allhex.[AircraftID] = logLLp.[ID]" &
                             " SET Allhex.UserTag = null, Allhex.LastModified = Now()" &
                             " WHERE (((logLLp.ID)=[Allhex].[AircraftID]) AND ((logLLp.Registration)=[Allhex].[Registration]));", DBName)
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

            If RadioButton10.Checked = False Then
                SetLabelText_ThreadSafe(Label1, vbCrLf + "Updating Interested field for Mil only", Color.Yellow, 0)
                AccessSQL("UPDATE Allhex SET Interested = TRUE, LastModified = Now()" &
                             " WHERE UserInt1 = 502 ;", DBName)
                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
            End If


            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting UserTag = Ps as appropriate", Color.Yellow, 0)
            AccessSQL("UPDATE Allhex INNER JOIN loggedhex ON (loggedhex.ID = Allhex.AircraftID) SET Allhex.UserTag = " & """Ps""" & ", Allhex.LastModified = Now()" &
                        " WHERE (((Allhex.AircraftID)=[loggedhex].[ID]) And ((Allhex.Registration)<>[loggedhex].[Registration]));", DBName)

            'Removing Ps if previously logged with this registration (i.e. repeated lease) Part 1
            AccessSQL("INSERT INTO Ps_Reset ( Registration, AircraftID, Hex )" &
                      " SELECT DISTINCT logLLp.Registration,  logLLp.ID AS AircraftID, tbldataset.Hex" &
                      " FROM ((logLLp INNER JOIN tblOperatorHistory ON logLLp.ID = tblOperatorHistory.ID) INNER JOIN tbldataset ON logLLp.ID = tbldataset.ID)" &
                      " INNER JOIN Allhex ON logLLp.ID = Allhex.AircraftID" &
                      " WHERE (((logLLp.Registration) In (select tblOperatorHistory.previous from tblOperatorHistory)));", DBName)

            'Removing Ps if previously logged with this registration (i.e. repeated lease) Part 2
            AccessSQL("UPDATE Allhex INNER JOIN Ps_Reset ON Allhex.AircraftID = Ps_Reset.AircraftID" &
                      " SET allhex.UserTag = NULL, allhex.LastModified = Now()" &
                      " Where allhex.Registration in (select registration from Ps_Reset);", DBName)
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
        End If

        'Update Allhex with correct Airbus NEO and B737Max ICAO Type codes
        SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting correct ICAO Type codes for Airbus NEO family", Color.Yellow, 0)
        AccessSQL("UPDATE Allhex INNER JOIN tbldataset ON Allhex.AircraftID = tbldataset.ID " &
                  "SET Allhex.ICAOTypeCode = 'A20N' WHERE (((tbldataset.FKvariant) in (16897, 9000539)));", DBName)
        AccessSQL("UPDATE Allhex INNER JOIN tbldataset ON Allhex.AircraftID = tbldataset.ID " &
                  "SET Allhex.ICAOTypeCode = 'A21N' WHERE (((tbldataset.FKvariant) in (9000540, 9000541, 9000977)));", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting correct ICAO Type codes for Boeing 737 Max family", Color.Yellow, 0)
        AccessSQL("UPDATE Allhex INNER JOIN tbldataset ON Allhex.AircraftID = tbldataset.ID " &
                  "SET Allhex.ICAOTypeCode = 'B37M' WHERE (((tbldataset.FKvariant) in (9001052)));", DBName)
        AccessSQL("UPDATE Allhex INNER JOIN tbldataset ON Allhex.AircraftID = tbldataset.ID " &
                  "SET Allhex.ICAOTypeCode = 'B38M' WHERE (((tbldataset.FKvariant) in (9001053)));", DBName)
        AccessSQL("UPDATE Allhex INNER JOIN tbldataset ON Allhex.AircraftID = tbldataset.ID " &
                  "SET Allhex.ICAOTypeCode = 'B39M' WHERE (((tbldataset.FKvariant) in (9000980)));", DBName)
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton10.Checked = True Then
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting all RQ/Ps to Interested", Color.Yellow, 0)
            AccessSQL("UPDATE Allhex SET Interested = TRUE, LastModified = Now()" &
                             " WHERE (UserTag = " & """Ps""" & ")" &
                    " OR (UserTag = " & """RQ""" & ");", DBName)
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        End If

        If CheckBox1.Checked = True Then
            Select Case RadioButton7.Checked
                Case True
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Loading PlanePlotter v1 Symbols", Color.Yellow, 0)

                    AccessSQL("UPDATE Allhex INNER JOIN PP_SymbolsByType On Allhex.ICAOTypeCode = PP_SymbolsByType.ICAOTypeCode Set" &
                          " allhex.UserString1 = PP_SymbolsByType.UserString1_v1;", DBName)
                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
                Case Else
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Loading PlanePlotter v3 Symbols", Color.Yellow, 0)
                    AccessSQL("UPDATE Allhex INNER JOIN PP_SymbolsByType On Allhex.ICAOTypeCode = PP_SymbolsByType.ICAOTypeCode Set" &
                          " allhex.UserString1 = PP_SymbolsByType.UserString1_v3;", DBName)
                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
            End Select


            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting PlanePlotter Symbols", Color.Yellow, 0)
            AccessSQL("UPDATE Allhex Set Allhex.UserTag = Allhex.UserTag & AllHex.UserString1;", DBName)
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        End If

        If RadioButton5.Checked = True Then

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Kinetic flags Part 1", Color.Yellow, 0)

            AccessSQL("INSERT INTO Kinetic_Operator_Flags Select Allhex.AircraftID, PRO_tbloperator.Operator, PRO_tbloperator.[3L]" &
                            "FROM (PRO_tbloperator RIGHT JOIN tbldataset ON PRO_tbloperator.FKoperator = tbldataset.FKoperator)" &
                            "INNER JOIN Allhex On tbldataset.ID = Allhex.AircraftID;", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Kinetic flags Part 2", Color.Yellow, 0)

            AccessSQL("UPDATE Allhex INNER JOIN Kinetic_Operator_Flags On Allhex.AircraftID = Kinetic_Operator_Flags.AircraftID Set Allhex.OperatorFlagCode = Kinetic_Operator_Flags.[3L]" &
                            "WHERE (((Kinetic_Operator_Flags.[3L])<>'')" &
                            "AND ((Kinetic_Operator_Flags.AircraftID)=[Allhex].[AircraftID]));", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
        End If

        If CheckBox3.Checked = True Then
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting No Null Operator Flags", Color.Yellow, 0)

            AccessSQL("UPDATE Allhex Set Allhex.OperatorFlagCode = ICAOTypeCode WHERE" &
                      " (Allhex.OperatorFlagCode = '-' or Allhex.OperatorFlagCode IS NULL);", DBName)

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
        End If

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Building Type List", Color.Yellow, 0)

        AccessSQL("SELECT DISTINCT tblManufacturer.Builder+' '+tblmodel.Model AS Types INTO TypeList
                    FROM (tblmodel INNER JOIN tblManufacturer ON tblmodel.UID = tblManufacturer.UID) INNER JOIN tbldataset ON (tblmodel.FKmodel = tbldataset.FKmodel) AND (tblManufacturer.UID = tbldataset.UID)
                    ORDER BY tblManufacturer.Builder+' '+tblmodel.Model", DBName)

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Building Operator List", Color.Yellow, 0)

        AccessSQL("SELECT PRO_tbloperator.FKoperator, PRO_tbloperator.Operator INTO OperatorList
                    FROM PRO_tbloperator ORDER BY PRO_tbloperator.Operator", DBName)

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        'Open connection to BaseStation
        Try
            'Error handling keeps our software from crashing
            'when an error occurs
            BS_Con.Open()
            BS_Con_mem.Open()
            'MsgBox("Connection Success", MsgBoxStyle.OkOnly)
        Catch ex As Exception
            'You could opt to do something here
            'This code is called when an error occurs
            'I usually do a MsgBox
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "BaseStation Connection Error")
        End Try

        Try

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Attach Current BaseStation database", Color.Yellow, 0)
            BS_SQL1 = "ATTACH DATABASE '" + BSLoc + "' AS currentDB"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Create other tables in memory", Color.Yellow, 0)

        Try
            BS_SQL1 = "CREATE TABLE [DBHistory] (" &
                        "[DBHistoryID] integer primary key,[TimeStamp] datetime Not null, [Description] varchar(100) Not null);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into DBHistory Select * From currentDB.DBHistory;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        'Try
        '    DBHistory_Text = "Created by PACModes2020 V" & My.Application.Info.Version.ToString
        '    BS_SQL1 = "Insert into DBHistory (TimeStamp, Description) values ( Current_TIMESTAMP, '" + DBHistory_Text + " ');"
        '    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
        '    daBSCommand.ExecuteNonQuery()
        'Catch ex As System.Exception
        '    System.Windows.Forms.MessageBox.Show(ex.Message)
        'End Try

        Try
            BS_SQL1 = "CREATE TABLE DBInfo(OriginalVersion smallint not null,CurrentVersion smallint not null);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into DBInfo Select * From currentDB.DBInfo;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TABLE Flights(FlightID integer primary key,SessionID integer not null,
                    AircraftID integer not null,StartTime datetime not null,EndTime datetime,
                    Callsign varchar(20),NumPosMsgRec integer,NumADSBMsgRec integer,NumModeSMsgRec integer,
                    NumIDMsgRec integer,NumSurPosMsgRec integer,NumAirPosMsgRec integer,NumAirVelMsgRec integer,
                    NumSurAltMsgRec integer,NumSurIDMsgRec integer,NumAirToAirMsgRec integer,NumAirCallRepMsgRec integer,
                    FirstIsOnGround boolean not null default 0,LastIsOnGround boolean not null default 0,FirstLat real,
                    LastLat real,FirstLon real,LastLon real,FirstGroundSpeed real,LastGroundSpeed real,
                    FirstAltitude integer,LastAltitude integer,FirstVerticalRate integer,LastVerticalRate integer,
                    FirstTrack real,LastTrack real,FirstSquawk integer,LastSquawk integer,HadAlert boolean not null default 0,
                    HadEmergency boolean not null default 0,HadSPI boolean not null default 0,UserNotes varchar(300),
                    CONSTRAINT SessionIDfk FOREIGN KEY (SessionID) REFERENCES Sessions,
                    CONSTRAINT AircraftIDfk FOREIGN KEY (AircraftID) REFERENCES Aircraft);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [FlightsStartTime] ON [Flights] ([StartTime]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [FlightsSessionID] ON [Flights] ([SessionID]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [FlightsEndTime] ON [Flights] ([EndTime]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [FlightsAircraftID] ON [Flights] ([AircraftID]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [FlightsCallsign] ON [Flights] ([Callsign]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into Flights Select * From currentDB.Flights;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TABLE Locations(LocationID integer primary key,LocationName varchar(20) not null,
                        Latitude real not null,Longitude real not null,Altitude real not null);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [LocationsLocationName] ON [Locations] ([LocationName]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into Locations Select * From currentDB.Locations;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TABLE Sessions(SessionID integer primary key,LocationID integer not null,
                    StartTime datetime not null,EndTime datetime,CONSTRAINT LocationIDfk FOREIGN KEY (LocationID) REFERENCES Locations);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [SessionsEndTime] ON [Sessions] ([EndTime]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [SessionsLocationID] ON [Sessions] ([LocationID]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [SessionsStartTime] ON [Sessions] ([StartTime]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TRIGGER SessionIDdeltrig BEFORE DELETE ON Sessions " &
                " FOR EACH ROW BEGIN DELETE FROM Flights WHERE SessionID = OLD.SessionID;END;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into Sessions Select * From currentDB.Sessions;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TABLE SystemEvents(SystemEventsID integer primary key,
                TimeStamp datetime not null,App varchar(15) not null,Msg varchar(100) not null);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [SystemEventsApp] ON [SystemEvents] ([App]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [SystemEventsTimeStamp] ON [SystemEvents] ([TimeStamp]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "Insert into SystemEvents Select * From currentDB.SystemEvents;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "SELECT * from sqlite_master where name = 'Alerts';"
            Using BSCommand As SQLiteCommand = New SQLiteCommand(BS_SQL1, BS_Con)
                Using reader As SQLiteDataReader = BSCommand.ExecuteReader
                    bRet = reader.HasRows
                End Using
            End Using

            If bRet = True Then
                Try
                    BS_SQL1 = "CREATE TABLE Alerts (ID integer primary key, AlertType char(1) null, AlertValue char(10) null);"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "CREATE INDEX [AlertsAlertValue] ON [Alerts] ([AlertValue]);
                                CREATE INDEX AlertsAlertType ON Alerts(AlertType);"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "Insert into Alerts Select * From currentDB.Alerts;"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                bRet = False

            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "SELECT * from sqlite_master where name = 'Active';"
            Using BSCommand As SQLiteCommand = New SQLiteCommand(BS_SQL1, BS_Con)
                Using reader As SQLiteDataReader = BSCommand.ExecuteReader
                    bRet = reader.HasRows
                End Using
            End Using

            If bRet = True Then
                Try
                    BS_SQL1 = "CREATE TABLE Active (ID integer primary key, ModeS varchar(7) not null, AircraftID integer,
                                FlightID integer not null, SessionID integer not null, FirstCreated DateTime not null,
                                LastModified DateTime not null, Country varchar(24) null, Registration varchar(20) null,
                                Callsign varchar(20) null, Type varchar(50) null, ICAOType varchar(8) null,
                                ConstructionNumber varchar(30) null, Operator varchar(100) null, ICAOOperator varchar(20) null,
                                SubOperator varchar(20) null, RadioCallsign varchar(20) null, Route varchar(12) null,
                                UserTag varchar(5) null, Interested boolean default 0, Populated boolean default 0,
                                Alert boolean default 0, FirstAltitude integer default 0, LastAltitude integer default 0,
                                FirstLongitude real null, LastLongitude real null, FirstLatitude real null, LastLatitude real null,
                                FirstGroundSpeed real null, LastGroundSpeed real null, FirstVerticalRate integer default 0,
                                LastVerticalRate integer default 0, FirstTrack real null, LastTrack real null,
                                FirstSquawk integer default 0, LastSquawk integer default 0, UserString1 varchar(20) null,
                                UserString2 varchar(20) null, UserString3 varchar(20) null, UserString4 varchar(20) null,
                                UserString5 varchar(20) null, UserInt1 integer default 0, UserInt2 integer default 0,
                                UserInt3 integer default 0, UserInt4 integer default 0, UserInt5 integer default 0,
                                UserBool1 boolean default 0, UserBool2 boolean default 0, UserBool3 boolean default 0,
                                UserBool4 boolean default 0, UserBool5 boolean default 0, UserNotes varchar(300) null,
                                NeedsPopulating boolean default 0, NewEntry boolean default 0, Miscode boolean default 0,
                                GroundCode boolean default 0, Currency boolean default 0, DataCurrency varchar(40) null,
                                MiscodedRegistration varchar(20) null, PreviousRegistration varchar(20) null,
                                PreviousCode varchar(7) null, SquawkTrans varchar(50) null, BizJet boolean default 0,
                                BizProp boolean default 0, LightAircraft boolean default 0, Helicopter boolean default 0,
                                UserType boolean default 0, DeletedRecord boolean default 0, PressureSetting integer default 1013);"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "CREATE INDEX ActiveInterested ON Active(Interested);

                                CREATE INDEX ActiveModeS ON Active(ModeS);

                                CREATE INDEX ActiveRadioCallsign ON Active(RadioCallsign);

                                CREATE INDEX ActivePreviousCode ON Active(PreviousCode);

                                CREATE INDEX ActiveGroundCode ON Active(GroundCode);

                                CREATE INDEX ActiveUserType ON Active(UserType);

                                CREATE INDEX ActiveCallsign ON Active(Callsign);

                                CREATE INDEX ActiveNewEntry ON Active(NewEntry);

                                CREATE INDEX ActiveCurrency ON Active(Currency);

                                CREATE INDEX ActiveBizProp ON Active(BizProp);

                                CREATE INDEX ActiveDataCurrency ON Active(DataCurrency);

                                CREATE INDEX ActiveAlert ON Active(Alert);

                                CREATE INDEX ActiveBizJet ON Active(BizJet);

                                CREATE INDEX ActiveICAOOperator ON Active(ICAOOperator);

                                CREATE INDEX ActiveNeedsPopulating ON Active(NeedsPopulating);

                                CREATE INDEX ActiveOperator ON Active(Operator);

                                CREATE INDEX ActiveCountry ON Active(Country);

                                CREATE INDEX ActiveSquawkTrans ON Active(SquawkTrans);

                                CREATE INDEX ActivePreviousRegistration ON Active(PreviousRegistration);

                                CREATE INDEX ActivePopulated ON Active(Populated);

                                CREATE INDEX ActiveSubOperator ON Active(SubOperator);

                                CREATE INDEX ActiveICAOType ON Active(ICAOType);

                                CREATE INDEX ActiveLastModified ON Active(LastModified);

                                CREATE INDEX ActiveConstructionNumber ON Active(ConstructionNumber);

                                CREATE INDEX ActiveRoute ON Active(Route);

                                CREATE INDEX ActiveType ON Active(Type);

                                CREATE INDEX ActiveMiscode ON Active(Miscode);

                                CREATE INDEX ActiveRegistration ON Active(Registration);

                                CREATE INDEX ActiveMiscodedRegistration ON Active(MiscodedRegistration);

                                CREATE INDEX ActiveHelicopter ON Active(Helicopter);

                                CREATE INDEX ActiveLightAircraft ON Active(LightAircraft);"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "Insert into Active Select * From currentDB.Active;"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                bRet = False

            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "SELECT * from sqlite_master where name = 'XP_PROC';"
            Using BSCommand As SQLiteCommand = New SQLiteCommand(BS_SQL1, BS_Con)
                Using reader As SQLiteDataReader = BSCommand.ExecuteReader
                    bRet = reader.HasRows
                End Using
            End Using

            If bRet = True Then
                Try
                    BS_SQL1 = "CREATE TABLE XP_PROC ( view_name TEXT, param_list TEXT, xSQL TEXT,
                              def_param TEXT, opt_param TEXT, comment TEXT, PRIMARY KEY (view_name) );"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "Insert into XP_PROC Select * From currentDB.XP_PROC;"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                bRet = False

            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "SELECT * from sqlite_master where name = 'SQLITEADMIN_QUERIES';"
            Using BSCommand As SQLiteCommand = New SQLiteCommand(BS_SQL1, BS_Con)
                Using reader As SQLiteDataReader = BSCommand.ExecuteReader
                    bRet = reader.HasRows
                End Using
            End Using

            If bRet = True Then
                Try
                    BS_SQL1 = "CREATE TABLE SQLITEADMIN_QUERIES(ID INTEGER PRIMARY KEY,NAME VARCHAR(100),SQL TEXT);"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    BS_SQL1 = "Insert into SQLITEADMIN_QUERIES Select * From currentDB.SQLITEADMIN_QUERIES;"
                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                bRet = False

            End If
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        Try
            SetLabelText_ThreadSafe(Label1, vbCrLf + "Re-building BaseStation Aircraft table", Color.Yellow, 0)
            BS_SQL1 = "CREATE TABLE [Aircraft] (" &
                      "[AircraftID] integer PRIMARY KEY, " &
                      "[FirstCreated] datetime, " &
                      "[LastModified] datetime, " &
                      "[ModeS] varchar(6) Not NULL UNIQUE, " &
                      "[ModeSCountry] varchar(24), " &
                      "[Registration] varchar(20), " &
                      "[ICAOTypeCode] varchar(10), " &
                      "[SerialNo] varchar(30), " &
                      "[OperatorFlagCode] varchar(20), " &
                      "[Manufacturer] varchar(60), " &
                      "[Type] varchar(40), " &
                      "[FirstRegDate] varchar(10), " &
                      "[CurrentRegDate] varchar(10), " &
                      "[Country] varchar(24), " &
                      "[PreviousID] varchar(10), " &
                      "[DeRegDate] varchar(10), " &
                      "[Status] varchar(10), " &
                      "[PopularName] varchar(20), " &
                      "[GenericName] varchar(20), " &
                      "[AircraftClass] varchar(20), " &
                      "[Engines] varchar(40), " &
                      "[OwnershipStatus] varchar(10), " &
                      "[RegisteredOwners] varchar(100), " &
                      "[MTOW] varchar(10), " &
                      "[TotalHours] varchar(20), " &
                      "[YearBuilt] varchar(4), " &
                      "[CofACategory] varchar(30), " &
                      "[CofAExpiry] varchar(10), " &
                      "[UserNotes] varchar(300), " &
                      "[Interested] boolean Not NULL DEFAULT (0), " &
                      "[UserTag] varchar(5), " &
                      "[InfoURL] varchar(150), " &
                      "[PictureURL1] varchar(150), " &
                      "[PictureURL2] varchar(150), " &
                      "[PictureURL3] varchar(150), " &
                      "[UserBool1] boolean Not NULL DEFAULT (0), " &
                      "[UserBool2] boolean Not NULL DEFAULT (0), " &
                      "[UserBool3] boolean Not NULL DEFAULT (0), " &
                      "[UserBool4] boolean Not NULL DEFAULT (0), " &
                      "[UserBool5] boolean Not NULL DEFAULT (0), " &
                      "[UserString1] varchar(20), " &
                      "[UserString2] varchar(20), " &
                      "[UserString3] varchar(20), " &
                      "[UserString4] varchar(20), " &
                      "[UserString5] varchar(20), " &
                      "[UserInt1] integer DEFAULT (0), " &
                      "[UserInt2] integer DEFAULT (0), " &
                      "[UserInt3] integer DEFAULT (0), " &
                      "[UserInt4] integer DEFAULT (0), " &
                      "[UserInt5] integer DEFAULT (0));"

            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)



        SetLabelText_ThreadSafe(Label1, vbCrLf + "Selecting from Allhex datatable", Color.Yellow, 0)
        Dim strSql As String = "SELECT * FROM ALLHEX"
        Dim dtb As New DataTable
        Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
            con.Open()
            Using dad As New OleDbDataAdapter(strSql, con)
                dad.Fill(dtb)
            End Using
            con.Close()
        End Using
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        Dim Allhex_rowcount As Int32
        Allhex_rowcount = dtb.Rows.Count
        SetLabelText_ThreadSafe(Label4, Chr(32) & Allhex_rowcount, Color.White, 0)

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)


        SetLabelText_ThreadSafe(Label1, vbCrLf + "Load BaseStation Aircraft table in memory", Color.Yellow, 0)

        daBSCommand = New SQLiteCommand("INSERT into Aircraft (AircraftID, Modes, ModeSCountry, Registration, SerialNo, ICAOTypecode, UserInt1, UserTag, UserString1, Interested, FirstCreated, LastModified, RegisteredOwners, OperatorFlagCode)" &
                                   "VALUES (:AircraftID, :Modes, :ModeSCountry, :Registration, :SerialNo, :ICAOTypecode, :UserInt1, :UserTag, :UserString1, :Interested, :FirstCreated, :LastModified, :RegisteredOwners, :OperatorFlagCode);", BS_Con_mem)
        Dim dr1 As DataRow
        'Dim dr1count As Int32 = 0
        Try

            ' Add the parameters for the InsertCommand.
            daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")
            daBSCommand.Parameters.Add(":ModeS", DbType.String, 6, "ModeS")
            daBSCommand.Parameters.Add(":ModeSCountry", DbType.String, 24, "ModeSCountry")
            daBSCommand.Parameters.Add(":Registration", DbType.String, 20, "Registration")
            daBSCommand.Parameters.Add(":SerialNo", DbType.String, 30, "Serial")
            daBSCommand.Parameters.Add(":ICAOTypeCode", DbType.String, 10, "ICAOTypecode")
            daBSCommand.Parameters.Add(":UserInt1", DbType.Int32, 0, "UserInt1")
            daBSCommand.Parameters.Add(":UserTag", DbType.String, 5, "UserTag")
            daBSCommand.Parameters.Add(":UserString1", DbType.String, 20, "UserString1")
            daBSCommand.Parameters.Add(":Interested", DbType.Boolean, 0, "Interested")
            daBSCommand.Parameters.Add(":FirstCreated", DbType.DateTime, 0, "FirstCreated")
            daBSCommand.Parameters.Add(":LastModified", DbType.DateTime, 0, "LastModified")
            daBSCommand.Parameters.Add(":RegisteredOwners", DbType.String, 100, "RegisteredOwners")
            daBSCommand.Parameters.Add(":OperatorFlagCode", DbType.String, 20, "OperatorFlagCode")

            Using t As SQLiteTransaction = BS_Con_mem.BeginTransaction()

                For Each dr1 In dtb.Rows

                    daBSCommand.Parameters(":AircraftID").Value = (dr1("AircraftID"))
                    daBSCommand.Parameters(":ModeS").Value = (dr1("ModeS"))
                    daBSCommand.Parameters(":ModeSCountry").Value = (dr1("ModeSCountry"))
                    daBSCommand.Parameters(":Registration").Value = (dr1("Registration"))
                    daBSCommand.Parameters(":SerialNo").Value = (dr1("SerialNo"))
                    daBSCommand.Parameters(":ICAOTypeCode").Value = (dr1("ICAOTypeCode"))
                    daBSCommand.Parameters(":UserTag").Value = (dr1("UserTag"))
                    daBSCommand.Parameters(":UserInt1").Value = (dr1("UserInt1"))
                    daBSCommand.Parameters(":UserString1").Value = (dr1("UserString1"))
                    daBSCommand.Parameters(":Interested").Value = (dr1("Interested"))
                    daBSCommand.Parameters(":FirstCreated").Value = Now
                    daBSCommand.Parameters(":LastModified").Value = Now
                    daBSCommand.Parameters(":RegisteredOwners").Value = (dr1("RegisteredOwners"))
                    daBSCommand.Parameters(":OperatorFlagCode").Value = (dr1("OperatorFlagCode"))

                    'dr1count = dr1count + 1

                    da1.InsertCommand = daBSCommand
                    da1.InsertCommand.ExecuteNonQuery()

                Next

                t.Commit()

                'Dim dr1count_string As String = dr1count.ToString
                'SetLabelText_ThreadSafe(Me.Label5, dr1count_string, Color.White, 1)

            End Using
            dtb.Dispose()
        Catch

        End Try

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Build Aircraft Indices", Color.Yellow, 0)


        Try

            BS_SQL1 = "CREATE INDEX [AircraftAircraftClass] ON [Aircraft] ([AircraftClass]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = ("CREATE INDEX [AircraftCountry] ON [Aircraft] ([Country]);")
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftGenericName] ON [Aircraft] ([GenericName]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftICAOTypeCode] ON [Aircraft] ([ICAOTypeCode]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try

            BS_SQL1 = "CREATE INDEX [AircraftInterested] ON [Aircraft] ([Interested]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftManufacturer] ON [Aircraft] ([Manufacturer]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftModeS] ON [Aircraft] ([ModeS]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftModeSCountry] ON [Aircraft] ([ModeSCountry]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftPopularName] ON [Aircraft] ([PopularName]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftRegisteredOwners] ON [Aircraft] ([RegisteredOwners]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftRegistration] ON [Aircraft] ([Registration]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftSerialNo] ON [Aircraft] ([SerialNo]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftType] ON [Aircraft] ([Type]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftUserTag] ON [Aircraft] ([UserTag]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE INDEX [AircraftYearBuilt] ON [Aircraft] ([YearBuilt]);"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        Try
            BS_SQL1 = "CREATE TRIGGER [AircraftIDdeltrig] BEFORE DELETE ON [Aircraft]" &
                    " FOR EACH ROW BEGIN DELETE FROM Flights WHERE AircraftID = OLD.AircraftID;END;"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)


        Dim BSReccount As Integer
        daBSCommand = New SQLiteCommand("SELECT COUNT(*) from Aircraft", BS_Con_mem)
        BSReccount = Convert.ToInt32(daBSCommand.ExecuteScalar)
        SetLabelText_ThreadSafe(Label5, BSReccount, Color.White, 1)

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton9.Checked = True Then
            Try

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Load Personal Operator Flag codes", Color.Yellow, 0)
                BS_SQL1 = "CREATE TABLE PersonalFlags ([ModeS] varchar(6) PRIMARY KEY, [PersonalFlagCode] varchar(20) );

                        INSERT INTO PersonalFlags SELECT ModeS, OperatorFlagCode AS 'PersonalFlagCode' from [currentDB].Aircraft;

                        UPDATE Aircraft SET OperatorFlagCode = (SELECT PersonalFlagCode from PersonalFlags
                        WHERE Aircraft.ModeS = PersonalFlags.ModeS);

                        DROP TABLE Personalflags;"

                daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                daBSCommand.ExecuteNonQuery()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
            End Try
            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

            If CheckBox3.Checked = True Then
                Try

                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Update Personal Flags to Set No Null Operator Flags", Color.Yellow, 0)
                    BS_SQL1 = "UPDATE Aircraft Set OperatorFlagCode = ICAOTypeCode WHERE (OperatorFlagCode = '-' or OperatorFlagCode ISNULL);"

                    daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
                    daBSCommand.ExecuteNonQuery()
                Catch ex As Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try
            End If

            SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
        End If

        Try

            SetLabelText_ThreadSafe(Label1, vbCrLf + "Detach Current BaseStation database", Color.Yellow, 0)
            BS_SQL1 = "DETACH DATABASE currentDB"
            daBSCommand = New SQLiteCommand(BS_SQL1, BS_Con_mem)
            daBSCommand.ExecuteNonQuery()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Save BaseStation Aircraft table in memory to disk", Color.Yellow, 0)

        Dim BS_new As New SQLiteConnection("Provider=System.Data.SQLite;DataSource='" + PACBSLoc + "\basestation_PAC.sqb';Version=3;New=True;")
        Using BS_Con
            'BS_Con_mem.Open()
            'BS_Con.Open()
            BS_new.Open()
            BS_Con_mem.BackupDatabase(BS_new, "main", "main", -1, Nothing, 0)
            BS_new.Dispose()

        End Using

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

        GoTo CompactStep

QuickUpdate:
        BS_Con.Open()

        'Removing Ps if previously logged with this registration (i.e. repeated lease) Part 1
        'SetLabelText_ThreadSafe(Me.Label1, vbCrLf + "Extracting to Ps_reset", Color.Yellow, 0)
        AccessSQL("INSERT INTO Ps_Reset ( AircraftID, Registration, Hex )" &
              " SELECT DISTINCT logLLp.ID AS AircraftID, logLLp.Registration, tbldataset.Hex" &
              " FROM (logLLp INNER JOIN tblOperatorHistory ON (logLLp.ID = tblOperatorHistory.ID) AND (logLLp.Registration = tblOperatorHistory.previous))" &
              " INNER JOIN tbldataset ON logLLp.ID = tbldataset.ID" &
              " WHERE (((logLLp.Registration)<>([tbldataset].[registration]) And (logLLp.Registration) In (select tblOperatorHistory.previous from tblOperatorHistory)));", DBName)
        'SetLabelText_ThreadSafe(Me.Label2, vbCrLf + "Completed", Color.Green, 0)

        If RadioButton4.Checked = True Then 'Loggings only

            If RadioButton2.Checked = True Then 'Set Interested for all required

                Try
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing Interested fields", Color.Yellow, 0)
                    BS_SQL1 = "Update Aircraft SET Interested = 0 WHERE Interested = 1;"
                    Dim cmd2 As New SQLiteCommand(BS_SQL1, BS_Con)
                    cmd2.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field", Color.Yellow, 0)
                strSql = "SELECT logllp.ID AS AircraftID, logllp.Registration, tbldataset.Hex" &
                      " FROM (logllp INNER JOIN (SELECT ID, Max([when]) as LastDate" &
                      " FROM logllp GROUP BY ID) AS B ON (logllp.ID = B.ID) AND" &
                      " (logllp.[when] = B.LastDate)) INNER JOIN tbldataset ON B.ID = tbldataset.ID" &
                      " where tbldataset.hex is not null;"
                dtb = New DataTable
                Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                    con.Open()
                    Using dad As New OleDbDataAdapter(strSql, con)
                        dad.Fill(dtb)
                    End Using
                    con.Close()
                End Using

                SQLTrans = BS_Con.BeginTransaction

                daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET Interested = 1, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                        Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID;", BS_Con)

                Try

                    ' Add the parameters for the InsertCommand.
                    daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")

                    For Each dr2 In dtb.Rows
                        daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                        da1.UpdateCommand = daBSCommand
                        da1.UpdateCommand.ExecuteNonQuery()
                    Next

                    SQLTrans.Commit()
                    dtb.Dispose()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field part 2", Color.Yellow, 0)
                strSql = "SELECT Ps_Reset.AircraftID, Ps_Reset.Registration" &
                            " FROM Ps_Reset INNER JOIN tbldataset ON Ps_Reset.AircraftID = tbldataset.ID" &
                            " Where Ps_reset.Registration <> tbldataset.registration;"
                dtc = New DataTable
                Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                    con.Open()
                    Using dad As New OleDbDataAdapter(strSql, con)
                        dad.Fill(dtc)
                    End Using
                    con.Close()
                End Using

                SQLTrans = BS_Con.BeginTransaction

                daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET Interested = 0, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                        Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID AND Registration <> :Registration;", BS_Con)
                Try

                    ' Add the parameters for the InsertCommand.
                    daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")
                    daBSCommand.Parameters.Add(":Registration", DbType.String, 20, "Registration")

                    For Each dr2 In dtb.Rows
                        daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                        daBSCommand.Parameters(":Registration").Value = (dr2("Registration"))
                        da1.UpdateCommand = daBSCommand
                        da1.UpdateCommand.ExecuteNonQuery()
                    Next

                    SQLTrans.Commit()
                    dtb.Dispose()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                If CheckBox1.Checked = False Then
                    Try
                        SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing PlanePotter Symbols", Color.Yellow, 0)
                        BS_SQL1 = "Update Aircraft SET UserTag = NULL;"
                        Dim cmd2 As New SQLiteCommand(BS_SQL1, BS_Con)
                        cmd2.ExecuteNonQuery()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                End If

                If CheckBox2.Checked = True Then

                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting Interested field for Ignore Ps", Color.Yellow, 0)
                    strSql = "Select AircraftID from Ps_Reset;"
                    dtb = New DataTable
                    Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                        con.Open()
                        Using dad As New OleDbDataAdapter(strSql, con)
                            dad.Fill(dtb)
                        End Using
                        con.Close()
                    End Using

                    SQLTrans = BS_Con.BeginTransaction

                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET Interested = 1, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                            Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID;", BS_Con)

                    Try

                        ' Add the parameters for the InsertCommand.
                        daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")

                        For Each dr2 In dtb.Rows
                            daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                            da1.UpdateCommand = daBSCommand
                            da1.UpdateCommand.ExecuteNonQuery()
                        Next

                        SQLTrans.Commit()
                        dtb.Dispose()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                End If

                If CheckBox1.Checked = True Then

                    'Select Case RadioButton7.Checked
                    '    Case True
                    '        SetLabelText_ThreadSafe(Label1, vbCrLf + "Loading PlanePlotter v1 Symbols", Color.Yellow, 0)

                    '        AccessSQL("UPDATE Allhex INNER JOIN PP_SymbolsByType On Allhex.ICAOTypeCode = PP_SymbolsByType.ICAOTypeCode Set" &
                    '              " allhex.UserString1 = PP_SymbolsByType.UserString1_v1;", DBName)
                    '        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
                    '    Case Else
                    '        SetLabelText_ThreadSafe(Label1, vbCrLf + "Loading PlanePlotter v3 Symbols", Color.Yellow, 0)
                    '        AccessSQL("UPDATE Allhex INNER JOIN PP_SymbolsByType On Allhex.ICAOTypeCode = PP_SymbolsByType.ICAOTypeCode Set" &
                    '              " allhex.UserString1 = PP_SymbolsByType.UserString1_v3;", DBName)
                    '        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
                    'End Select

                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting PlanePlotter Symbols", Color.Yellow, 0)

                    SQLTrans = BS_Con.BeginTransaction

                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET UserTag = UserString1, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                            Chr(39) & "localtime" & Chr(39) & ");", BS_Con)

                    Try
                        da1.UpdateCommand = daBSCommand
                        da1.UpdateCommand.ExecuteNonQuery()

                        SQLTrans.Commit()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)



                    'SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting PlanePlotter Symbols", Color.Yellow, 0)
                    'AccessSQL("UPDATE Allhex Set Allhex.UserTag = Allhex.UserTag & AllHex.UserString1;", DBName)
                    'SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                End If
                BS_Con.Dispose()
                BS_Con_mem.Dispose()

                GoTo ENDSUB

            ElseIf RadioButton1.Checked = True Or RadioButton10.Checked = True Then 'Set RQ/Ps for all and Interested for Mil only

                Try
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing User Tag fields", Color.Yellow, 0)
                    BS_SQL1 = "Update Aircraft SET UserTag = " & """RQ""" & ";"
                    Dim cmd2 As New SQLiteCommand(BS_SQL1, BS_Con)
                    cmd2.ExecuteNonQuery()
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                Try
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Updating UserTag field only", Color.Yellow, 0)
                    strSql = "SELECT logllp.ID AS AircraftID, logllp.Registration, tbldataset.Hex" &
                      " FROM (logllp INNER JOIN (SELECT ID, Max([when]) as LastDate" &
                      " FROM logllp GROUP BY ID) AS B ON (logllp.ID = B.ID) AND" &
                      " (logllp.[when] = B.LastDate)) INNER JOIN tbldataset ON B.ID = tbldataset.ID" &
                      " where tbldataset.hex is not null AND logllp.Registration = tbldataset.registration;"
                    dtb = New DataTable
                    Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                        con.Open()
                        Using dad As New OleDbDataAdapter(strSql, con)
                            dad.Fill(dtb)
                        End Using
                        con.Close()
                    End Using

                    SQLTrans = BS_Con.BeginTransaction

                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET UserTag = null, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                        Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID AND Registration = :Registration;", BS_Con)
                    Dim dr2 As DataRow
                    Try

                        ' Add the parameters for the InsertCommand.
                        daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")
                        daBSCommand.Parameters.Add(":Registration", DbType.String, 20, "Registration")

                        For Each dr2 In dtb.Rows
                            daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                            daBSCommand.Parameters(":Registration").Value = (dr2("Registration"))
                            da1.UpdateCommand = daBSCommand
                            da1.UpdateCommand.ExecuteNonQuery()
                        Next

                        SQLTrans.Commit()
                        dtb.Dispose()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                Try
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting UserTag to Ps where appropriate", Color.Yellow, 0)
                    strSql = "SELECT logllp.ID As AircraftID, logllp.Registration, tbldataset.Hex" &
                      " FROM (logllp INNER JOIN (SELECT ID, Max([when]) as LastDate" &
                      " FROM logllp GROUP BY ID) AS B ON (logllp.ID = B.ID) AND" &
                      " (logllp.[when] = B.LastDate)) INNER JOIN tbldataset ON B.ID = tbldataset.ID" &
                      " where tbldataset.hex is not null AND logllp.registration <> tbldataset.registration;"
                    dtb = New DataTable
                    Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                        con.Open()
                        Using dad As New OleDbDataAdapter(strSql, con)
                            dad.Fill(dtb)
                        End Using
                        con.Close()
                    End Using

                    SQLTrans = BS_Con.BeginTransaction

                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET UserTag = " & """Ps""" & ", LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                        Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID;", BS_Con)
                    Dim dr2 As DataRow
                    Try

                        ' Add the parameters for the InsertCommand.
                        daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")

                        For Each dr2 In dtb.Rows
                            daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                            da1.UpdateCommand = daBSCommand
                            da1.UpdateCommand.ExecuteNonQuery()
                        Next

                        SQLTrans.Commit()
                        dtb.Dispose()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                Try
                    strSql = "SELECT Ps_Reset.AircraftID, Ps_Reset.Registration" &
                            " FROM Ps_Reset INNER JOIN tbldataset ON Ps_Reset.AircraftID = tbldataset.ID;"
                    dtb = New DataTable
                    Using con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBName & "")
                        con.Open()
                        Using dad As New OleDbDataAdapter(strSql, con)
                            dad.Fill(dtb)
                        End Using
                        con.Close()
                    End Using

                    SQLTrans = BS_Con.BeginTransaction

                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET UserTag = NULL, LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                        Chr(39) & "localtime" & Chr(39) & ") where AircraftID = :AircraftID AND Registration = :Registration;", BS_Con)
                    Dim dr2 As DataRow
                    Try

                        ' Add the parameters for the InsertCommand.
                        daBSCommand.Parameters.Add(":AircraftID", DbType.String, 6, "AircraftID")
                        daBSCommand.Parameters.Add(":Registration", DbType.String, 20, "Registration")

                        For Each dr2 In dtb.Rows
                            daBSCommand.Parameters(":AircraftID").Value = (dr2("AircraftID"))
                            daBSCommand.Parameters(":Registration").Value = (dr2("Registration"))
                            da1.UpdateCommand = daBSCommand
                            da1.UpdateCommand.ExecuteNonQuery()
                        Next

                        SQLTrans.Commit()
                        dtb.Dispose()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try
                Catch ex As System.Exception
                    System.Windows.Forms.MessageBox.Show(ex.Message)
                End Try

                SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

                If RadioButton10.Checked = True Then
                    Try
                        SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting all RQ/Ps to Interested", Color.Yellow, 0)
                        BS_SQL1 = "Update Aircraft SET Interested = TRUE where UserTag Contains " & """RQ*"" " & " OR UserTag Contains " & """Ps*""" & ";"
                        Dim cmd2 As New SQLiteCommand(BS_SQL1, BS_Con)
                        cmd2.ExecuteNonQuery()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                End If


                If CheckBox1.Checked = True Then

                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Setting PlanePlotter Symbols", Color.Yellow, 0)

                    Try

                        SQLTrans = BS_Con.BeginTransaction
                        Try
                            daBSCommand = New SQLiteCommand("Update Aircraft Set UserTag = COALESCE((UserTag || UserString1), UserString1), LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                            Chr(39) & "localtime" & Chr(39) & ");", BS_Con)

                            da1.UpdateCommand = daBSCommand
                            da1.UpdateCommand.ExecuteNonQuery()

                            SQLTrans.Commit()
                        Catch ex As System.Exception
                            System.Windows.Forms.MessageBox.Show(ex.Message)
                        End Try
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
                Else
                    SetLabelText_ThreadSafe(Label1, vbCrLf + "Clearing PlanePlotter Symbols", Color.Yellow, 0)
                    SQLTrans = BS_Con.BeginTransaction
                    daBSCommand = New SQLiteCommand("UPDATE AIRCRAFT SET UserTag = ifnull(UserTag," & Chr(39) & Chr(39) & "), LastModified = DATETIME(" & Chr(39) & "now" & Chr(39) & "," &
                                            Chr(39) & "localtime" & Chr(39) & ");", BS_Con)
                    Try
                        da1.UpdateCommand = daBSCommand
                        da1.UpdateCommand.ExecuteNonQuery()

                        SQLTrans.Commit()
                    Catch ex As System.Exception
                        System.Windows.Forms.MessageBox.Show(ex.Message)
                    End Try

                    SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)
                End If

            End If
        End If

        BS_Con.Dispose()
        BS_Con_mem.Dispose()

        GoTo ENDSUB


CompactStep:
        Con.Close()
        BS_Con.Dispose()
        BS_Con_mem.Dispose()

        SetLabelText_ThreadSafe(Label1, vbCrLf + "Finalise basestation.sqb and compact Logged.mdb", Color.Yellow, 0)

        FileToDelete = BSLoc
        FileToRename = PACBSLoc + "\basestation_PAC.sqb"

        If System.IO.File.Exists(FileToDelete) = True Then

            System.IO.File.Delete(FileToDelete)

        End If
        My.Computer.FileSystem.RenameFile(FileToRename, "basestation.sqb")

        CompactDB()

        SetLabelText_ThreadSafe(Label2, vbCrLf + "Completed", Color.Green, 0)

ENDSUB:

        Con.Close()
        BS_Con.Dispose()
        BS_Con_mem.Dispose()

        stop_time = Now
        SetLabelText_ThreadSafe(Label6, vbCrLf + "End Time" & Chr(32) & stop_time, Color.Blue, 0)
        elapsed_time = stop_time.Subtract(start_time)
        elapsed = elapsed_time.TotalSeconds.ToString("0.00")
        SetLabelText_ThreadSafe(Label6, vbCrLf + "Elapsed" & Chr(32) & elapsed, Color.Blue, 0)

        Exit Sub

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        ' Is the Background Worker do some work?
        If BGW.IsBusy Then
            'If it supports cancellation, Cancel It
            If BGW.WorkerSupportsCancellation Then
                ' Tell the Background Worker to stop working.
                BGW.CancelAsync()
            End If
        End If
        ' Enable to Start Button
        Button3.Enabled = True
        ' Disable to Stop Button
        Button4.Enabled = False
        Close()

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub

    Private Sub BGW_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        AddHandler BGW.ProgressChanged, AddressOf BGW_ProgressChanged

        ProgressBar1.Refresh()
        ProgressBar1.Style = ProgressBarStyle.Marquee

        For I As Integer = 0 To 99
            ProgressBar1.Refresh()
        Next

        'Me.Label1.Text = e.ProgressPercentage.ToString() & "%"

    End Sub

    Private Sub BGW_RunWorkerCompleted(ByVal sender As Object, ByVal e As AsyncCompletedEventArgs) Handles BGW.RunWorkerCompleted

        AddHandler BGW.RunWorkerCompleted, AddressOf BGW_RunWorkerCompleted
        ProgressBar1.Hide()
        If e.Cancelled = True Then
            Label1.Text = "Canceled!"
        ElseIf e.Error IsNot Nothing Then
            Label1.Text = "Error: " & e.Error.Message
        Else
            Label2.Text += vbCrLf & "All Done!"

        End If
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Show()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Close()
    End Sub

    ' The delegate
    Public Delegate Sub SetLabelText_Delegate(ByVal [Label] As Label, ByVal [text] As String, ByVal [colour] As Color, ByVal [Select] As Integer)

    ' The delegates subroutine.
    Private Sub SetLabelText_ThreadSafe(ByVal [Label] As Label, ByVal [text] As String, ByVal [colour] As Color, ByVal [Select] As Integer)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If [Label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Invoke(MyDelegate, New Object() {[Label], [text], [colour], [Select]})
        ElseIf [Select] = 0 Then
            [Label].Text += [text]
            [Label].BackColor = [colour]
        ElseIf [Select] = 1 Then
            [Label].Text = "Records Loaded " & [text]
            [Label].BackColor = [colour]
        End If
        ''Write logfile
        'Using writer As StreamWriter = New StreamWriter(logfile, True)
        '    writer.WriteLine([text] & "-" & Date.Now)
        'End Using
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        My.Settings.BSloc = TextBox1.Text
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        My.Settings.BSBackupLoc = TextBox2.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Using dialog As New FolderBrowserDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            BSLoc = dialog.SelectedPath
            TextBox1.Text = dialog.SelectedPath
            My.Settings.BSloc = dialog.SelectedPath
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Using dialog As New FolderBrowserDialog
            If dialog.ShowDialog() <> DialogResult.OK Then Return
            BSBackupLoc = dialog.SelectedPath
            TextBox2.Text = dialog.SelectedPath
            My.Settings.BSBackupLoc = dialog.SelectedPath
        End Using
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        RadioButton1.Enabled = True
        My.Settings.RQPsButton = True
        My.Settings.InterestedButton = False
        My.Settings.RQPsandIButton = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        RadioButton2.Enabled = True
        My.Settings.InterestedButton = True
        My.Settings.RQPsButton = False
        My.Settings.RQPsandIButton = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            My.Settings.PPSymbols = True
            RadioButton7.Visible = True
            RadioButton8.Visible = True
        Else
            RadioButton7.Visible = False
            RadioButton8.Visible = False
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Enabled = True Then
            CheckBox3.Enabled = True
            My.Settings.OperatorFlags = "Kinetic"
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Enabled = True Then
            CheckBox3.Enabled = False
            My.Settings.OperatorFlags = "GFIA"
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If RadioButton9.Enabled = True Then
            CheckBox3.Enabled = True
            My.Settings.OperatorFlags = "Personal"
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Enabled = True Then
            My.Settings.PPSymbolsType = "v1"
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If RadioButton8.Enabled = True Then
            My.Settings.PPSymbolsType = "v3"
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            RadioButton5.Visible = False
            RadioButton6.Visible = False
        Else
            RadioButton5.Visible = True
            RadioButton6.Visible = True
        End If
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged
        RadioButton10.Enabled = True
        My.Settings.RQPsandIButton = True
        My.Settings.InterestedButton = False
        My.Settings.RQPsButton = False
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            My.Settings.NullOpFlags = True
        Else
            My.Settings.NullOpFlags = False
        End If
    End Sub

End Class
