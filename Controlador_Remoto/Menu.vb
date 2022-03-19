'<Fecha> 27/11/2021
'<Ver> 1.6'
Imports System.Net.Sockets
Imports System.Net
Imports System.IO
Imports System.Runtime.InteropServices
Public Class Menu

    Private port As Integer = 6969
    Private direction As IPAddress
    Private StatusRemote As Boolean = False
    Private tryRemoteConnection As Boolean = False
    Private remoteListener As TcpListener
    Dim allDrives As DriveInfo() = System.IO.DriveInfo.GetDrives
    Private readQueue As Boolean = False
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    Const appComands As UInteger = &H319
    Const Volume_UP As UInteger = &HA
    Const Volume_DOWN As UInteger = &H9
    Const Mute As UInteger = &H8
    Private StreamR As StreamReader
    Private StreamW As StreamWriter
    Private client As New TcpClient()
    Private StreamN As NetworkStream
    Private continueSwitch As Boolean
    Private stringReadBox As String



    Private Sub b_Remote_Click(sender As Object, e As EventArgs) Handles b_Remote.Click
        If (b_Remote.Text = "Remote") Then
            b_Controller.Visible = False
            b_Remote.Text = "Disconnect"
            lbl_My_Ip.Visible = True
            direction = Dns.GetHostEntry(My.Computer.Name).AddressList.FirstOrDefault(Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork)
            lbl_My_Ip.Text = direction.ToString
            start_Remote()
        ElseIf (b_Remote.Text = "Disconnect") Then
            disconnect_Client()

        End If


    End Sub

    Private Function disconnect_Client()
        If (StatusRemote = False) Then
            send_Controller("/Disconnect")
        End If
        StopServer()
        lbl_My_Ip.Text = ""
        b_Remote.Text = "Remote"
        b_Controller.Visible = True
    End Function


    Private Function start_Remote()
        If StatusRemote = False Then
            tryRemoteConnection = True
            Try
                remoteListener = New TcpListener(IPAddress.Any, port)
                remoteListener.Start()
                StatusRemote = True
                Threading.ThreadPool.QueueUserWorkItem(AddressOf cycle_Remote)
            Catch ex As Exception
                StatusRemote = False
            End Try
            tryRemoteConnection = False

        End If

        Return True
    End Function



    Private Function StopServer()
        If StatusRemote = True Then
            tryRemoteConnection = True
            Try
                client.Close()
                remoteListener.Stop()
                StatusRemote = False
            Catch ex As Exception
                StopServer()
            End Try
            tryRemoteConnection = False
        End If
        Return True
    End Function


    Private Sub cycle_Remote()
        Try
            Using client As TcpClient = remoteListener.AcceptTcpClient
                If tryRemoteConnection = False Then
                    Threading.ThreadPool.QueueUserWorkItem(AddressOf cycle_Remote)
                End If
                StreamR = New StreamReader(client.GetStream)
                StreamW = New StreamWriter(client.GetStream)
                Try
                    If StreamR.BaseStream.CanRead = True Then
                        While (StreamR.BaseStream.CanRead = True)
                            stringReadBox = StreamR.ReadLine

                            If (stringReadBox.Substring(0, 1) = "/") Then
                                internComands(stringReadBox)
                            Else
                                txt_Answers.Text = stringReadBox
                                stringReadBox = ""
                            End If


                        End While

                    End If
                Catch ex As Exception
                    client.Close()
                End Try


            End Using
        Catch ex As Exception

        End Try

    End Sub


    Private Sub internComands(comand As String)

        Select Case comand
            Case = "/Disconnect"
                StatusRemote = False
                disconnect_Client()
            Case = "/DesktopName"
                send_Controller(My.Computer.Name.ToString)

            Case = "/username"
                Dim partes() As String = Split(My.User.Name, "\")
                send_Controller(partes(partes.Length - 1))

            Case = "/OperativeSystem"
                send_Controller(My.Computer.Info.OSFullName)


            Case = "/OperativeSystemVersion"
                send_Controller(My.Computer.Info.OSVersion)


            Case = "/RAM"
                send_Controller(Math.Truncate((((My.Computer.Info.TotalPhysicalMemory / 1024) / 1024) / 1024) * 10000) / 10000 & " GB")

            Case = "/freeRAM"
                send_Controller(Math.Truncate((((My.Computer.Info.AvailablePhysicalMemory / 1024) / 1024) / 1024) * 10000) / 10000 & " GB")

            Case = "/ScreenResolution"
                send_Controller(Screen.PrimaryScreen.Bounds.Width & " x " & Screen.PrimaryScreen.Bounds.Height)

            Case = "/Disks"
                infoDiscos()

            Case = "/timeZone"
                send_Controller(TimeZoneInfo.Local.ToString)


            Case = "/date"
                send_Controller(DateTime.Now.ToString)

            Case = "/platform"
                send_Controller(My.Computer.Info.OSPlatform)

            Case = "/VolumenDown"
                SendMessage(Me.Handle, appComands, &H30292, Volume_DOWN * &H10000)

            Case = "/VolumenUP"
                SendMessage(Me.Handle, appComands, &H30292, Volume_UP * &H10000)

            Case = "/Mute"
                SendMessage(Me.Handle, appComands, &H200EB0, Mute * &H10000)


            Case = "/ShutDown"
                System.Diagnostics.Process.Start("shutdown", "-s -t 00")

            Case = "/Restart"
                System.Diagnostics.Process.Start("shutdown", "-r -t 00")

            Case = "/Logout"
                System.Diagnostics.Process.Start("shutdown.exe", "/l")

            Case = "/infoCPU"
                send_Controller(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "ProcessorNameString", Nothing))

            Case = "/processList"
                processList()
                Exit Select

        End Select

    End Sub


    Private Function send_Controller(ByVal dato As String)

        If StatusRemote = True Then
            Try

                StreamW.WriteLine(dato)
                StreamW.Flush()

            Catch ex As Exception
                send_Controller(dato)

            End Try
        End If

        Return True

    End Function


    Private Sub lbl_My_Ip_Click(sender As Object, e As EventArgs) Handles lbl_My_Ip.Click
        Clipboard.SetText(lbl_My_Ip.Text)
    End Sub

    Private Sub infoDiscos()

        send_Controller("______________List HDD____________")
        For Each DR As DriveInfo In allDrives

            send_Controller("- Letter :" + DR.Name)


            If (DR.IsReady = True) Then

                send_Controller(" Archive System: " + DR.DriveFormat.ToString)

                Dim disponible As Integer = Math.Truncate((((DR.AvailableFreeSpace / 1024) / 1024) / 1024) * 10000) / 10000
                Dim Total As Integer = Math.Truncate((((DR.TotalSize / 1024) / 1024) / 1024) * 10000) / 10000

                send_Controller(" Memory Free : " + disponible.ToString + " GB ")

                send_Controller(" Total Memory: " + Total.ToString + " GB ")

                send_Controller(" Used Memory: " + Math.Truncate(Total - disponible).ToString + " GB")


            End If


        Next DR
    End Sub



    Private Sub processList()
        Dim psLista() As Process
        Try
            psLista = Process.GetProcesses()
            send_Controller("ID   -      Name Process")


            For Each p As Process In psLista

                send_Controller(p.Id.ToString() + "  -     " + p.ProcessName)


            Next p

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub



    Private Sub b_Controlador_Click(sender As Object, e As EventArgs) Handles b_Controller.Click

        If (b_Controller.Text = "Controller") Then


            b_Remote.Visible = False
            Label_IP.Visible = True
            TextBox_IP.Visible = Not (TextBox_IP.Visible)
            b_Connect.Visible = Not (b_Connect.Visible)
            b_Controller.Text = "Disconnect"
        ElseIf (b_Controller.Text = "Disconnect") Then
            server_Disconnect()
        End If


    End Sub

    Private Function server_Disconnect()


        If (b_ShutDown.Visible = True) Then
            senderRemote("/Disconnect")
            menuOptions()
        End If
        txt_Answers.Clear()
        b_Connect.Visible = Not (b_Connect.Visible)
        TextBox_IP.Visible = Not (TextBox_IP.Visible)
        b_Controller.Text = "Controller"
        txt_Answers.Text = "Disconnected"

        client.Close()
        Label_IP.Visible = False
        b_Controller.Visible = True
        b_Remote.Visible = True
    End Function



    Private Sub Connect_Click(sender As Object, e As EventArgs) Handles b_Connect.Click
        If TextBox_IP.Text = "" Then
            txt_Answers.Text = "IP not found"
        Else

            Try
                direction = IPAddress.Parse(TextBox_IP.Text)
                client = New TcpClient(direction.ToString, port)

                If client.GetStream.CanRead = True Then
                    StreamR = New StreamReader(client.GetStream)
                    StreamW = New StreamWriter(client.GetStream)


                    Threading.ThreadPool.QueueUserWorkItem(AddressOf stateConnected)
                    menuOptions()
                End If
                txt_Answers.Text = "Connected!"

            Catch ex As Exception
                txt_Answers.Text = "IP failed " + vbCrLf + ex.Message
            End Try


        End If

    End Sub

    Private Function menuOptions()
        Array.ForEach(Me.Controls.OfType(Of Label).ToArray, Sub(lbl) lbl.Visible = Not lbl.Visible)
        b_DesktopName.Visible = Not (b_DesktopName.Visible)
        b_username.Visible = Not (b_username.Visible)
        b_OperativeSystemName.Visible = Not (b_OperativeSystemName.Visible)
        b_OperativeSystemVersion.Visible = Not (b_OperativeSystemVersion.Visible)
        b_Total_RAM.Visible = Not (b_Total_RAM.Visible)
        b_Free_RAM.Visible = Not (b_Free_RAM.Visible)
        b_ScreenResolution.Visible = Not (b_ScreenResolution.Visible)
        b_Disks.Visible = Not (b_Disks.Visible)
        b_timeZone.Visible = Not (b_timeZone.Visible)
        b_date.Visible = Not (b_date.Visible)
        b_platform.Visible = Not (b_platform.Visible)
        b_InfoCPU.Visible = Not (b_InfoCPU.Visible)
        b_processList.Visible = Not (b_processList.Visible)
        b_Volumen_UP.Visible = Not (b_Volumen_UP.Visible)
        b_Volumen_Down.Visible = Not (b_Volumen_Down.Visible)
        b_Mute.Visible = Not (b_Mute.Visible)
        b_LogOut.Visible = Not (b_LogOut.Visible)
        b_Restart.Visible = Not (b_Restart.Visible)
        b_ShutDown.Visible = Not (b_ShutDown.Visible)
        txt_console.Visible = Not (txt_console.Visible)
        btn_Enviar.Visible = Not (btn_Enviar.Visible)
        btn_verCaptura.Visible = Not (btn_verCaptura.Visible)
    End Function


    Private Sub stateConnected()
        If StreamR.BaseStream.CanRead = True Then

            Try
                While StreamR.BaseStream.CanRead = True

                    stringReadBox = StreamR.ReadLine + vbCrLf
                    txt_Answers.Text += stringReadBox

                    If (stringReadBox.Substring(0, 1) = "/") Then
                        server_Disconnect()
                    End If

                End While
            Catch ex As Exception


            End Try
        End If
    End Sub


    Private Sub senderRemote(ByVal dato As String)
        Try
            StreamW.WriteLine(dato)
            StreamW.Flush()
        Catch ex As Exception

        End Try

    End Sub


    Private Sub btn_send_Click(sender As Object, e As EventArgs) Handles btn_Enviar.Click
        stringReadBox = txt_console.Text
        If (stringReadBox <> "") Then
            senderRemote(stringReadBox)
            txt_console.Clear()
        End If
        stringReadBox = ""
    End Sub


    Private Sub b_DesktopName_Click(sender As Object, e As EventArgs) Handles b_DesktopName.Click
        txt_Answers.Clear()
        senderRemote("/DesktopName")
    End Sub


    Private Sub b_username_Click(sender As Object, e As EventArgs) Handles b_username.Click
        txt_Answers.Clear()
        senderRemote("/username")
    End Sub
    Private Sub b_OperativeSystemName_Click(sender As Object, e As EventArgs) Handles b_OperativeSystemName.Click
        txt_Answers.Clear()
        senderRemote("/OperativeSystem")
    End Sub
    Private Sub b_OperativeSystemVersion_Click(sender As Object, e As EventArgs) Handles b_OperativeSystemVersion.Click
        txt_Answers.Clear()
        senderRemote("/OperativeSystemVersion")

    End Sub

    Private Sub b_Total_RAM_Click(sender As Object, e As EventArgs) Handles b_Total_RAM.Click
        txt_Answers.Clear()
        senderRemote("/RAM")

    End Sub

    Private Sub b_Free_RAM_Click(sender As Object, e As EventArgs) Handles b_Free_RAM.Click
        txt_Answers.Clear()
        senderRemote("/freeRAM")

    End Sub

    Private Sub b_ScreenResolution_Click(sender As Object, e As EventArgs) Handles b_ScreenResolution.Click
        txt_Answers.Clear()
        senderRemote("/ScreenResolution")

    End Sub

    Private Sub b_Disks_Click(sender As Object, e As EventArgs) Handles b_Disks.Click
        txt_Answers.Clear()
        senderRemote("/Disks")
    End Sub

    Private Sub b_timeZone_Click(sender As Object, e As EventArgs) Handles b_timeZone.Click
        txt_Answers.Clear()
        senderRemote("/timeZone")

    End Sub

    Private Sub b_date_Click(sender As Object, e As EventArgs) Handles b_date.Click
        txt_Answers.Clear()
        senderRemote("/date")
    End Sub


    Private Sub b_platform_Click(sender As Object, e As EventArgs) Handles b_platform.Click
        txt_Answers.Clear()
        senderRemote("/platform")

    End Sub
    Private Sub b_InfoCPU_Click(sender As Object, e As EventArgs) Handles b_InfoCPU.Click
        txt_Answers.Clear()
        senderRemote("/infoCPU")
    End Sub

    Private Sub b_processList_Click(sender As Object, e As EventArgs) Handles b_processList.Click
        txt_Answers.Clear()
        senderRemote("/processList")
    End Sub
    Private Sub b_Mute_Click(sender As Object, e As EventArgs) Handles b_Mute.Click
        txt_Answers.Clear()
        senderRemote("/Mute")
    End Sub

    Private Sub b_Volumen_UP_Click(sender As Object, e As EventArgs) Handles b_Volumen_UP.Click
        txt_Answers.Clear()
        senderRemote("/VolumenUP")
    End Sub

    Private Sub b_Volumen_Down_Click(sender As Object, e As EventArgs) Handles b_Volumen_Down.Click
        txt_Answers.Clear()
        senderRemote("/VolumenDown")
    End Sub

    Private Sub b_LogOut_Click(sender As Object, e As EventArgs) Handles b_LogOut.Click
        txt_Answers.Clear()
        senderRemote("/Logout")
    End Sub

    Private Sub b_ShutDown_Click(sender As Object, e As EventArgs) Handles b_ShutDown.Click
        txt_Answers.Clear()
        senderRemote("/ShutDown")
    End Sub

    Private Sub b_Restart_Click(sender As Object, e As EventArgs) Handles b_Restart.Click
        txt_Answers.Clear()
        senderRemote("/Restart")
    End Sub


    Private Sub Menu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
    End Sub


End Class
