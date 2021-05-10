Imports System.Net.Mail
Public Class Form1
    Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    Dim contador As Integer
    Dim flag As Boolean


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        If TextBox2.Text = "00" And TextBox3.Text = "00" Then
            MsgBox("You must to select the frecuency.", MsgBoxStyle.Information, "Select frecuency")
        Else
            Timer1.Interval = 1
            Button1.Enabled = False
            cmbHost.Enabled = False
            GroupBox1.Enabled = False
            ubiWinFilesTxt.Enabled = False
            Button4.Enabled = True
            Button2.Enabled = False
            Timer1.Start()
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        FolderBrowserDialog1.SelectedPath = Trim(ubiWinFilesTxt.Text)
        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()
        If (result = DialogResult.OK) Then
            ubiWinFilesTxt.Text = Trim(FolderBrowserDialog1.SelectedPath)
            Me.SaveSettings()
        End If

    End Sub

    Private Sub SaveSettings()
        Utils.SalvarSetting(appName, "UbiPWRBFilesLocation", ubiWinFilesTxt.Text)
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim errCode As Integer
        'Dim errorStr As String = ""
        'Dim hostNumberArray(5) As Integer

        Dim todayDate As Date
        Dim cdcStr As String

        'Utils.getHostMode("10.1.5.11", "prosys", "Numb3r1j0b", hostNumberArray, errCode, errorStr)
        'ComboBox1.Text = hostNumberArray(0)


        Me.GetSettings()

        TextBox2.Text = Utils.GetSetting(appName, "HoursFrecuency", "").ToString()
        TextBox3.Text = Utils.GetSetting(appName, "MinutesFrecuency", "").ToString()
        yesRadioButton.Select()

        todayDate = Format(Now, "MM/dd/yyyy")
        If TimeOfDay >= "00:00:01" And TimeOfDay < "03:55:00" Then
            cdcStr = returnCDC(todayDate) - 1
        Else
            cdcStr = returnCDC(todayDate)
        End If
        Label3.Text = cdcStr
        Timer1.Enabled = False
        Timer2.Enabled = False
        Label4.Text = ""
        Button4.Enabled = False

    End Sub

    Private Sub GetSettings()
        ubiWinFilesTxt.Text = Utils.GetSetting(appName, "UbiPWRBFilesLocation", "").ToString()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        'Dim objShell
        Dim strPathRemoteFile, strPathFile, hostStr As String
        Dim errMessageStr As String = ""
        Dim Success As Boolean = False
        Dim resInt As Integer
        Dim todayDate As Date
        Dim cdcStr As String
        Dim minInt, segInt, hourInt As Integer
        Dim errCode As Integer
        Dim errorStr As String = ""
        Dim conStrinHostStr As String
        Dim substrings() As String

        Timer1.Stop()
        Timer2.Stop()
        hostStr = Trim(cmbHost.Text)

        todayDate = Format(Now, "MM/dd/yyyy")
        If TimeOfDay >= "00:00:01" And TimeOfDay < "03:55:00" Then
            cdcStr = returnCDC(todayDate) - 1
        Else
            cdcStr = returnCDC(todayDate)
        End If
        Label3.Text = cdcStr

        'substrings(0) -> username
        'substrings(1) -> IP ESTE1
        'substrings(2) -> Password ESTE1
        'substrings(3) -> IP ESTE2
        'substrings(4) -> Password ESTE2
        'substrings(5) -> IP ESTE3
        'substrings(6) -> Password ESTE3
        'substrings(7) -> IP ESTE4
        'substrings(8) -> Password ESTE4
        'substrings(9) -> IP ESTE5
        'substrings(10) -> Password ESTE5

        conStrinHostStr = GetSettingConfigHost(appNameSettings, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")

        Select Case hostStr

            Case ""
                MsgBox("You must to select a Host number", MsgBoxStyle.Information, "Select Host Number")

                Button1.Enabled = True
                cmbHost.Enabled = True
                ubiWinFilesTxt.Enabled = True
                Button1.Enabled = True
                Button4.Enabled = False

                Exit Sub


            Case "1"

                Utils.autoMJFDumpsNew(substrings(1), substrings(0), substrings(2), errCode, errorStr)

                If errCode <> 0 Then
                    MsgBox("Error: " & errorStr, MsgBoxStyle.Information, "Error")
                Else
                    If (yesRadioButton.Checked) Then
                        strPathRemoteFile = "/d2/mjfxfer"
                        strPathFile = Trim(ubiWinFilesTxt.Text)
                        resInt = downLoadFileDumpsSFTP(substrings(1), "22", "xfer", "Welcome1", strPathRemoteFile, strPathFile, errMessageStr)
                        If resInt <> 0 Then
                            MsgBox("Error: " & errMessageStr, MsgBoxStyle.Information, "Error")
                            'Else
                            '    Me.TextBox1.Text = Me.TextBox1.Text & Now & ".- The files were transfered succefully." & vbCrLf
                            '    Me.Refresh()
                        End If
                    End If

                End If

            Case "2"

                Utils.autoMJFDumpsNew(substrings(3), substrings(0), substrings(4), errCode, errorStr)

                If errCode <> 0 Then
                    MsgBox("Error: " & errorStr, MsgBoxStyle.Information, "Error")
                Else
                    If (yesRadioButton.Checked) Then
                        strPathRemoteFile = "/d2/mjfxfer"
                        strPathFile = Trim(ubiWinFilesTxt.Text)
                        resInt = downLoadFileDumpsSFTP(substrings(3), "22", "xfer", "Welcome1", strPathRemoteFile, strPathFile, errMessageStr)
                        If resInt <> 0 Then
                            MsgBox("Error: " & errMessageStr, MsgBoxStyle.Information, "Error")
                            'Else
                            '    Me.TextBox1.Text = Me.TextBox1.Text & Now & ".- The files were transfered succefully." & vbCrLf
                            '    Me.Refresh()
                        End If
                    End If

                End If




            Case "3"

                Utils.autoMJFDumpsNew(substrings(5), substrings(0), substrings(6), errCode, errorStr)

                If errCode <> 0 Then
                    MsgBox("Error: " & errorStr, MsgBoxStyle.Information, "Error")
                Else
                    If (yesRadioButton.Checked) Then
                        strPathRemoteFile = "/d2/mjfxfer"
                        strPathFile = Trim(ubiWinFilesTxt.Text)

                        resInt = downLoadFileDumpsSFTP(substrings(5), "22", "xfer", "Welcome1", strPathRemoteFile, strPathFile, errMessageStr)
                        If resInt <> 0 Then
                            MsgBox("Error: " & errMessageStr, MsgBoxStyle.Information, "Error")
                            'Else
                            '    Me.TextBox1.Text = Me.TextBox1.Text & Now & ".- The files were transfered succefully." & vbCrLf
                            '    Me.Refresh()
                        End If
                    End If

                End If



            Case "4"

                Utils.autoMJFDumpsNew(substrings(7), substrings(0), substrings(8), errCode, errorStr)

                If errCode <> 0 Then
                    MsgBox("Error: " & errorStr, MsgBoxStyle.Information, "Error")
                Else
                    If (yesRadioButton.Checked) Then
                        strPathRemoteFile = "/d2/mjfxfer"
                        strPathFile = Trim(ubiWinFilesTxt.Text)
                        resInt = downLoadFileDumpsSFTP(substrings(7), "22", "xfer", "Welcome1", strPathRemoteFile, strPathFile, errMessageStr)
                        If resInt <> 0 Then
                            MsgBox("Error: " & errMessageStr, MsgBoxStyle.Information, "Error")
                            'Else
                            '    Me.TextBox1.Text = Me.TextBox1.Text & Now & ".- The files were transfered succefully." & vbCrLf
                            '    Me.Refresh()
                        End If
                    End If

                End If


            Case "5"

                Utils.autoMJFDumpsNew(substrings(9), substrings(0), substrings(10), errCode, errorStr)

                If errCode <> 0 Then
                    MsgBox("Error: " & errorStr, MsgBoxStyle.Information, "Error")
                Else
                    If (yesRadioButton.Checked) Then
                        strPathRemoteFile = "/d2/mjfxfer"
                        strPathFile = Trim(ubiWinFilesTxt.Text)
                        resInt = downLoadFileDumpsSFTP(substrings(9), "22", "xfer", "Welcome1", strPathRemoteFile, strPathFile, errMessageStr)
                        If resInt <> 0 Then
                            MsgBox("Error: " & errMessageStr, MsgBoxStyle.Information, "Error")
                            'Else
                            '    Me.TextBox1.Text = Me.TextBox1.Text & Now & ".- The files were transfered succefully." & vbCrLf
                            '    Me.Refresh()
                        End If
                    End If

                End If


        End Select

        minInt = CInt(TextBox3.Text)
        segInt = minInt * 60

        hourInt = CInt(TextBox2.Text)
        segInt = segInt + (hourInt * 3600)

        Timer1.Interval = segInt * 1000
        Timer1.Start()
        contador = segInt
        Timer2.Start()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Timer1.Stop()
        Timer2.Stop()
        Label4.Text = ""
        Button1.Enabled = True
        cmbHost.Enabled = True
        GroupBox1.Enabled = True
        ubiWinFilesTxt.Enabled = True
        Button4.Enabled = False
        Button2.Enabled = True


    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Dim dayWeekInt As Integer
        Dim timeofdayMilitary As String
        contador = contador - 1
        Label4.Text = contador

        dayWeekInt = Weekday(Now)
        timeofdayMilitary = Format(TimeOfDay, "HH:mm:ss")
        Label9.Text = timeofdayMilitary

        Select Case dayWeekInt

            Case 1 'Sunday

            Case 2 'Monday 
                If timeofdayMilitary = "20:45:15" Then
                    Timer1.Stop()
                    Timer2.Stop()
                    Label4.Text = ""
                    Timer1.Interval = 1
                    Timer1.Start()

                End If

            Case 5 'Thursday

                If timeofdayMilitary = "20:45:15" Then
                    Timer1.Stop()
                    Timer2.Stop()
                    Label4.Text = ""
                    Timer1.Interval = 1
                    Timer1.Start()

                End If


            Case 3, 6 'Tuesday and Friday

                If timeofdayMilitary = "22:45:15" Then
                    Timer1.Stop()
                    Timer2.Stop()
                    Label4.Text = ""
                    Timer1.Interval = 1
                    Timer1.Start()

                End If

            Case 4, 7 'Wednesday and saturday

                If timeofdayMilitary = "22:00:15" Then
                    Timer1.Stop()
                    Timer2.Stop()
                    Label4.Text = ""
                    Timer1.Interval = 1
                    Timer1.Start()

                End If

        End Select



    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Dispose()
    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

    '    Dim resCode As Integer

    '    resCode = sent_Email("156.24.14.132", "carlos.vegabello@igt.com, Kavita.Persaud@IGT.com", "C:\SharedFiles\Vacation_form.doc", "Prueba numero 1", "Prueba Hola Mundo")

    '    If resCode = 0 Then
    '        MessageBox.Show("Done!  Hola", "Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '    Else

    '        MessageBox.Show("Error!", "Error Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End If

    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim f As New Form2(Me)
        f.ShowDialog()
    End Sub

    Private Sub yesRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yesRadioButton.CheckedChanged
        Label5.Visible = True
        Button3.Visible = True
        ubiWinFilesTxt.Visible = True
    End Sub

    Private Sub noRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles noRadioButton.CheckedChanged
        Label5.Visible = False
        Button3.Visible = False
        ubiWinFilesTxt.Visible = False
    End Sub
End Class
