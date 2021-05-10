
Module Sftp

    Function downloadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        Dim handle As String

        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 1
        End If


        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.


        success = sftp.Connect(hostName, port)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 2
        Else
            'MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
            'Console.WriteLine("Exito")
        End If


        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 3
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 4
        Else
            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")
        End If


        handle = sftp.OpenFile(remoteFile, "readOnly", "openExisting")
        If (handle = vbNullString) Then
            strError = sftp.LastErrorText
            Return 5
        Else
            'MsgBox("Funciono abrir el archivo", MsgBoxStyle.OkOnly, "Funciono abrir el archivo. ")
        End If

        '  Download the file:
        success = sftp.DownloadFile(handle, localFile)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 6
        Else
            'MsgBox("Funciono bajar el archivo", MsgBoxStyle.OkOnly, "Funciono bajar el archivo. ")

        End If


        '  Close the file.
        success = sftp.CloseHandle(handle)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 7

        Else
            'MsgBox("Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "Funciono cerrar el archivo. ")
        End If
        Return 0

    End Function



    Function upLoadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 1
        End If

        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.

        success = sftp.Connect(hostName, port)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 2
        Else

        End If


        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 3
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 4
        Else
            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")
        End If

        success = sftp.UploadFileByName(remoteFile, localFile)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 5
        End If

        Return 0

        'MsgBox("Funciono todo", MsgBoxStyle.OkOnly, "EXITO TOTAL")


    End Function

    Function downLoadFileDumpsSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal strPathRemoteFile As String, ByVal strPathLocalFile As String, ByRef strError As String) As Integer


        Dim sftp As New Chilkat.SFtp
        Dim dirListing As Chilkat.SFtpDir
        Dim handle, serieStr, cdcStr As String
        Dim numFilesInt As Integer
        Dim todayDate As Date

        '  Any string automatically begins a fully-functional 30-day trial.
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 1
        End If


        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.

        success = sftp.Connect(hostName, port)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 2

        End If


        '  Authenticate with the SSH server.  Chilkat SFTP supports
        '  both password-based authenication as well as public-key
        '  authentication.  This example uses password authenication.
        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 3
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 4
        Else
            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")
        End If


        '  Open a directory on the server...
        '  Paths starting with a slash are "absolute", and are relative
        '  to the root of the file system. Names starting with any other
        '  character are relative to the user's default directory (home directory).
        '  A path component of ".." refers to the parent directory,
        '  and "." refers to the current directory.

        handle = sftp.OpenDir(strPathRemoteFile)
        If (sftp.LastMethodSuccess <> True) Then
            strError = sftp.LastErrorText
            Return 5
        Else

        End If


        '  Download the directory listing:

        dirListing = sftp.ReadDir(handle)
        If (sftp.LastMethodSuccess <> True) Then
            strError = sftp.LastErrorText
            Return 6
        Else

        End If




        '  Iterate over the files.
        Dim i As Integer
        Dim conRemote As Integer = dirListing.NumFilesAndDirs
        
        'MsgBox("Hay " & Str(conRemote) & " archivos en la parte remota.")

        todayDate = Format(Now, "MM/dd/yyyy")
        
        cdcStr = Trim(Form1.Label3.Text)

        If (conRemote <> 0) Then
            numFilesInt = 0
            For i = 0 To conRemote - 1
                serieStr = CStr(i + 1).PadLeft(3, "0")
                If (sftp.FileExists(strPathRemoteFile & "/" & "mjf_dump" & cdcStr & "v" & serieStr & ".fil", True) = 1) Then
                    numFilesInt = numFilesInt + 1
                End If
            Next

            If numFilesInt > 1 Then
                If numFilesInt = conRemote Then
                    For i = 0 To numFilesInt - 2
                        serieStr = CStr(i + 1).PadLeft(3, "0")
                        If Not (My.Computer.FileSystem.FileExists(strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")) Then
                            Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- Transferring 'mjf_dump" & cdcStr & "v" & serieStr & ".fil' from " & strPathRemoteFile & " to " & strPathLocalFile & vbCrLf
                            Form1.Refresh()
                            success = sftp.DownloadFileByName(strPathRemoteFile & "/" & "mjf_dump" & cdcStr & "v" & serieStr & ".fil", strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")
                            If (success <> True) Then
                                strError = sftp.LastErrorText
                                Return 7
                            Else
                                Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully." & vbCrLf
                                Form1.Refresh()
                                'MsgBox("mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully.")
                            End If

                        End If
                    Next
                Else
                    If conRemote > numFilesInt And cdcStr <> returnCDC(todayDate) Then

                        For i = 0 To numFilesInt - 1
                            serieStr = CStr(i + 1).PadLeft(3, "0")
                            If Not (My.Computer.FileSystem.FileExists(strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")) Then
                                Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- Transferring 'mjf_dump" & cdcStr & "v" & serieStr & ".fil' from " & strPathRemoteFile & " to " & strPathLocalFile & vbCrLf
                                Form1.Refresh()
                                success = sftp.DownloadFileByName(strPathRemoteFile & "/" & "mjf_dump" & cdcStr & "v" & serieStr & ".fil", strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")
                                If (success <> True) Then
                                    strError = sftp.LastErrorText
                                    Return 7
                                Else

                                    Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully." & vbCrLf
                                    Form1.Refresh()
                                    'MsgBox("mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully.")
                                End If

                            End If
                        Next

                    Else
                        If conRemote > numFilesInt And cdcStr = returnCDC(todayDate) Then
                            For i = 0 To numFilesInt - 2
                                serieStr = CStr(i + 1).PadLeft(3, "0")
                                If Not (My.Computer.FileSystem.FileExists(strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")) Then
                                    Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- Transferring 'mjf_dump" & cdcStr & "v" & serieStr & ".fil' from " & strPathRemoteFile & " to " & strPathLocalFile & vbCrLf
                                    Form1.Refresh()
                                    success = sftp.DownloadFileByName(strPathRemoteFile & "/" & "mjf_dump" & cdcStr & "v" & serieStr & ".fil", strPathLocalFile & "mjf_dump" & cdcStr & "v" & serieStr & ".fil")
                                    If (success <> True) Then
                                        strError = sftp.LastErrorText
                                        Return 7
                                    Else
                                        Form1.TextBox1.Text = Form1.TextBox1.Text & Now & ".- mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully." & vbCrLf
                                        Form1.Refresh()
                                        'MsgBox("mjf_dump" & cdcStr & "v" & serieStr & ".fil" & " was transfered succefully.")
                                    End If

                                End If
                            Next

                        End If



                    End If


                End If


            End If
        End If


        '  Close the directory
        success = sftp.CloseHandle(handle)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 8
        Else
            Return 0
        End If

    End Function



End Module
