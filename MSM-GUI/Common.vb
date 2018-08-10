Imports System.Net

Module Common

    Declare Sub Sleep Lib "kernel32.dll" (ByVal Milliseconds As Integer)

    ''' <summary>
    ''' Ping d'une machine par son nom d'hôte ou son adresse IP. Retourne TRUE si l'hôte a répondu.
    ''' </summary>
    ''' <param name="host">Nom d'hôte ou adresse IP</param>
    ''' <param name="timeo">Time-out du ping en millisecondes (par défaut 500)</param>
    ''' <param name="nbretry">Nombre de tentatives en cas d'échec (par défaut 0)</param>
    ''' <returns>Retourne TRUE si l'hôte a répondu.</returns>
    ''' <remarks></remarks>
    Public Function Ping(ByVal host As String, Optional ByVal timeo As Integer = 500, Optional ByVal nbretry As Integer = 0) As Boolean

        Dim _re As Boolean

        For i = 0 To nbretry
            Try
                _re = My.Computer.Network.Ping(host, timeo)
                If _re Then
                    Exit For                                        'On quitte la boucle si on a une réponse
                Else
                    _re = False
                End If
            Catch ex As NetworkInformation.PingException 'Se déclenche si l'hote ne répond pas
                _re = False
                Exit Try
            Catch ex As InvalidOperationException            'Se déclenche s'il n'y a pas de réseau
                _re = False
                Exit Try
            Catch ex As Exception
                _re = False
                Exit Try
            End Try
        Next

        Return _re

    End Function

    ''' <summary>
    ''' Retourne une adresse IP depuis un nom DNS. Si le nom n'a pas pus être résolu, on retourne une chaine vide.
    ''' </summary>
    ''' <param name="host">Nom d'hôte de la machine ciblée</param>
    ''' <returns>Retourne l'adresse IP</returns>
    ''' <remarks></remarks>
    Public Function DnsToIP(ByVal host As String) As String
        Dim _ret As String = ""

        Try
            Dim hostEntry As IPHostEntry = Dns.GetHostEntry(host)
            If hostEntry.AddressList.Count > 0 Then
                _ret = hostEntry.AddressList(0).ToString()
            Else
                _ret = hostEntry.HostName
            End If
        Catch ex As Exception
#If DEBUG Then
            MsgBox("Fonction DnsToIP : " & ex.Message)
#End If
        End Try

        Return _ret

    End Function

    ''' <summary>
    ''' Vérifie si la chaîne entrée en argument est une adresse IPv4 valide.
    ''' </summary>
    ''' <param name="IPAddress">Adresse IP à tester</param>
    ''' <returns>Retourne TRUE si l'IP est valide</returns>
    ''' <remarks></remarks>
    Public Function ISValidIPv4(ByVal IPAddress As String) As Boolean
        Dim validFormat As String = "\b(([01]?\d?\d|2[0-4]\d|25[0-5])\.){3}([01]?\d?\d|2[0-4]\d|25[0-5])\b"
        Dim rxn As Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(validFormat)
        If rxn.IsMatch(IPAddress) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Vérifie si la chaîne entrée en argument est une adresse MAC valide.
    ''' </summary>
    ''' <param name="MACAddress">Adresse MAC à tester</param>
    ''' <returns>Retourne TRUE si l'adresse est valide</returns>
    ''' <remarks></remarks>
    Public Function ISValidMAC(ByVal MACAddress As String) As Boolean
        Dim validFormat As String = "^([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})$"
        Dim rxn As Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(validFormat)
        If rxn.IsMatch(MACAddress) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function NetUser(ByVal user As String, Optional ByVal domain As Boolean = True) As String
        ' Retrouve les informations sur l'utilisateur d'entrée via la commande "net user"
        ' Retourne "message" as string comme texte de la fenêtre CMD

        Dim arg As String = ""

        If domain Then arg = "/DOMAIN "

        Dim Proc As Process = New Process
        Dim message As String
        With Proc.StartInfo
            .StandardOutputEncoding = System.Text.Encoding.GetEncoding(850)
            .FileName = "c:\windows\system32\cmd.exe"
            .UseShellExecute = False
            .CreateNoWindow = True
            .RedirectStandardInput = True
            .RedirectStandardOutput = True
            .RedirectStandardError = True
        End With

        Proc.Start()

        Dim sIn As IO.StreamWriter = Proc.StandardInput
        Dim sOut As IO.StreamReader = Proc.StandardOutput
        Dim sErr As IO.StreamReader = Proc.StandardError

        sIn.AutoFlush = True
        sIn.Write("net user " & arg & user & Environment.NewLine)
        sIn.Write("exit" & Environment.NewLine)
        message = sOut.ReadToEnd()

        If Not Proc.HasExited Then
            Proc.Kill()
        End If
        sIn.Close()
        sOut.Close()
        sErr.Close()
        Proc.Close()

        Dim messageAr = Split(message, vbCrLf)
        message = ""

        For i = 4 To messageAr.Length - 3
            message = message & vbCrLf & messageAr(i)
        Next

        Return message

    End Function

    Public Function IpStringToInteger(ByVal ip As String) As List(Of Integer)
        If ISValidIPv4(ip) Then
            Dim ipInt As New List(Of Integer)

            For Each ipByte As String In ip.Split(".")
                ipInt.Add(Convert.ToInt32(ipByte))
            Next

            Return ipInt
        End If
        Return Nothing
    End Function

    Public Function CheckIpRange(ByVal rangeFrom As List(Of Integer), ByVal rangeTo As List(Of Integer)) As Boolean
        Dim isRangeValid As Boolean = False

        If rangeTo.ElementAt(0) > rangeFrom.ElementAt(0) Then
            isRangeValid = True
        ElseIf rangeTo.ElementAt(0) = rangeFrom.ElementAt(0) Then
            If rangeTo.ElementAt(1) > rangeFrom.ElementAt(1) Then
                isRangeValid = True
            ElseIf rangeTo.ElementAt(1) = rangeFrom.ElementAt(1) Then
                If rangeTo.ElementAt(2) > rangeFrom.ElementAt(2) Then
                    isRangeValid = True
                ElseIf rangeTo.ElementAt(2) = rangeFrom.ElementAt(2) Then
                    If rangeTo.ElementAt(3) > rangeFrom.ElementAt(3) Then
                        isRangeValid = True
                    End If
                End If
            End If
        End If

        Return isRangeValid

    End Function

    Public Function GetSystemDir() As String
        ' https://ovidiupl.wordpress.com/2008/07/11/useful-wow64-file-system-trick/
        ' /!\ Sur un système 64bits, Windows fait des redirections automatiques en fonction de l'architecture du programme (32 ou 64 bits)
        '     "C:\Windows\System32" est accessible aux applications 64bits
        '     "C:\Windows\SysWOW64" est accessible aux applications 32bits
        '     "C:\Windows\sysnative" est un lien qui existe pour les applications 32bits lancées sur un OS 64bits, le lien pointe sur les applications systèmes 64 bits depuis une application 32 bits


        '----------------------------------------
        ' Ex : Sur un Windows 7 64 bits, MSG.exe est une application 64 bits.
        '
        ' MSG.exe est réellement localisé dans le dossier C:\Windows\System32\ .
        ' Avec une application 32 bits, on le cherchera logiquement dans C:\Windows\System32\ mais on ne pourra pas le trouver.
        ' Windows nous aura déjà redirigé de manière transparente dans le dossier C:\Windows\SysWOW64 qui contient les applications systèmes 32 bits.
        '
        ' Comme MSG.exe n'existe qu'en 64 bits (sur Windows 7 x64 au moins), on doit utiliser le dossier C:\Windows\Sysnative pour accéder aux applications
        ' systèmes 64 bits depuis notre application 32 bits.
        '
        ' Ce qui est trompeur, c'est que l'explorateur Windows verra le fichier dans System32 car il est lui même en 64 bits et ne reflète donc pas les chemins
        ' et fichiers auxquels notre application 32 bits a accès.
        '

        If Environment.Is64BitOperatingSystem Then
            Return Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\sysnative"
        Else
            Return Environment.GetFolderPath(Environment.SpecialFolder.System)
        End If

    End Function

End Module
