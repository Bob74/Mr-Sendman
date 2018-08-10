Imports System.Management

Public Class Wmi

    ' Liste des versions de build WMI
    Public Const WMIBuild_W2000 As Double = 1085.0005
    Public Const WMIBuild_XP As Double = 2600.0
    Public Const WMIBuild_Vista As Double = 6002.18005
    Public Const WMIBuild_Seven As Double = 7601.17514

    Public Class ScopeID
        Public host As String = "127.0.0.1"
        Public username As String = ""
        Public password As String = ""
        Public domain As String = ""
        Public path As String = "\root\CIMV2"

        Public Sub reset()
            host = "127.0.0.1"
            username = ""
            password = ""
            domain = ""
            path = "\root\CIMV2"
        End Sub

    End Class

    ''' <summary>
    ''' Récupère des informations via une requête WQL.
    ''' Retourne un ArrayList de tableaux. Les tableaux contiennent les informations relatives à chaque éléments (1 élément avec ses propriétés = 1 tableau).
    ''' </summary>
    ''' <param name="_scID">Informations d'identification pour la création de la ManagementScope</param>
    ''' <param name="_req">Requête type WQL</param>
    ''' <param name="_auth">Niveau d'identification. Certains services nécessitent des droits différents (ex : terminal service requiert PacketPrivacy ( = 6))</param>
    ''' <returns>Retourne un ArrayList de tableaux. Les tableaux contiennent les informations relatives à chaque éléments (1 élément avec ses propriétés = 1 tableau).</returns>
    ''' <remarks></remarks>
    Public Function GetQuery(ByVal _scID As ScopeID, ByVal _req As String, Optional ByVal _auth As Integer = 0) As ArrayList
        ' -----------------------------------------------------------------
        ' | Renvois une ArrayList de Arrays
        ' |     -> Chaque Array contient les informations d'un service
        ' |         -> Chaque information est un objet
        ' |
        ' | Les informations entrées dans un Array ainsi que leur ordre
        ' | sont définis par la clause SELECT de la requête
        ' -----------------------------------------------------------------

        Dim _arCont As New ArrayList
        Dim _select As String()

        Try
            _select = Split(Replace(Mid(Mid(_req, 8), 1, _req.IndexOf(" FROM") - 7), " ", ""), ",")
            Dim scope = SetConnection(_scID, _auth)
            scope.Connect()

            Dim query As New ObjectQuery(_req)
            Dim searcher As New ManagementObjectSearcher(scope, query)

            If searcher.Get.Count <> 0 Then
                For Each queryObj As ManagementObject In searcher.Get()

                    If UBound(_select) = 0 Then
                        _arCont.Add(queryObj(_select(0)))
                    Else

                        Dim _ar As New ArrayList

                        For i = 0 To UBound(_select)
                            _ar.Add(queryObj(_select(i)))
                        Next

                        _arCont.Add(_ar)

                    End If

                Next
            End If

        Catch err As ManagementException
#If DEBUG Then
            MsgBox("WMI (GetQuery [" & _req & "]) : " & err.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        Catch unauthorizedErr As System.UnauthorizedAccessException
#If DEBUG Then
            MsgBox("WMI (GetQuery [" & _req & "]) : " & unauthorizedErr.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Accès refusé")
#End If
        Catch ex As Exception
#If DEBUG Then
            MsgBox("WMI (GetQuery [" & _req & "]) : " & ex.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        End Try

        Return _arCont

    End Function

    ''' <summary>
    ''' Appelle une méthode WMI sans arguments d'entrée.
    ''' </summary>
    ''' <param name="_scID">Informations d'identification pour la création de la ManagementScope</param>
    ''' <param name="_manObj">Pointe sur l'objet avec lequel on veut intéragir. Ex : ManagObj = Win32_Service.Name='nom du service'</param>
    ''' <param name="_method">Méthode à éxécuter</param>
    ''' <param name="_auth">Niveau d'identification. Certains services nécessitent des droits différents (ex : terminal service requiert PacketPrivacy ( = 6))</param>
    ''' <remarks></remarks>
    Public Function CallMethod(ByVal _scID As ScopeID, ByVal _manObj As String, ByVal _method As String, Optional ByVal _auth As Integer = 0) As ManagementBaseObject
        '-----------------------------------------------------------------------
        '| _scID comporte toutes les infos pour créer la scope (user, poste à contacter, etc.)
        '| _manObj pointe sur l'objet avec lequel on veut intéragir (appliquer une méthode sur lui)
        '| _method désigne la fonction à éxécuter
        '| ---
        '| ex : ManagObj = Win32_Service.Name='<nom du service>'
        '|                  |=> pointe sur un service désigné par son nom (pour start/stop...)
        '----------------------------------------------------------------------- 

        Try

            Dim objManagementScope As ManagementScope
            Dim classinstance   ' Sera ManagementObject ou ManagementClass suivant le besoin

            ' Création de la scope et connexion :
            objManagementScope = SetConnection(_scID)
            objManagementScope.Connect()

            ' On pointe ensuite sur l'objet avec lequel on veut intéragir :
            If InStr(_manObj, ".") Then
                ' Si on a un "point", on pointe sur un objet (ex : Win32_Process.Name)
                classinstance = New ManagementObject _
                (objManagementScope, New ManagementPath(_manObj), Nothing)
            Else
                ' Sinon on pointe sur une classe (ex : Win32_Process)
                classinstance = New ManagementClass _
                (objManagementScope, New ManagementPath(_manObj), Nothing)
            End If

            ' Exécuter la méthode et récupèrer les paramètres de sortie :
            Dim outParams As ManagementBaseObject = classinstance.InvokeMethod(_method, Nothing, Nothing)

            Return outParams

        Catch err As ManagementException
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & err.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        Catch unauthorizedErr As System.UnauthorizedAccessException
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & unauthorizedErr.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Accès refusé")
#End If
        Catch ex As Exception
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & ex.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        End Try

        Return Nothing

    End Function

    ''' <summary>
    ''' Appelle une méthode WMI avec arguments d'entrée.
    ''' </summary>
    ''' <param name="_scID">Informations d'identification pour la création de la ManagementScope</param>
    ''' <param name="_manObj">Pointe sur l'objet avec lequel on veut intéragir
    ''' ex : ManagObj = Win32_Service.Name='nom du service'
    ''' </param>
    ''' <param name="_method">Méthode à éxécuter</param>
    ''' <param name="_params">Arguments d'entrée (Syntaxe : {{"Paramètre", "Valeur"}}).
    ''' ex : Params(,) a 2 dimensions avec comme valeurs à l'index 0 ["StartMode"]["Auto"] (On assigne la valeur "Auto" au paramètre "StartMode". C'est un tableau à 2 colonnes.)</param>
    ''' <param name="_auth">Niveau d'identification. Certains services nécessitent des droits différents (ex : terminal service requiert PacketPrivacy ( = 6))</param>
    ''' <remarks></remarks>
    Public Function CallMethod(ByVal _scID As ScopeID, ByVal _manObj As String, ByVal _method As String, ByVal _params As Object(,), Optional ByVal _auth As Integer = 0) As ManagementBaseObject
        '-----------------------------------------------------------------------
        '| _scID comporte toutes les infos pour créer la scope (ex : user, poste à contacter, etc.)
        '| _manObj pointe sur l'objet avec lequel on veut intéragir (appliquer une méthode sur lui)
        '| _method désigne la fonction à exécuter
        '| _params contient les paramètres d'entrée (ex : type de démarrage à appliquer au service)
        '| ---
        '| ex : ManagObj = Win32_Service.Name='<nom du service>'
        '|                 |=> pointe sur un service désigné par son nom (pour start/stop...)
        '|
        '| ex : Params(,) à 2 dimensions avec comme valeurs ["StartMode"]["Auto"]
        '|          |=> inParams(Params(0,0)) = Params(0,1)
        '|          |==> Ligne 0 -> Première case : Paramètre
        '|          |==> Ligne 0 -> Deuxième case : Valeur
        '|          |===> On assigne la valeur "Auto" au paramètre "StartMode"
        '----------------------------------------------------------------------- 

        Try
            Dim objManagementScope As ManagementScope
            Dim objManagementBaseObject As ManagementBaseObject
            Dim classinstance   ' Sera ManagementObject ou ManagementClass suivant le besoin

            ' Création de la scope et connexion :
            objManagementScope = SetConnection(_scID)
            objManagementScope.Connect()

            ' On pointe ensuite sur l'objet avec lequel on veut intéragir :
            If InStr(_manObj, ".") Then
                ' Si on a un "point", on pointe sur un objet (ex : Win32_Process.Name)
                classinstance = New ManagementObject _
                (objManagementScope, New ManagementPath(_manObj), Nothing)
            Else
                ' Sinon on pointe sur une classe (ex : Win32_Process)
                classinstance = New ManagementClass _
                (objManagementScope, New ManagementPath(_manObj), Nothing)
            End If

            With classinstance
                .Scope = objManagementScope
                ' Obtention des paramètres d'entrée attendus par la Méthode :
                objManagementBaseObject = .GetMethodParameters(_method)

                With objManagementBaseObject
                    For i = 0 To UBound(_params)
                        .SetPropertyValue(_params(i, 0), _params(i, 1))
                    Next
                End With

                ' Exécuter la méthode et récupèrer les paramètres de sortie :
                Dim OutParams As ManagementBaseObject = .InvokeMethod(_method, objManagementBaseObject, Nothing)

                Return OutParams
            End With

        Catch err As ManagementException
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & err.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        Catch unauthorizedErr As System.UnauthorizedAccessException
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & unauthorizedErr.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Accès refusé")
#End If
        Catch ex As Exception
#If DEBUG Then
            MsgBox("WMI (CallMethod [" & _manObj & ", " & _method & "]) : " & ex.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        End Try

        Return Nothing

    End Function

    ''' <summary>
    ''' Retourne les entrées Ouverture/Fermeture/Déverrouillage de session du journal d'évênement Sécurité sur une période donnée.
    ''' </summary>
    ''' <param name="_scID">Informations d'identification pour la création de la ManagementScope</param>
    ''' <param name="dateFrom">Date de début</param>
    ''' <param name="dateTo">Date de fin</param>
    ''' <param name="deverrou">Indique s'il faut ou non afficher les déverrouillages de session</param>
    ''' <returns>Retourne un ArrayList de tableaux. Les tableaux contiennent les informations relatives à chaque éléments (1 élément avec ses propriétés = 1 tableau).</returns>
    ''' <remarks></remarks>
    Public Function GetSessionEventLog(ByVal _scID As ScopeID,
                                        ByVal dateFrom As Date,
                                        ByVal dateTo As Date,
                                        Optional ByVal deverrou As Boolean = False)

        Dim _arCont As New ArrayList

        Try

            Dim scope = SetConnection(_scID)    ' Création de la scope
            scope.Connect()                     ' Connexion..
            Dim query As New ObjectQuery("SELECT CategoryString, TimeGenerated, EventIdentifier, Message" &
                                         " FROM Win32_NTLogEvent" &
                                         " WHERE LogFile='Security'" &
                                         " AND TimeGenerated>='" & dateFrom & "'" &
                                         " AND TimeGenerated<='" & dateTo & "'")

            Dim searcher As New ManagementObjectSearcher(scope, query)

            For Each queryObj As ManagementObject In searcher.Get()
                ' On prend seulement les Ouvertures/Fermetures de session
                If queryObj("EventIdentifier") = 4624 Or queryObj("EventIdentifier") = 4647 Then

                    Dim _msglines As String() = Split(queryObj("Message"), vbCrLf)           ' Message découpé par lignes

                    Dim _str = Mid(queryObj("TimeGenerated").ToString, 1, 14)                                            ' Format retourné : 20130419125413.225811-000
                    Dim _LoggingDate As Date = Mid(_str, 7, 2) & "/" & Mid(_str, 5, 2) & "/" & Mid(_str, 1, 4) & " " &
                                               Mid(_str, 9, 2) & ":" & Mid(_str, 11, 2) & ":" & Mid(_str, 13, 2)           ' Remet en forme l'heure (format "lisible" + ajout de 2h)

                    Dim _ar As New ArrayList
                    Dim _UserAccount As String = ""
                    Dim _add As Boolean = False
                    Dim _type As String = ""        ' Type d'évênement (Ouverture/Fermeture/Déverrouillage)

                    ' Ouverture/Déverrouillage de session (4624)
                    If queryObj("EventIdentifier") = 4624 Then

                        ' winlogon.exe figure dans les évênements d'ouverture/déverrouillage de session
                        If InStr(_msglines(19), "winlogon.exe") <> 0 Then

                            If InStr(_msglines(8), "7") Then
                                ' Déverrouillage de session
                                _type = "Déverrouillage"
                                _add = True
                                If Not deverrou Then _add = False ' Si on ne veut pas de Déverrouillage
                            ElseIf InStr(_msglines(8), "2") Then
                                ' Ouverture de session (utilisateur ou system)
                                _type = "Ouverture"
                                _add = True
                                ' Si "GUID d’ouverture de session" = {00000000-0000-0000-0000-000000000000}
                                ' on ne prend pas l'évênement (ouverture de session System et non User)
                                If InStr(_msglines(15), "{00000000-0000-0000-0000-000000000000}") Then _add = False
                            ElseIf InStr(_msglines(8), "3") Then
                                ' A titre d'info
                                _type = "Réseau"
                            End If

                            If _add Then
                                _UserAccount = Split(_msglines(13), vbTab & vbTab)(1) & "\" & Split(_msglines(12), vbTab & vbTab)(1)  'DOMAIN\USER
                            End If

                        End If

                    Else
                        ' Fermeture de session (4647)
                        _add = True
                        _UserAccount = Split(_msglines(5), vbTab & vbTab)(1) & "\" & Split(_msglines(4), vbTab & vbTab)(1)  'DOMAIN\USER
                        _type = "Fermeture"

                    End If

                    If _add Then
                        Dim _NetUser = ""
                        If UBound(Split(_UserAccount, "\")) >= 1 Then ' S'il y'a un '\'
                            _NetUser = NetUser(Split(_UserAccount, "\")(1))

                            If Split(_NetUser, vbCrLf).Length <= 10 Then
                                ' Si le message fait moins de 10 lignes, c'est que la recherche n'a rien donné, on réessaye en local
                                _NetUser = NetUser(Split(_UserAccount, "\")(1), False)
                            End If
                        End If

                        ' Remplissage du tableau avec les infos
                        _ar.Add(queryObj("CategoryString"))
                        _ar.Add(_type)
                        _ar.Add(_UserAccount)
                        _ar.Add(_LoggingDate.AddHours(2).ToString())
                        _ar.Add(queryObj("Message"))
                        _ar.Add(_NetUser)

                        ' Ajout de cet évênement au tableau principal
                        _arCont.Add(_ar)

                    End If


                End If
            Next

        Catch err As ManagementException
#If DEBUG Then
            MsgBox("WMI (EventLog) : " & err.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur")
#End If
        Catch unauthorizedErr As System.UnauthorizedAccessException
#If DEBUG Then
            MsgBox("WMI (EventLog) : " & unauthorizedErr.Message, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Accès refusé")
#End If
        End Try

        Return _arCont

    End Function

    ''' <summary>
    ''' Génère la ManagementScope nécessaire à l'éxécution de la requête/méthode WMI. Retourne la ManagementScope à utiliser.
    ''' </summary>
    ''' <param name="_scID">Informations d'identification pour la création de la ManagementScope</param>
    ''' <param name="_auth">Niveau d'identification. Certains services nécessitent des droits différents (ex : terminal service requiert PacketPrivacy ( = 6))</param>
    ''' <returns>Retourne la ManagementScope à utiliser.</returns>
    ''' <remarks></remarks>
    Private Function SetConnection(ByVal _scID As ScopeID, Optional ByVal _auth As Integer = 0) As ManagementScope
        '-----------------------------------------------------------------------
        '| Création d'une scope avec les informations d'identification
        '| et pointant sur le poste ciblé.
        '-----------------------------------------------------------------------
        '| options.Authentication = certains services nécessitent des droits
        '| différents (ex : terminal service requiert PacketPrivacy ( = 6))
        '-----------------------------------------------------------------------

        Dim objManagementScope As New ManagementScope

        With objManagementScope
            With .Path
                .Server = _scID.host
                .NamespacePath = _scID.path
            End With

            With .Options
                .EnablePrivileges = True
                .Impersonation = ImpersonationLevel.Impersonate
                .Authentication = _auth

                If _scID.username <> "" Then      ' Si username n'est pas renseigné, on laisse les options vides
                    .Username = _scID.username
                    .Password = _scID.password
                    .Authority = "ntlmdomain:" & _scID.domain
                End If

            End With
        End With

        Return objManagementScope

    End Function

    ''' <summary>
    ''' Détaille les erreurs retournées par l'éxécution de méthodes WMI.
    ''' </summary>
    ''' <param name="outParams">Objet contenant la valeur de retour de la méthode</param>
    ''' <remarks></remarks>
    Private Sub ErrorOut(ByVal outParams As ManagementBaseObject)
        ' -------------------------------------------------------
        '| outParams("ReturnValue") :
        '| 0    ->  Success
        '| 2    ->  Access Denied
        '| 3    ->  Insufficient privilege
        '| 8    ->  Unknown failure
        '| 9    ->  Path not found
        '| 21   ->  Invalid parameter
        ' -------------------------------------------------------

        Dim err As Integer = Convert.ToInt32(outParams("returnValue"))
        Dim msg As String = "FAILED: An Unspecified error has occured (code: " & err & ")"

        Select Case err
            Case 0
                msg = "SUCCESS"
            Case 2
                msg = "FAILED: Access denied (code: " & err & ")"
            Case 3
                msg = "FAILED: Insufficient privilege (code: " & err & ")"
            Case 8
                msg = "FAILED: Unknown failure (code: " & err & ")"
            Case 9
                msg = "FAILED: Path not found (code: " & err & ")"
            Case 21
                msg = "FAILED: Invalid parameter (code: " & err & ")"
        End Select

        ' Pas de message si la fonction s'est bien déroulée (err <> 0) :
        If err <> 0 Then MsgBox(msg, MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, "Erreur lors de l'appel de la méthode")

    End Sub

End Class




