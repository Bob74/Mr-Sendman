
' Objet "Hôte" passé dans le thread
Public Class TargetObj
    Public Sub New(Optional ByVal host As String = "127.0.0.1", Optional ByVal msg As String = "", Optional ByVal duration As Integer = 0)
        Name = host
        Message = msg
        MsgDuration = duration
    End Sub
    Public Sub ClearData()
        WmiBuild = 0.0
        IP = ""
        Message = ""
        Status = ""
        User = ""
        Image = ""
        Color = Brushes.Black
        ErrorCode = 0
    End Sub
    Public Property ScopeID As Wmi.ScopeID
    Public Property WmiBuild As Double = 0.0
    Public Property Name As String
    Public Property IP As String
    Public Property Message As String
    Public Property MsgDuration As Integer
    Public Property Status As String
    Public Property User As String
    Public Property Image As String
    Public Property Color As SolidColorBrush
    Public Property ErrorCode As Integer = 0
    Public Property Counted As Boolean = True
End Class


' Objet "Liste d'hôtes" passé dans le thread
Public Class TargetListObj
    Public Sub New(ByVal list As ArrayList, Optional ByVal duration As Integer = 0, Optional ByVal msg As String = "")
        For i = 0 To list.Count - 1
            TargetList.Add(New TargetObj(list(i), msg, duration))
        Next
        MsgDuration = duration
    End Sub
    Public Sub New(ByVal list As ArrayList, Optional ByVal duration As Integer = 0, Optional ByVal retry As Boolean = False, Optional ByVal delay As Integer = 5)
        For i = 0 To list.Count - 1
            TargetList.Add(New TargetObj(list(i), "", duration))
        Next
        RetryDelay = delay
        NeedRetry = retry
        MsgDuration = duration
    End Sub

    Public Sub SetMessage(ByVal msg As String)
        For i = 0 To TargetList.Count - 1
            TargetList(i).Message = msg
        Next
    End Sub
    Public Sub SetCounted(ByVal b As Boolean)
        For i = 0 To TargetList.Count - 1
            TargetList(i).Counted = b
        Next
    End Sub
    Public Sub ClearTargetsData()
        For i = 0 To TargetList.Count - 1
            TargetList(i).ClearData()
        Next
    End Sub
    Public Function GetOfflineCount() As Integer
        Dim offlineCount As Integer = 0

        For i = 0 To TargetList.Count - 1
            If TargetList(i).ErrorCode = -1 Or TargetList(i).ErrorCode = -3 Then
                offlineCount += 1
            End If
        Next

        Return offlineCount

    End Function

    Public Property TargetList As New List(Of TargetObj)
    Public Property MsgDuration As Integer
    Public Property NeedRetry As Boolean
    Public Property RetryDelay As Integer = 5
End Class

' Contenu de listView : Liste des destinataires
Public Class HostInformation
    Public Sub New(ByVal n As String, ByVal t As Integer)
        Name = n
        Type = t
        If Type = 1 Then Image = "resources/range24.png" Else Image = "resources/host24.png"
    End Sub
    Public Property Name As String      ' Nom d'hôte ou plage d'adresse
    Public Property Type As Integer     ' 0 = Nom d'hôte / 1 = Plage d'adresse
    Public Property Image As String     ' Icône dans la listView
End Class

' Contenu de listView : Résultats
Public Class ResultInformation
    Public Property Name As String      ' Nom d'hôte
    Public Property IP As String        ' Adresse IP
    Public Property Status As String    ' État de l'hôte (Ping, WMI, Création du processus, etc)
    Public Property Message As String   ' Message délivré
    Public Property User As String      ' Utilisateur connecté sur la machine
    Public Property Image As String
    Public Property Color As SolidColorBrush
    Public Property ErrorCode As Integer = 0
    Public Property Time As String
End Class

' Contenu de listView : Liste des horaires
Public Class TimingInformations
    Public Sub New(ByVal d As Date, ByVal msg As String)
        TimingDate = d
        Name = "Envoi le " + d.ToString
        Message = msg
    End Sub
    Public Property TimingDate As Date  ' Heure de déclenchement (Valeur)
    Public Property Name As String      ' Heure de déclenchement (String)
    Public Property Message As String   ' Message à afficher
End Class

