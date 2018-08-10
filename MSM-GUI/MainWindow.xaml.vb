' *********************************************************************************************************************
' *
' *     Projet commencé le 18/08/2017
' *
' *********************************************************************************************************************
'
' 10/10/2017 2.0.1 :    - Affichage du rappel de l'envoi dans l'onglet Envoi après l'élévation
'                       - Ajout d'une colonne Adresse IP dans les résultats
'
' 16/09/2017 2.0.0 :    - Ajout de la planification des envois
'                       - Possibilité de faire disparaître le message après X secondes (même durée pour l'ensemble des messages envoyés)
'                       - Sauvegarde des données en XML (Elévation + Import/Export)
'                       - Données utilisateurs supplémentaires (Retry et Duration : checkbox + value)
'                       - Fonction d'update de la ProgressBar
'
' 29/08/2017 1.1.0 :    - Ajout de /Time:0 pour MSG.exe
'
' 28/08/2017 1.0.0 :    - Ajout d'icônes pour la listview des résultats
'                       - Correction d'un "bug" dans la classe WMI : Les résultats des Query WMI retournaient
'                         obligatoirement un tableau non vide (.Count = 1), ce qui empêchait de vérifier de manière
'                         sûre si la requête WMI avait échouée
'                       - Modification du design des boutons "Ajouter" pour les destinataires (design rond)
'                       - Ajout de deux boutons "Importer" et "Exporter" pour permettre de charger/sauvegarder les
'                         listes de destinataires
'                       - Sauvegarde des plages d'adresses IP dans les paramètres Utilisateur
'
' 25/08/2017 0.3.0 :    - Réécriture au propre de l'interface graphique en XAML
'                       - ListView custom avec icônes dans la liste
'
' 24/08/2017 0.2.0 :    - Utilisation de la librairie "IPAddressRange" pour générer les plages d'adresses
'                       - Création de l'icône "Vault boy"
'                       - Destinataires supprimables de la liste via la touche "Suppr"
'
' 23/08/2017 0.1.0 :    - Interface MahApps Metro terminée
'                       - Multi-threading en place
'                       - Wmi incorporé
'                       - Rechargement de l'interface lors du passage en Admin (fichier de conf dans ProgramData)
'                       - Manque encore la génération de plages d'adresses
'

'   A FAIRE :
'       - Logger les envois sous forme de fichier Excel
'       - Tester sur Windows XP
'       - Assembly version (automatiser la modification ?)
'
'       + Permettre de programmer et répéter les envoi :
'           - Notification quand envoi message en bas à droite
'           - Poste ayant reçu le message seulement / Tous les postes
'
'


' Rappel concernant les Threads :
' -------------------------------
' On communique entre le Thread principal et les multithreads via un objet que l'on ballade de fonctions en fonctions.
' Cet objet peut être passé par valeur (ByVal) car il est déjà présent sous la forme de pointeur (constructeur New).
' Pas besoin donc de le passer par référence (ByRef).
'


Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Net
Imports Microsoft.Win32
Imports NetTools
Imports System.ComponentModel
Imports System.Xml

Class MainWindow
    Inherits MetroWindow

    ' Version du programme
    Private Property _version As String = "2.0.1"

    Public Property Hosts As ObservableCollection(Of HostInformation)
    Public Property HostsResults As ObservableCollection(Of ResultInformation)
    Public Property TimingMsg As ObservableCollection(Of TimingInformations)

    Private Property IsCurrentUserAdmin As Boolean = False
    Private Property tempFile As String = "C:\ProgramData\MSM_tmp.cfg"

    ' Instance de la classe WMI :
    Private Property _Wmi As New Wmi

    ' Variable pour permettre d'utiliser l'ObservableCollection depuis des threads
    Private ReadOnly HostsResultsLock As New Object

    ' Mutex pour empêcher la modification simultanée de la liste des résultats
    Private Property mutex As New Threading.Mutex


#Region "Core & Threading"

    Private Sub StartSending()
        Dim TPool As Threading.ThreadPool

        ' Compteur d'envois traités
        labelResultCount.Content = "0 / " & Hosts.Count

        If checkboxProgMsg.IsChecked Then
            If TimingMsg.Count > 0 Then
                Dim targetListObj As TargetListObj = New TargetListObj(GenerateTargetList(), CInt(NumericUpDownMsgDuration.Value), CheckboxRetry.IsChecked, CInt(NumericUpDownRetryDelay.Value))
                TPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf DetectTimeAndSend), targetListObj)
                Return
            End If
        End If

        ' Génération de la liste des hôtes et inclusion du message dans le TargetListObj
        SendToList(New TargetListObj(GenerateTargetList(), CInt(NumericUpDownMsgDuration.Value), textBoxMsgMessage.Text.Replace("""", "\""")))
    End Sub


    Private Sub DetectTimeAndSend(ByVal targetList As TargetListObj)
        ' Enumération des heures dans l'ordre de saisie
        For i = 0 To TimingMsg.Count - 1
            ' Reset du log
            HostsResults.Clear()
            labelResultCount.Invoke(Sub() labelResultCount.Content = "0 / " & Hosts.Count)

            ' On patiente jusqu'à l'heure demandée
            While Now < TimingMsg(i).TimingDate
                Sleep(1000)
            End While

            ' Oh ! Il est 9h !
            targetList.ClearTargetsData()               ' On supprime les données reçues par le dernier envoi (état, icône, errorcode, etc.)
            targetList.SetMessage(TimingMsg(i).Message) ' Définir le message pour l'envoi en cours...
            SendToList(targetList)                      ' Et envoi !

            ' Tentatives sur les hôtes qui n'ont pas répondus
            If targetList.NeedRetry Then
                ' Si il reste un horaire à traiter
                If TimingMsg.Count - 1 >= i + 1 Then
                    ' Tant qu'on a pas atteint l'horaire suivant, on réessaye les postes offline

                    ' On patiente jusqu'à ce que les envois soient tous terminés
                    While Hosts.Count < targetList.TargetList.Count
                        Sleep(1000)
                    End While

                    ' On ajoute RetryDelay "+ 1" pour éviter que ça passe pile et que le temps de réponse d'un poste offline rende le résultat après l'enclenchement de
                    ' l'horaire suivant. On se retrouverait avec le résultat du précédent horaire dans la nouvelle liste.
                    While Now.AddMinutes(targetList.RetryDelay + 1) < TimingMsg(i + 1).TimingDate
                        ' Attente avant nouvel envoi
                        Sleep(targetList.RetryDelay * 60000)

                        ' Nouvel essais
                        Dim offlineList As TargetListObj = New TargetListObj(GenerateTargetList("ResultInformation", True), targetList.MsgDuration, targetList.NeedRetry, targetList.RetryDelay)
                        offlineList.SetMessage(TimingMsg(i).Message)
                        offlineList.SetCounted(False)
                        SendToList(offlineList)

                        ' On patiente jusqu'à ce que les envois soient tous terminés
                        While offlineList.GetOfflineCount() < offlineList.TargetList.Count
                            Sleep(1000)
                        End While

                    End While
                End If
            End If
        Next
    End Sub


    ' Envoi du message à tous les destinataires
    Private Sub SendToList(ByVal obj As TargetListObj)
        If obj.TargetList.Count > 0 Then
            Dim TPool As Threading.ThreadPool

            For i = 0 To obj.TargetList.Count - 1
                ' Mettre une tâche en file d'attente
                TPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf WorkingThread), obj.TargetList(i))
            Next
        End If
    End Sub

    Private Sub WorkingThread(ByVal obj As TargetObj)
        If obj.Name <> "" Then
            obj.ScopeID = GetScopeID(obj.Name)

            Dim targetState As Integer = IsTargetAvailable(obj)
            obj.Status = ErrorToString(targetState)

            If targetState = 1 Then
                ExecWmiOnTarget(obj)
            End If

            mutex.WaitOne()

            ' Ajout du résultat dans la listView
            HostsResults.Add(New ResultInformation() With {
                        .Name = obj.Name,
                        .IP = obj.IP,
                        .Status = obj.Status,
                        .Message = obj.Message,
                        .User = obj.User,
                        .Image = obj.Image,
                        .Color = obj.Color,
                        .ErrorCode = obj.ErrorCode,
                        .Time = Now.ToString()})

            ' Incrémente le compteur d'envoi traités
            If obj.Counted Then
                labelResultCount.Invoke(Sub() labelResultCount.Content = HostsResults.Count & " / " & Hosts.Count)
            End If

            mutex.ReleaseMutex()
        End If
    End Sub



    ' Génération de la liste depuis HostInformation (nouvel envoi) ou ResultInformation (réenvoi aux hôtes non disponibles)
    Private Function GenerateTargetList(Optional ByVal hostSource As String = "HostInformation", Optional ByVal offlineOnly As Boolean = True) As ArrayList
        Dim targetList As New ArrayList

        ' Nouvel envoi, on prend les hôtes ciblés
        If hostSource = "HostInformation" Then
            For Each h As HostInformation In Hosts
                If h.Type = 0 Then
                    ' Ajout direct quand c'est un nom d'hôte/IP
                    targetList.Add(h.Name)
                Else
                    For Each ip As IPAddress In IPAddressRange.Parse(h.Name)
                        Dim target As String = ip.ToString
                        If Not target.EndsWith(".255") And Not target.EndsWith(".0") Then
                            targetList.Add(ip.ToString)
                        End If
                    Next
                End If
            Next
        Else
            ' Sinon on prend les hôtes dans la liste des résultats

            ' Hôtes hors ligne uniquement
            If offlineOnly Then
                For Each h As ResultInformation In HostsResults
                    ' Pas de résolution DNS / Pas de ping
                    If h.ErrorCode = -1 Or h.ErrorCode = -3 Then
                        If Not targetList.Contains(h.Name) Then targetList.Add(h.Name)  ' Ajout sans doublons
                    End If
                Next
            Else
                ' Tous les hôtes des résultats
                For Each h As ResultInformation In HostsResults
                    If Not targetList.Contains(h.Name) Then targetList.Add(h.Name)  ' Ajout sans doublons
                Next
            End If

        End If

        Return targetList

    End Function


    ' Retourne si la cible est valide (Ping + Wmi)
    Private Function IsTargetAvailable(ByVal obj As TargetObj) As Integer

        If Not ISValidIPv4(obj.Name) Then   ' Si c'est un nom DNS, on résout
            Dim targetDns As String = DnsToIP(obj.Name)
            If targetDns = "" Then
                obj.Color = Brushes.Gray    ' Impossible de résoudre le nom d'hôte
                obj.ErrorCode = -1
                Return -1
            Else
                obj.IP = targetDns
            End If
        Else
            obj.IP = obj.Name
        End If

        ' => A partir d'ici, IP = adresse IP

        If Ping(obj.IP, 500, 1) Then
            ' On vérifie que WMI est bien accessible :
            Dim WmiBuildArray As ArrayList = _Wmi.GetQuery(obj.ScopeID, "SELECT BuildVersion FROM Win32_WMISetting")

            If WmiBuildArray.Count > 0 Then
                obj.WmiBuild = Convert.ToDouble(Replace(WmiBuildArray(0), ".", ","))
            Else
                obj.Image = "resources/lock.png"
                obj.Color = Brushes.Orange  ' Pas de WMI
                obj.ErrorCode = -2
                Return -2
            End If
        Else
            obj.Color = Brushes.Gray        ' Ne ping pas
            obj.ErrorCode = -3
            Return -3
        End If

        Return 1

    End Function


    Private Sub ExecWmiOnTarget(ByVal obj As TargetObj)
        ' Utilisateur connecté
        Dim curuserArray As ArrayList = _Wmi.GetQuery(obj.ScopeID, "SELECT UserName FROM Win32_ComputerSystem")
        If curuserArray.Count > 0 Then obj.User = curuserArray(0)

        ' Création du message
        Dim cmd As String = ""

        If obj.WmiBuild > Wmi.WMIBuild_XP Then
            cmd = "MSG * /Time:" & obj.MsgDuration & " " & obj.Message
        Else
            cmd = "net send " & obj.Name & " " & obj.Message
        End If

        Dim outParams As Management.ManagementBaseObject = Nothing
        outParams = _Wmi.CallMethod(obj.ScopeID, "Win32_Process", "Create", {{"CommandLine", cmd}})

        If Not outParams Is Nothing Then
            ' Récupération de "ReturnValue"
            Dim prop As Management.PropertyData = outParams.Properties.Item("ReturnValue")
            If TypeOf (prop.Value) Is Array Then
                For Each subitem In prop.Value
                    If Not prop.Value Is Nothing Then
                        obj.Status += subitem.ToString() & " / "
                    End If
                Next

                obj.Color = Brushes.Green
                obj.ErrorCode = 1
                obj.Image = "resources/sendOk.png"
                obj.Status = "Message envoyé"
            Else
                If prop.Value = 0 Then
                    obj.Color = Brushes.Green
                    obj.ErrorCode = 1
                    obj.Image = "resources/sendOk.png"
                    obj.Status = "Message envoyé"
                Else
                    obj.Color = Brushes.Red
                    obj.ErrorCode = -4
                    obj.Image = "resources/sendNok.png"
                    obj.Status = "Erreur : " & prop.Value
                End If
            End If
        Else
            obj.Color = Brushes.Red
            obj.ErrorCode = -5
            obj.Image = "resources/sendNok.png"
            obj.Status = "Échec lors de la création du processus distant."
        End If

    End Sub


    ' Génère un ManagementClass avec les arguments entrés
    Public Function CreateManagementClass(ByVal managementPath As String, args As String()) As Management.ManagementClass
        Dim result As Management.ManagementClass = New Management.ManagementClass(managementPath)

        For i = 0 To args.Length - 1 Step 2
            result(args(i)) = args(i + 1)
        Next

        Return result
    End Function

    ' Création d'une ScopeID avec les paramètres courants
    Private Function GetScopeID(Optional ByVal host As String = "", Optional ByVal path As String = "\root\CIMV2") As Wmi.ScopeID
        Dim ScopeID As New Wmi.ScopeID
        ScopeID.host = host
        ScopeID.path = path
        Return ScopeID
    End Function

    ' Traduit le code erreur en message
    Private Function ErrorToString(ByVal id As Integer) As String
        Select Case id
            Case -1
                Return "Impossible de résoudre le nom d'hôte"
            Case -2
                Return "Accès Wmi refusé"
            Case -3
                Return "Ne ping pas"
        End Select

        Return "Ok"

    End Function
#End Region


#Region "Événements UI"
#Region " -------> Menu"
    ' Modification du menu du bas automatique en fonction de l'onglet en cours
    Private Sub metroAnimatedTabControlMain_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles metroAnimatedTabControlMain.SelectionChanged
        If Me.IsLoaded Then
            Dim currentTab As Integer = metroAnimatedTabControlMain.SelectedIndex
            If currentTab = 0 Then          ' Aide
                labelUAC.Visibility = Visibility.Hidden
                progressBarMenu.Visibility = Visibility.Visible
                buttonNext.Visibility = Visibility.Visible
                labelNext.Visibility = Visibility.Visible
                labelNext.Content = "Commencer"
            ElseIf currentTab = 1 Then      ' Message
                labelUAC.Visibility = Visibility.Hidden
                progressBarMenu.Visibility = Visibility.Visible
                buttonNext.Visibility = Visibility.Visible
                labelNext.Visibility = Visibility.Visible
                labelNext.Content = "Suivant"
            ElseIf currentTab = 2 Then      ' Destinataires
                labelUAC.Visibility = Visibility.Hidden
                progressBarMenu.Visibility = Visibility.Visible
                buttonNext.Visibility = Visibility.Visible
                labelNext.Visibility = Visibility.Visible
                labelNext.Content = "Suivant"
            ElseIf currentTab = 3 Then      ' Envoi
                If Not IsCurrentUserAdmin Then labelUAC.Visibility = Visibility.Visible
                progressBarMenu.Visibility = Visibility.Visible
                buttonNext.Visibility = Visibility.Visible
                labelNext.Visibility = Visibility.Visible
                labelNext.Content = "Envoyer !"
            ElseIf currentTab = 4 Then      ' Résultats
                labelUAC.Visibility = Visibility.Hidden
                progressBarMenu.Visibility = Visibility.Hidden
                buttonNext.Visibility = Visibility.Hidden
                labelNext.Visibility = Visibility.Hidden
                labelNext.Content = ""
            End If
        End If
    End Sub

    ' Bouton Suivant
    Private Sub labelNext_MouseEnter(sender As Object, e As MouseEventArgs) Handles labelNext.MouseEnter
        labelNext.Foreground = Brushes.Black
    End Sub
    Private Sub labelNext_MouseLeave(sender As Object, e As MouseEventArgs) Handles labelNext.MouseLeave
        labelNext.Foreground = New SolidColorBrush(Color.FromRgb(69, 69, 69))
    End Sub
    Private Sub labelNext_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles labelNext.MouseLeftButtonUp
        buttonNext_Click(sender, e)
    End Sub
    Private Sub buttonNext_Click(sender As Object, e As RoutedEventArgs) Handles buttonNext.Click
        Dim currentTab As Integer = metroAnimatedTabControlMain.SelectedIndex
        If currentTab = 0 Then          ' Aide
            HelpTabNext()
        ElseIf currentTab = 1 Then      ' Message
            MessageTabNext()
        ElseIf currentTab = 2 Then      ' Destinataires
            DestTabNext()
        ElseIf currentTab = 3 Then      ' Envoi
            SendTabNext()
        End If
    End Sub


    ' Onglet Aide - Bouton "Commencer"
    Private Sub HelpTabNext()
        Dim nextTab As TabItem = metroAnimatedTabControlMain.Items.GetItemAt(1)
        nextTab.IsEnabled = True
        metroAnimatedTabControlMain.SelectedItem = nextTab
        progressBarUpdate()     ' 30%
    End Sub
    ' Onglet Message - Bouton "Suivant"
    Private Sub MessageTabNext()
        If listViewTiming.Items.Count <= 0 Then
            checkboxProgMsg.IsChecked = False
        End If

        If textBoxMsgMessage.Text <> "" Or TimingMsg.Count > 0 Then
            Dim nextTab As TabItem = metroAnimatedTabControlMain.Items.GetItemAt(2)
            nextTab.IsEnabled = True
            metroAnimatedTabControlMain.SelectedItem = nextTab
        Else
            ShowMessageAsync("Information", "Vous devez saisir un message à afficher.")
        End If
    End Sub
    ' Onglet Destinataires - Bouton "Suivant"
    Private Sub DestTabNext()
        If Not listViewDest.HasItems Then
            ShowMessageAsync("Information", "Vous devez saisir au moins un destinataire pour le message." & vbCrLf &
                             "Un destinataire se désigne via son adresse IP ou son nom d'hôte.")
        Else
            Dim nextTab As TabItem = metroAnimatedTabControlMain.Items.GetItemAt(3)
            nextTab.IsEnabled = True
            metroAnimatedTabControlMain.SelectedItem = nextTab
            progressBarUpdate()     ' 100%
        End If
    End Sub
    'Onglet Envoi - Bouton "Envoyer"
    Private Sub SendTabNext()
#If DEBUG Then
        IsCurrentUserAdmin = True
#End If
        If IsCurrentUserAdmin Then
            ' Reset des résultats précédents
            HostsResults.Clear()

            ' Bascule sur l'onglet caché pour afficher les résultats (dernier onglet)
            Dim tabCount As Integer = metroAnimatedTabControlMain.Items.Count()
            Dim resultTab As TabItem = metroAnimatedTabControlMain.Items.GetItemAt(tabCount - 1)
            resultTab.Header = "Résultats"                          ' Affichage de l'onglet caché
            resultTab.Focusable = True
            resultTab.IsHitTestVisible = True
            metroAnimatedTabControlMain.SelectedItem = resultTab    ' Focus sur l'onglet

            ' Envoi multi-threadé
            StartSending()
        Else
            SaveFormData()
            RestartElevated()
        End If
    End Sub
#End Region
#Region " -------> Événements"
    ' Onglet - Message
    Private Sub textBoxMsgMessage_TextChanged(sender As Object, e As TextChangedEventArgs) Handles textBoxMsgMessage.TextChanged
        refreshEndMessage()
        labelMsgCharLimit.Content = textBoxMsgMessage.Text.Length.ToString & " / 255"

        If labelMsgCharLimit.Content = "255 / 255" Then
            labelMsgCharLimit.Foreground = Brushes.Red
        Else
            labelMsgCharLimit.Foreground = Brushes.Black
        End If

        progressBarUpdate()     ' 30% ou 50%
    End Sub
    Private Sub checkboxProgMsg_Checked(sender As Object, e As RoutedEventArgs) Handles checkboxProgMsg.Checked
        gridProgMsg.Visibility = Visibility.Visible
        gridProgMsg.IsEnabled = True
    End Sub
    Private Sub checkboxProgMsg_Unchecked(sender As Object, e As RoutedEventArgs) Handles checkboxProgMsg.Unchecked
        gridProgMsg.IsEnabled = False
        gridProgMsg.Visibility = Visibility.Hidden
    End Sub
    Private Sub buttonProgAddTime_Click(sender As Object, e As RoutedEventArgs) Handles buttonProgAddTime.Click
        If Not timePickerProgMsg.SelectedTime Is Nothing Then
            If textBoxMsgMessage.Text <> "" Then
                TimingMsg.Add(New TimingInformations(timePickerProgMsg.SelectedDate, textBoxMsgMessage.Text))
                textBoxMsgMessage.Clear()
                progressBarUpdate()     ' 50%
            Else
                ShowMessageAsync("Information", "Vous devez saisir un message à afficher.")
            End If
            refreshEndMessage()
        End If
    End Sub
    ' Supprimer un horaire
    Private Sub listViewTiming_KeyUp(sender As Object, e As KeyEventArgs) Handles listViewTiming.KeyUp
        If e.Key = Key.Delete Then
            If listViewTiming.SelectedIndex >= 0 Then
                While listViewTiming.SelectedItems().Count > 0
                    Dim n As Integer = listViewTiming.SelectedItems().Count - 1
                    Dim lastitem = listViewTiming.SelectedItems().Item(n)
                    TimingMsg.RemoveAt(listViewTiming.Items.IndexOf(lastitem))
                End While
            End If
            If TimingMsg.Count = 0 Then progressBarUpdate() ' 30%
            refreshEndMessage()
        End If
    End Sub
    ' Clic sur un horaire
    Private Sub listViewTiming_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles listViewTiming.SelectionChanged
        If listViewTiming.SelectedIndex >= 0 Then
            textBoxMsgMessage.Text = TimingMsg(listViewTiming.SelectedIndex).Message
        End If
    End Sub
    ' Refresh du message récapitulatif des envoi
    Private Sub refreshEndMessage()
        If TimingMsg.Count = 0 Then
            textBlockEndMessage.Text = textBoxMsgMessage.Text
        Else
            textBlockEndMessage.Text = ""
            For Each h As TimingInformations In TimingMsg
                textBlockEndMessage.Text += h.Name & " : " & h.Message & vbCrLf & vbCrLf
            Next
        End If
    End Sub


    ' Update de la ProgressBar
    Private Sub progressBarUpdate()
        Dim currentTab As Integer = metroAnimatedTabControlMain.SelectedIndex
        If currentTab = 0 Then          ' Aide
            progressBarMenu.Value = 0
        ElseIf currentTab = 1 Then      ' Message
            If textBoxMsgMessage.Text <> "" Or TimingMsg.Count > 0 Then
                progressBarMenu.Value = 50
            Else
                progressBarMenu.Value = 30
            End If
        ElseIf currentTab = 2 Then      ' Destinataires
            If listViewDest.Items.Count > 0 Then
                progressBarMenu.Value = 80
            Else
                progressBarMenu.Value = 50
            End If
        ElseIf currentTab = 3 Then      ' Envoi
            progressBarMenu.Value = 100
        ElseIf currentTab = 4 Then      ' Résultats
            progressBarMenu.Value = 100
        End If
    End Sub

    Private Sub buttonDestAddHost_Click(sender As Object, e As RoutedEventArgs) Handles buttonDestAddHost.Click
        For Each host As String In textBoxDestHost.Text.Split(vbCrLf)
            host = host.Replace(vbCr, "")
            host = host.Replace(vbLf, "")

            If Not host Is Nothing And host <> "" Then
                Hosts.Add(New HostInformation(host, 0))
            End If
        Next

        progressBarUpdate()
        textBoxDestHost.Clear()
    End Sub

    Private Sub buttonDestAddRange_Click(sender As Object, e As RoutedEventArgs) Handles buttonDestAddRange.Click
        If ISValidIPv4(textBoxDestIPRangeFrom.Text) And ISValidIPv4(textBoxDestIPRangeTo.Text) Then
            Dim ipFrom As List(Of Integer) = IpStringToInteger(textBoxDestIPRangeFrom.Text)
            Dim ipTo As List(Of Integer) = IpStringToInteger(textBoxDestIPRangeTo.Text)

            If ipFrom.Count = 4 And ipTo.Count = 4 Then
                If CheckIpRange(ipFrom, ipTo) Then
                    Hosts.Add(New HostInformation(textBoxDestIPRangeFrom.Text & " - " & textBoxDestIPRangeTo.Text, 1))
                Else
                    ShowMessageAsync("Erreur", "La plage d'adresse spécifiée est incorrecte (l'adresse de fin de la plage est plus petite ou égale à l'adresse de départ).")
                End If
            Else
                ShowMessageAsync("Erreur", "Au moins une des adresses IP renseignée n'est pas valide.")
            End If

            progressBarUpdate() ' 50% ou 80%
        Else
            ShowMessageAsync("Erreur", "Au moins une des adresses IP renseignée n'est pas valide.")
        End If
    End Sub

    ' Supprimer un destinataire
    Private Sub listViewDest_KeyUp(sender As Object, e As KeyEventArgs) Handles listViewDest.KeyUp
        If e.Key = Key.Delete Then
            If listViewDest.SelectedIndex >= 0 Then

                While listViewDest.SelectedItems().Count > 0
                    Dim n As Integer = listViewDest.SelectedItems().Count - 1
                    Dim lastitem = listViewDest.SelectedItems().Item(n)
                    Hosts.RemoveAt(listViewDest.Items.IndexOf(lastitem))
                End While

            End If
        End If

        progressBarUpdate()
    End Sub

    ' Importer une liste de destinataires
    Private Sub buttonImportDest_Click(sender As Object, e As RoutedEventArgs) Handles buttonImportDest.Click
        ImportList()
    End Sub

    ' Exporter une liste de destinataires
    Private Sub buttonExportDest_Click(sender As Object, e As RoutedEventArgs) Handles buttonExportDest.Click
        ExportList()
    End Sub



    ' Onglet - Envoi
    Private Sub labelUAC_MouseLeftButtonUp(sender As Object, e As MouseButtonEventArgs) Handles labelUAC.MouseLeftButtonUp
        buttonNext_Click(sender, e)
    End Sub

    Private Sub buttonEndPreview_Click(sender As Object, e As RoutedEventArgs) Handles buttonEndPreview.Click
        Dim msgList As New List(Of String)

        If TimingMsg.Count = 0 Then
            Dim message = textBoxMsgMessage.Text
            message = message.Replace("""", "\""")
            msgList.Add(message)
        Else
            For Each h As TimingInformations In TimingMsg
                Dim message = h.Message
                message = message.Replace("""", "\""")
                msgList.Add(message)
            Next
        End If

        For Each m As String In msgList
            Try
                Process.Start(New ProcessStartInfo(GetSystemDir() & "\msg.exe", "* /Time:" & CInt(NumericUpDownMsgDuration.Value) & " " & m))
            Catch ex As Exception
                ShowMessageAsync("Erreur", ex.Message & vbCrLf & vbCrLf & ex.ToString)
            End Try
            Sleep(500)  ' Permet de faire arriver les messages dans l'ordre d'envoi
        Next
    End Sub
#End Region
#End Region


#Region "Fonctions UI"
    ' Chargement de l'interface
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ' Ajout du n° de version dans le titre
        Me.Title += " (" & _version & ")"

        ' Chargement des paramètres Utilisateur
        textBoxDestIPRangeFrom.Text = My.Settings.RangeFrom
        textBoxDestIPRangeTo.Text = My.Settings.RangeTo
        checkboxMsgDuration.IsChecked = My.Settings.Duration
        NumericUpDownMsgDuration.Value = My.Settings.DurationDelay
        CheckboxRetry.IsChecked = My.Settings.Retry
        NumericUpDownRetryDelay.Value = My.Settings.RetryDelay

        ' Masquage des options de programmation du message
        gridProgMsg.Visibility = Visibility.Hidden

        ' Masquage de l'icône UAC (n'est activé qu'après détection si compte admin ou non)
        labelUAC.Visibility = Visibility.Hidden

        Hosts = New ObservableCollection(Of HostInformation)()              ' Listview : Destinataires
        HostsResults = New ObservableCollection(Of ResultInformation)()     ' listView : Résultats
        TimingMsg = New ObservableCollection(Of TimingInformations)()       ' listview : Timings


        ' Autoriser la MAJ de la collection depuis les threads
        BindingOperations.EnableCollectionSynchronization(HostsResults, HostsResultsLock)
        DataContext = Me

        If Not IsAdmin() Then
            IsCurrentUserAdmin = False
            Try
                If File.Exists(tempFile) Then File.Delete(tempFile)
            Catch ex As Exception
                ShowMessageAsync("Erreur", ex.Message & vbCrLf & vbCrLf & ex.ToString)
            End Try
        Else
            IsCurrentUserAdmin = True
            If File.Exists(tempFile) Then
                LoadFormData()
                refreshEndMessage()

                ' Activation des Tabs
                For i = 0 To 3
                    Dim tab As TabItem = metroAnimatedTabControlMain.Items.GetItemAt(i)
                    tab.IsEnabled = True
                Next

                ' Et envoi des messages
                SendTabNext()
                progressBarUpdate()     ' 100%
            End If
        End If

    End Sub

    ' Fermeture de la fenêtre
    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        My.Settings.RangeFrom = textBoxDestIPRangeFrom.Text
        My.Settings.RangeTo = textBoxDestIPRangeTo.Text
        My.Settings.Duration = checkboxMsgDuration.IsChecked
        My.Settings.DurationDelay = NumericUpDownMsgDuration.Value
        My.Settings.Retry = CheckboxRetry.IsChecked
        My.Settings.RetryDelay = NumericUpDownRetryDelay.Value

        My.Settings.Save()
    End Sub

    ' Sauvegarde des données
    Private Sub SaveFormData()
        '   <data>
        '       <message>Le message à afficher !</message>
        '       <hostList>
        '           <host>
        '               <name>DSP0558405</name>
        '               <type>0</type>
        '           </host>
        '           <host>
        '               ...
        '           </host>
        '       </hostList>
        '       <timingList>
        '           <timing>
        '               <date>01/01/17 12:00:00</date>
        '               <message>Il est midi !</message>
        '           </timing>
        '           <timing>
        '               ...
        '           </timing>
        '       </timingList>
        '       <settings>
        '           <duration>0</duration>
        '           <retry>TRUE</retry>
        '           <retryDelay>5</retryDelay>
        '       </settings>
        '   </data>

        ' http://selkis.developpez.com/tutoriels/dotnet/Xmlpart1/
        Dim XmlDoc As XmlDocument = New XmlDocument()

        'création du document
        XmlDoc.LoadXml("<data></data>")

        ' Général
        Dim elemMsg As XmlElement                   ' - Message (dans le cas d'un envoi simple)
        elemMsg = XmlDoc.CreateElement("message")
        elemMsg.InnerText = textBoxMsgMessage.Text
        XmlDoc.DocumentElement.AppendChild(elemMsg)

        ' Liste des hôtes
        Dim elemHostList As XmlElement              ' + Listes des hôtes et leur type
        elemHostList = XmlDoc.CreateElement("hostList")
        XmlDoc.DocumentElement.AppendChild(elemHostList)

        For Each h As HostInformation In Hosts
            Dim elemHost As XmlElement                  '   + Hôte
            elemHost = XmlDoc.CreateElement("host")

            Dim elemHostName As XmlElement              '       - Nom d'hôte ou plage d'adresses
            elemHostName = XmlDoc.CreateElement("name")
            elemHostName.InnerText = h.Name
            elemHost.AppendChild(elemHostName)

            Dim elemHostType As XmlElement              '       - Type d'hôte
            elemHostType = XmlDoc.CreateElement("type")
            elemHostType.InnerText = h.Type.ToString()
            elemHost.AppendChild(elemHostType)

            elemHostList.AppendChild(elemHost)
        Next

        ' Liste des horaires
        Dim elemTimingList As XmlElement            ' + Listes des horaires
        elemTimingList = XmlDoc.CreateElement("timingList")
        XmlDoc.DocumentElement.AppendChild(elemTimingList)

        For Each t As TimingInformations In TimingMsg
            Dim elemTiming As XmlElement                '   + Horaire
            elemTiming = XmlDoc.CreateElement("timing")

            Dim elemTimingTime As XmlElement            '       - Date et heure
            elemTimingTime = XmlDoc.CreateElement("date")
            elemTimingTime.InnerText = t.TimingDate.ToString()
            elemTiming.AppendChild(elemTimingTime)

            Dim elemTimingMessage As XmlElement         '       - Message à afficher
            elemTimingMessage = XmlDoc.CreateElement("message")
            elemTimingMessage.InnerText = t.Message
            elemTiming.AppendChild(elemTimingMessage)

            elemTimingList.AppendChild(elemTiming)
        Next

        ' Settings
        Dim elemSettings As XmlElement              ' + Settings
        elemSettings = XmlDoc.CreateElement("settings")
        XmlDoc.DocumentElement.AppendChild(elemSettings)

        Dim elemSettingDuration As XmlElement       '   - Temps limité pour afficher le message ?
        elemSettingDuration = XmlDoc.CreateElement("duration")
        elemSettingDuration.InnerText = checkboxMsgDuration.IsChecked.ToString()
        elemSettings.AppendChild(elemSettingDuration)

        Dim elemSettingDurationDelay As XmlElement  '   - Durée d'affichage du message
        elemSettingDurationDelay = XmlDoc.CreateElement("durationDelay")
        elemSettingDurationDelay.InnerText = NumericUpDownMsgDuration.Value.ToString()
        elemSettings.AppendChild(elemSettingDurationDelay)

        Dim elemSettingProg As XmlElement           '   - Messages programmés ?
        elemSettingProg = XmlDoc.CreateElement("prog")
        elemSettingProg.InnerText = checkboxProgMsg.IsChecked.ToString()
        elemSettings.AppendChild(elemSettingProg)

        Dim elemSettingRetry As XmlElement          '   - Besoin de retry ?
        elemSettingRetry = XmlDoc.CreateElement("retry")
        elemSettingRetry.InnerText = CheckboxRetry.IsChecked.ToString()
        elemSettings.AppendChild(elemSettingRetry)

        Dim elemSettingRetryDelay As XmlElement     '   - Delay entre chaque essais
        elemSettingRetryDelay = XmlDoc.CreateElement("retryDelay")
        elemSettingRetryDelay.InnerText = NumericUpDownRetryDelay.Value.ToString()
        elemSettings.AppendChild(elemSettingRetryDelay)

        ' Sauvegarde
        XmlDoc.Save(tempFile)
    End Sub

    ' Restauration des données sauvegardées
    Private Sub LoadFormData()
        Try
            If File.Exists(tempFile) Then
                Dim xelement As XElement = XElement.Load(tempFile)
                Dim data As IEnumerable(Of XElement) = xelement.Elements()

                For Each d In data
                    If d.Name.LocalName = "message" Then
                        textBoxMsgMessage.Text = d.Value
                    ElseIf d.Name.LocalName = "hostList" Then
                        For Each e As XElement In d.Elements("host")
                            Hosts.Add(New HostInformation(e.Element("name").Value, CInt(e.Element("type").Value)))
                        Next
                    ElseIf d.Name.LocalName = "timingList" Then
                        For Each e As XElement In d.Elements("timing")
                            TimingMsg.Add(New TimingInformations(CDate(e.Element("date").Value), e.Element("message").Value))
                        Next
                    ElseIf d.Name.LocalName = "settings" Then
                        If Not d.Element("duration") Is Nothing Then checkboxMsgDuration.IsChecked = CBool(d.Element("duration").Value)
                        If Not d.Element("durationDelay") Is Nothing Then NumericUpDownMsgDuration.Value = CInt(d.Element("durationDelay").Value)
                        If Not d.Element("prog") Is Nothing Then checkboxProgMsg.IsChecked = CBool(d.Element("prog").Value)
                        If Not d.Element("retry") Is Nothing Then CheckboxRetry.IsChecked = CBool(d.Element("retry").Value)
                        If Not d.Element("retryDelay") Is Nothing Then NumericUpDownRetryDelay.Value = CInt(d.Element("retryDelay").Value)
                    End If
                Next

            End If
        Catch ex As Exception
            ShowMessageAsync("Erreur", ex.Message & vbCrLf & vbCrLf & ex.ToString)
        End Try

        If File.Exists(tempFile) Then File.Delete(tempFile)

    End Sub


    Private Sub ExportList()
        Dim listFile As String = ""
        Dim fileDialog As SaveFileDialog = New SaveFileDialog()
        fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        fileDialog.Filter = "Listes MSM (*.lst)|*.lst"

        If fileDialog.ShowDialog() Then
            listFile = fileDialog.FileName

            Try
                ' http://selkis.developpez.com/tutoriels/dotnet/Xmlpart1/
                Dim XmlDoc As XmlDocument = New XmlDocument()

                'création du document
                XmlDoc.LoadXml("<data></data>")

                ' Liste des hôtes
                Dim elemHostList As XmlElement              ' + Listes des hôtes et leur type
                elemHostList = XmlDoc.CreateElement("hostList")
                XmlDoc.DocumentElement.AppendChild(elemHostList)

                For Each h As HostInformation In Hosts
                    Dim elemHost As XmlElement                  '   + Hôte
                    elemHost = XmlDoc.CreateElement("host")

                    Dim elemHostName As XmlElement              '       - Nom d'hôte ou plage d'adresses
                    elemHostName = XmlDoc.CreateElement("name")
                    elemHostName.InnerText = h.Name
                    elemHost.AppendChild(elemHostName)

                    Dim elemHostType As XmlElement              '       - Type d'hôte
                    elemHostType = XmlDoc.CreateElement("type")
                    elemHostType.InnerText = h.Type.ToString()
                    elemHost.AppendChild(elemHostType)

                    elemHostList.AppendChild(elemHost)
                Next

                ' Sauvegarde
                XmlDoc.Save(listFile)
            Catch ex As Exception
                ShowMessageAsync("Erreur", ex.Message & vbCrLf & vbCrLf & ex.ToString)
            End Try

        End If

    End Sub

    Private Sub ImportList()
        Dim listFile As String = ""
        Dim fileDialog As OpenFileDialog = New OpenFileDialog()
        fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        fileDialog.Filter = "Listes MSM (*.lst)|*.lst"

        If fileDialog.ShowDialog() Then
            listFile = fileDialog.FileName

            Try
                If File.Exists(listFile) Then
                    Dim xelement As XElement = XElement.Load(listFile)
                    Dim data As IEnumerable(Of XElement) = xelement.Elements()
                    For Each d In data
                        If d.Name.LocalName = "hostList" Then
                            For Each e As XElement In d.Elements("host")
                                Hosts.Add(New HostInformation(e.Element("name").Value, CInt(e.Element("type").Value)))
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                ShowMessageAsync("Erreur", ex.Message & vbCrLf & vbCrLf & ex.ToString)
            End Try

            progressBarUpdate()
        End If

    End Sub
#End Region

End Class
