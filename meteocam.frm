VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} meteocam 
   Caption         =   "UserForm1"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9045.001
   OleObjectBlob   =   "meteocam.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "meteocam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim olapp As New Outlook.Application
Dim NS As Object, dossier As Object, dossierArchives As Object, dossierMessagesEnvoyes As Object
Dim OlExp As Object
Dim i As Object
Dim mybody() As String
Dim fromsender As String
Dim ReceivedTime As Date
Dim j As Variant
Dim strFile As String
Dim strFolderpath As String
Dim objAttachments As Outlook.Attachments
Dim objMsg As Outlook.MailItem
Dim datejours As Date
Dim RecupDate As Date
Dim ReceivedTimecorrige As Variant
Dim datetampon(1) As Variant
Dim l As Integer
Dim NbrDateTrouve(1) As Integer
Dim k As Integer
Dim dossiersave As Variant
Dim lngCount As Integer

Dim largeur As Integer
Dim hauteur As Integer
Dim waitwaitTemps As Integer
Dim repertoire As String, fichier As String
Dim MyDestFolder As Variant

Dim compteur As Integer

Private Sub CBAnnule_Click()
    Unload Me
End Sub
Private Sub CBOK_Click()
'detecte si une date est rentrée et lance le prog LireMessage
'sinon lance le prog tridates pour toutes les dates

If TBDate <> "" Then
    RecupDate = TBDate.Value
    '*** Appel du programme de sauvegarde des photos de la date saisie
    Call LireMessages(RecupDate)
Else
    '*** Appel du programme de tri de toutes les dates presentes dans la boite de reception
    Call TriDates
End If
End Sub
Sub LireMessages(RecupDate)
    Dim MonApp As Outlook.Application
    Dim MonNameSpace As Outlook.NameSpace
    Dim MonDossier As Outlook.Folder
    Dim MonMail As Outlook.MailItem
    Dim numero As Integer
    'Dim strAttachment As String
    'Dim NbAttachments As Integer
    
    'Instance des objets
    Set MonApp = Outlook.Application
    Set MonNameSpace = MonApp.GetNamespace("MAPI")
    Set MonDossier = MonNameSpace.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")  'GetDefaultFolder(olFolderInbox)
    numero = MonDossier.Items.Count
    Set MonMail = MonDossier.Items(numero)
    Set MyDestFolder = MonDossier.Folders("Archives")
    
    compteur = 0
For Each i In MonDossier.Items
    ReceivedTime = Left(i.ReceivedTime, 10) ' isole la date
    ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
    dossiersave = TBChemin.Value + ReceivedTimecorrige + "\"
    Debug.Print ReceivedTime
    
    If ReceivedTime = RecupDate Then    ' compare la date du mail et la date du jour où est lancé la macro
        If Not Len(Dir(dossiersave, vbDirectory)) > 0 Then
            MkDir dossiersave
        End If
        
        compteur = compteur + 1 ' compte le nombre de mail sur lesquels on a agit
        Set objAttachments = i.Attachments
        lngCount = objAttachments.Count
        ' mettre la piece jointe dans le dossier
        For k = lngCount To 1 Step -1
            strFile = objAttachments.Item(k).FileName

            ' Combine with the path to the Temp folder.
            strFile = dossiersave & strFile

            ' Save the attachment as a file.
            objAttachments.Item(k).SaveAsFile strFile

        Next k
    End If


'i.Move MyDestFolder   ' fonctionne
Next i

Call getEmailsSelected(RecupDate)

End Sub
Sub getEmailsSelected(RecupDate)

    Dim oFolder As Folder
    Dim oItems As Items
    Dim j As Variant
    Dim NS As Object, dossier As Object
    Dim olapp As New Outlook.Application
    Dim MyDestFolder As Variant
    Dim M As Variant
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
 
    Set NS = olapp.GetNamespace("MAPI")
    Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")
    Set MyDestFolder = dossier.Folders("Archives")
    
    ActiveExplorer.ClearSelection
    
    Set oFolder = ActiveExplorer.CurrentFolder
    For Each j In dossier.Items
        ReceivedTime = Left(j.ReceivedTime, 10)
        If ReceivedTime = RecupDate Then
            ActiveExplorer.AddToSelection j
        End If
    Next j
    Set myOlExp = Application.ActiveExplorer
    
    Set myOlSel = myOlExp.Selection
    For Each M In myOlSel
        M.Move MyDestFolder
    Next M
End Sub



Private Sub CBTouteLesDates_Click()
If CBTouteLesDates = True Then
    TBDate.Value = ""
End If
End Sub

Private Sub CBVerification_Click()
If CBVerification.Value = True Then
    F_Tree_Recursif.Show
End If

End Sub

Private Sub TBDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
CtrL_KeyDown TBDate, KeyCode
End Sub

Private Sub CtrL_KeyDown(ByVal TxtB As MSForms.TextBox, ByVal KeyCode As MSForms.ReturnInteger)
    Dim X&, Xl&, D$, M$, A, T$, mask, C2, D2: mask = "__/__/____"
    'pour ceux qui n'ont pas le pavé numerique conversion du keycode du pavé haut du clavier
    If KeyCode >= 48 And KeyCode <= 57 Then KeyCode = KeyCode + 48
    'c'est parti on démarre le controle!!
    With TxtB
        Xl = .SelLength: If Xl = 0 Then Xl = 1    'Xl= la longeur de texte selectionné
        .Value = IIf(.Value = "", mask, .Value): If KeyCode = 8 And Xl > 1 Then KeyCode = 46
        T = .Value: .SelStart = IIf(T = mask, 0, .SelStart): X = .SelStart:
        Select Case KeyCode
        Case 96 To 105  'pavé numerique haut et bas (Attention!!!pas besoins de bloquer la touche MAJ!!!!!!!!le code se charge de convertir)
            If X = 10 Then KeyCode = 0: Exit Sub
            If X = 2 Or X = 5 Then X = X + 1
            Mid(T, X + 1, Xl) = Chr(KeyCode - 48) & Mid(mask, X + 2, Xl - 1)
            X = X + 1: Xl = 0: If X = 2 Or X = 5 Then X = X + 1
            'le plus gros tu traitement se passe avec controle de validité de date en fait!!!
            If Val(T) > 31 Or Val(Mid(T, 1, 1)) > 3 Then X = 0: Xl = 2: Mid(T, 1, 2) = Mid(mask, 1, 2): Beep
            If Val(Mid(T, 4, 2)) > 12 Or Val(Mid(T, 4, 1)) > 1 Then Mid(T, 4.2) = Mid(mask, 4, 2): X = 3: Xl = 2: Beep
            D = Mid(T, 1, 2): M = Mid(T, 4, 2): A = Mid(T, 7, 4)
            If IsDate(D & "/" & M) And Not IsDate(D & "/" & M & "/2000") Then Mid(T, 4, 2) = Mid(mask, 4, 2): X = 3: Xl = 2: Beep: KeyCode = 0
            If X = 10 And Not IsDate(T) Then Mid(T, 7, 10) = Mid(mask, 7, 10): X = 6: Xl = 4: Beep

        Case 8:
            If X = 0 Then KeyCode = 0:: .Value = "": Exit Sub Else Mid(T, X, 1) = Mid(mask, X, 1): X = X - 1: Xl = 0
            If T = mask Then T = ""
        Case 46:
            If X = 10 Then Exit Sub Else Mid(T, X + 1, Xl) = Mid(mask, X + 1, Xl): X = X: Xl = 0: If X = 2 Or X = 5 Then X = X + 1
            If T = mask Then T = ""
        Case Else: KeyCode = 0    ' a pour effet d'inhiber toutes les autre touches
        End Select
        .Value = T 'restitution
        .SelStart = X: .SelLength = Xl: KeyCode = 0
    End With
End Sub
Private Sub UserForm_Initialize()
largeur = 464 '171
hauteur = 221.25 '196.5

meteocam.Width = largeur
meteocam.Height = hauteur

CBTouteLesDates.Value = False
CBVerification.Value = False

End Sub
Sub TriDates()
Dim oFolder As Folder
Dim oItems As Items
Dim j As Variant
Dim NS As Object, dossier As Object
Dim olapp As New Outlook.Application
Dim MyDestFolder As Variant
Dim M As Variant
Dim myOlExp As Outlook.Explorer
Dim myOlSel As Outlook.Selection

Set NS = olapp.GetNamespace("MAPI")
Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")
Set dossierMessagesEnvoyes = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Messages envoyés")

l = 0
k = 0
dossiersave = TBChemin.Value
Set MyDestFolder = dossier.Folders("Archives")

If Not Len(Dir(dossiersave, vbDirectory)) > 0 Then
  MkDir dossiersave

End If

For Each i In dossier.Items
    ReceivedTime = Left(i.ReceivedTime, 10) ' isole la date
    ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
    
    dossiersave = TBChemin.Value + ReceivedTimecorrige + "\"
        
    If Not Len(Dir(dossiersave, vbDirectory)) > 0 Then
        MkDir dossiersave
    End If
        
    Set objAttachments = i.Attachments
    lngCount = objAttachments.Count
    ' mettre la piece jointe dans le dossier
    For k = lngCount To 1 Step -1
    strFile = objAttachments.Item(k).FileName

    ' Combine with the path to the Temp folder.
     strFile = dossiersave & strFile

    ' Save the attachment as a file.
     objAttachments.Item(k).SaveAsFile strFile
    Next k
Next i


ActiveExplorer.ClearSelection

Set oFolder = ActiveExplorer.CurrentFolder
For Each j In dossier.Items
    ReceivedTime = Left(j.ReceivedTime, 10)
    'If ReceivedTime = RecupDate Then
        ActiveExplorer.AddToSelection j
    'End If
Next j
Set myOlExp = Application.ActiveExplorer

Set myOlSel = myOlExp.Selection
For Each M In myOlSel
    M.Move MyDestFolder
Next M


' attention a ne pas recevoir un mail pendant le mouvement des mails car le nouveau mail va bouger aussi sans avoir pris la photo!
' dans ce cas comparer l'heure de reception avec l'heure du mail...


Unload Me
End Sub
