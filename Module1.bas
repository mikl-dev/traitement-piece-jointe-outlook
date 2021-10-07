Attribute VB_Name = "Module1"
'la procédure se lance à la réception d'un nouveau mail
Private Sub Application_NewMail()
Call sauvegardePJ
End Sub
'procédure de sauvegarde
Sub sauvegardePJ()
    Dim MonApp As Outlook.Application
    Dim MonNameSpace As Outlook.NameSpace
    Dim MonDossier As Outlook.Folder
    Dim MonMail As Outlook.MailItem
    Dim numero As Integer
    Dim strAttachment As String
    Dim NbAttachments As Integer
    Dim chemin As String
    'Instance des objets
    Set MonApp = Outlook.Application
    Set MonNameSpace = MonApp.GetNamespace("MAPI")
    Set MonDossier = MonNameSpace.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")  'GetDefaultFolder(olFolderInbox)
    numero = MonDossier.Items.Count
    Set MonMail = MonDossier.Items(numero)
    
    'chemin de destination des pièces jointes
    chemin = "C:\Users\marchessouxm\OneDrive - CONAIR CORPORATION\Home\Desktop\meteo\"
    
    'compte le nombre de pieces jointes
    NbAttachments = MonMail.Attachments.Count
    
    'contrôles possibles:nom de l'expéditeur, adresse mail expéditeur et sujet du mail
        'MonMail.SenderName= ""
        'MonMail.SenderEmailAddress
        'MonMail.Subject
        'If MonMail.Subject = "test" Then
        For j = 1 To numero
            i = 1
                Do While i <= NbAttachments
                strAttachment = MonMail.Attachments.Item(i).FileName
                MonMail.Attachments.Item(i).SaveAsFile chemin & strAttachment
                i = i + 1
                Loop
        Next j
        'End If
End Sub
Sub getEmailsSelected()
    
    Dim myOlSel As Outlook.Selection
    Dim myOlExp As Outlook.Explorer
    Dim gtStartDate As String
    Dim gtEndDate As String
    Dim ReceivedTime As Variant
    Dim ReceivedTimecorrige As Variant
    Dim NS As Object, dossier As Object
    Dim olapp As New Outlook.Application
    
    Set NS = olapp.GetNamespace("MAPI")
    Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")

    'gtStartDate = InputBox("Type the start date (format MM/DD/YYYY)")
    'gtEndDate = InputBox("Type the end date (format MM/DD/YYYY)")
    gtStartDate = "04/10/2021"
    gtEndDate = "04/10/2021"
    Set myOlExp = Application.ActiveExplorer
    
    For Each i In dossier.Items
        ReceivedTime = Left(i.ReceivedTime, 10) ' isole la date
        
        ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
        
        Set myOlSel = myOlExp.Selection("[Received] >= '" & gtStartDate & "' And [Received] <= '" & gtEndDate & "'")
    Next i

End Sub
Sub getEmailsSelecteda()

    Dim oFolder As Folder
    Dim oItems As Items
    Dim i As Long
    Dim j As Object
    Dim gtStartDate As String
    Dim gtEndDate As String
    Dim ReceivedTime As Variant
    Dim ReceivedTimecorrige As Variant
    Dim NS As Object, dossier As Object
    Dim olapp As New Outlook.Application
    Dim k As Integer
    
    Set NS = olapp.GetNamespace("MAPI")
    Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")

'    gtStartDate = InputBox("Type the start date (format MM/DD/YYYY)")
'    gtEndDate = InputBox("Type the end date (format MM/DD/YYYY)")
    gtStartDate = "04/10/2021"
    gtEndDate = "04/10/2021"
    Set oFolder = ActiveExplorer.CurrentFolder
    k = 0
    
    For Each j In dossier.Items
        ReceivedTime = Left(j.ReceivedTime, 10) ' isole la date
        
        ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
        
        
        If ReceivedTime = gtStartDate Then
            k = 1
            Set oItems = oFolder.Items.Restrict("[ReceivedTime] >= '" & gtStartDate & "' And [ReceivedTime] <= '" & gtEndDate & "'")
             For i = 1 To oItems.Count
                ActiveExplorer.AddToSelection oItems(i)
            Next i
            
        End If
        ActiveExplorer.ClearSelection
    Next j

    
End Sub
Sub getEmailsSelectedBB()

    Dim oFolder As Folder
    Dim oItems As Items
    Dim i As Long
    Dim j As Variant
    Dim gtStartDate As String
    Dim gtEndDate As String
    Dim cpt As Integer
    Dim NS As Object, dossier As Object
    Dim olapp As New Outlook.Application
    Dim k As Integer
    Dim MyDestFolder As Variant
        
    Set NS = olapp.GetNamespace("MAPI")
    Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception")
    Set MyDestFolder = dossier.Folders("Archives")
    cpt = 0
    k = 0
    'gtStartDate = InputBox("Type the start date (format MM/DD/YYYY)")
    'gtEndDate = InputBox("Type the end date (format MM/DD/YYYY)")
    gtStartDate = "06/10/2021"
    gtEndDate = "06/10/2021"
    
    Set oFolder = ActiveExplorer.CurrentFolder
    For Each j In dossier.Items
        cpt = cpt + 1
        ReceivedTime = Left(j.ReceivedTime, 10)
        'Set oItems = oFolder.Items.Restrict("[ReceivedTime] >= '" & gtStartDate & "' And [ReceivedTime] <= '" & gtEndDate & "'")
        If ReceivedTime = gtStartDate Then
            k = k + 1
            cpt = cpt - 1
            'ActiveExplorer.AddToSelection j
            j.Move MyDestFolder
            
        End If
    Next j

'j.Move MyDestFolder
fin:
End Sub
Sub DeleteMailCorbeille()
Dim olapp As New Outlook.Application
Dim MyDestFolder As Variant
Dim dossier As Object
Dim i As Variant
Dim myOlExp As Outlook.Explorer
Dim myOlSel As Outlook.Selection
 
 
Set NS = olapp.GetNamespace("MAPI")
Set dossier = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Corbeille")
'Set MyDestFolder = dossier.Folders("Corbeille")

For Each i In dossier.Items
    ActiveExplorer.AddToSelection i
Next i
Set myOlExp = Application.ActiveExplorer

Set myOlSel = myOlExp.Selection
For Each M In myOlSel
    M.Delete
Next M
End Sub
Sub MoveArchivestoCorbeille()
'prend tous les mails de archives vers la corbeille
Dim olapp As New Outlook.Application
Dim MyDestFolder As Variant
Dim dossierEmeteur As Object 'emeteur
Dim dossierDest As Object  ' corbeille
Dim i As Variant
Dim myOlExp As Outlook.Explorer
Dim myOlSel As Outlook.Selection
Dim M As Object
 
Set NS = olapp.GetNamespace("MAPI")
Set dossierEmeteur = NS.Folders("meteocaudry@gmail.com").Folders("Boîte de réception").Folders("Archives")
Set dossierDest = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Corbeille")

For Each i In dossierEmeteur.Items
    ActiveExplorer.AddToSelection i
Next i
Set myOlExp = Application.ActiveExplorer

Set myOlSel = myOlExp.Selection
For Each M In myOlSel
    M.Move dossierDest
Next M
End Sub


