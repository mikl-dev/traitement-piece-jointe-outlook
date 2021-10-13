VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USFSupp 
   Caption         =   "Corbeille"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3390
   OleObjectBlob   =   "USFSupp.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USFSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecupDate As Variant
Private Sub CBSupp_Click()

'detecte si une date est rentrée et lance le prog LireMessage
'sinon lance le prog tridates pour toutes les dates

If TBDateSupp <> "" Then
    RecupDate = TBDateSupp.Value
    '*** Appel du programme de sauvegarde des photos de la date saisie
    Call SuppDateDefinitivement(RecupDate)
Else
    '*** Appel du programme de tri de toutes les dates presentes dans la boite de reception
    Call SuppToutDefinitivement
End If
End Sub

Private Sub CBSuppAnnule_Click()
    Unload Me
End Sub
Sub SuppDateDefinitivement(RecupDate)
Dim MonApp As Outlook.Application
Dim MonNameSpace As Outlook.NameSpace
Dim Mondossier As Outlook.Folder
Dim MonMail As Outlook.MailItem
Dim i As Object
Dim ReceivedTime As Variant
Dim ReceivedTimecorrige As Variant
Dim MonDossiercorbeille As Outlook.Folder
Dim MonDossierenvoye As Outlook.Folder
'Instance des objets
Set MonApp = Outlook.Application
Set MonNameSpace = MonApp.GetNamespace("MAPI")
Set Mondossier = MonNameSpace.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Archives")
Set MonDossiercorbeille = MonNameSpace.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Corbeille")
Set MonDossierenvoye = MonNameSpace.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Messages envoyés")

For Each i In Mondossier.Items
    ReceivedTime = Left(i.ReceivedTime, 10) ' isole la date
    ReceivedTimecorrige = Replace(ReceivedTime, "/", "")
    If ReceivedTime = RecupDate Then    ' compare la date du mail et la date du jour où est lancé la macro
        i.Delete     ' supprimer le mail
    End If
Next i
For Each i In MonDossierenvoye.Items
    If ReceivedTime = RecupDate Then    ' compare la date du mail et la date du jour où est lancé la macro
        i.Delete     ' supprimer le mail
    End If
Next i
For Each i In MonDossiercorbeille.Items
    If ReceivedTime = RecupDate Then    ' compare la date du mail et la date du jour où est lancé la macro
        i.Delete     ' supprimer le mail
    End If
Next i
End Sub
Sub SuppToutDefinitivement()
Dim oFolder As Folder
Dim oItems As Items
Dim NS As Object, Mondossier As Object, MonDossiercorbeille As Object, MonDossierenvoye As Object
Dim olapp As New Outlook.Application
Dim MyDestFolder As Variant
Dim i As Object
Dim myOlExp As Outlook.explorer
Dim myOlSel As Outlook.Selection

Set NS = olapp.GetNamespace("MAPI")
Set Mondossier = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Archives")
Set MonDossiercorbeille = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Corbeille")
Set MonDossierenvoye = NS.Folders("meteocaudry@gmail.com").Folders("[Gmail]").Folders("Messages envoyés")

For Each i In Mondossier.Items
    i.Delete     ' supprimer le mail
Next i
For Each i In MonDossierenvoye.Items
    i.Delete     ' supprimer le mail
Next i
For Each i In MonDossiercorbeille.Items
    i.Delete     ' supprimer le mail
Next i
Unload Me
End Sub
