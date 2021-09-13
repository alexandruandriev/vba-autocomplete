VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Word Autocomplete Alpha 1.0"
   ClientHeight    =   9180.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12480
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ResetInformatiiClient_textbox()
    Me.nume_textbox.Enabled = True
    Me.nume_label.Enabled = True
    Me.prenume_label.Enabled = True
    Me.prenume_textbox.Enabled = True
End Sub

Private Sub DisableNumePrenumeClient_textbox()
    Me.nume_textbox.Enabled = False
    Me.nume_label.Enabled = False
    Me.prenume_label.Enabled = False
    Me.prenume_textbox.Enabled = False
End Sub

Private Sub InformatiiClient_frame_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub codclient_label_Click()

End Sub

Private Sub doamna_box_Click()

ResetInformatiiClient_textbox

End Sub

Private Sub domnul_box_Click()

ResetInformatiiClient_textbox

End Sub

Private Sub email_option_Click()

End Sub

Private Sub firma_box_Click()

DisableNumePrenumeClient_textbox

End Sub

Private Sub furnizor_ComboBox_Change()

End Sub

Private Sub icaz_btn_Click()
Dim matches2 As MatchCollection
Set matches2 = regex_handler.find_given_pattern("C[\d]{8}")
Dim fnd As Match

   With Selection
        .HomeKey wdStory
        With .Find
            .ClearFormatting
            .Forward = True
            .Format = False
            .MatchCase = True
            For Each fnd In matches2
                Debug.Print (fnd)
                With Selection
                    .MoveRight wdCharacter
                End With
            Next fnd
        End With
        .HomeKey wdStory
    End With
End Sub

Private Sub iclient_btn_Click()

Dim myClient As New Client


'Stores all the answers in a list to check if they are empty
'Dim coll As Object
'Set coll = CreateObject("System.Collections.ArrayList")
'coll.Add Trim(Me.codclient_textbox.Value)
'coll.Add Trim(Me.nume_textbox.Value)
'coll.Add Trim(Me.prenume_textbox.Value)


'If safety_checks.isEmpty_list(coll) Then
'    MsgBox "Toate spatiile trebuie completate!", vbCritical, "Eroare 101"
'    Exit Sub
'End If



If Me.domnul_box Then

    myClient.InitializeWithValues Trim(Me.codclient_textbox.Value), Trim(Me.prenume_textbox), Trim(Me.nume_textbox)
    'If the client code was incorect, exit the sub
    If document_changes.change_client_code(myClient.client_code) = False Then
        Exit Sub
    End If
    document_changes.change_name_domnul myClient.first_name, myClient.last_name
End If
   
    


If Me.doamna_box Then

    myClient.InitializeWithValues Trim(Me.codclient_textbox.Value), Trim(Me.prenume_textbox), Trim(Me.nume_textbox)
    'If the client code was incorect, exit the sub
    If document_changes.change_client_code(myClient.client_code) = False Then
        Exit Sub
    End If
    document_changes.change_name_doamna myClient.first_name, myClient.last_name
End If

    
    
    
    
    

If Me.firma_box Then

    myClient.InitializeWithValues Trim(Me.codclient_textbox.Value), Trim(Me.prenume_textbox), , True
    'If the client code was incorect, exit the sub
    If document_changes.change_client_code(myClient.client_code) = False Then
        Exit Sub
    End If
    document_changes.change_name_firma (myClient.getName())
End If
    
    
    







End Sub

Private Sub iclient_btn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub ifurnizor_btn_Click()
Dim nume_furnizor As String
nume_furnizor = Me.furnizor_ComboBox.Value
If Me.furnizor_ComboBox.ListIndex = 2 Then
    nume_furnizor = "Enel Energie S.A"
End If

document_changes.change_furnizor_name (nume_furnizor)


End Sub

Private Sub sex_label_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

UserForm1.furnizor_ComboBox.AddItem ("Enel Energie S.A")
UserForm1.furnizor_ComboBox.AddItem ("Enel Energie Muntenia S.A")
UserForm1.furnizor_ComboBox.AddItem ("Zona New Enel")
UserForm1.codclient_textbox.ControlTipText = "Codul de client pe care vrei sa il schimbi."

UserForm1.Show
End Sub
