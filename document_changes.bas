Attribute VB_Name = "document_changes"
Function change_client_code(client_code As String)
'Performs a check if the input is correct
Dim isCorrect As Boolean
Debug.Print ("Client code is : " + client_code)
isCorrect = regex_handler.test_client_code(client_code)


If isCorrect = True Then
    Dim matches As MatchCollection
    
    Set matches = regex_handler.find_given_pattern(regex_handler.get_client_code_pattern)
    regex_handler.change_matches_to matches, client_code
    
Else

MsgBox "Codul de client este invalid!"

End If


End Function

Function change_name_domnul(nume As String, prenume As String)
Dim isCorrect As Boolean
Dim completeName As String
completeName = "Domnule " + nume + " " + prenume
isCorrect = regex_handler.test_domnule(completeName)

If isCorrect Then
    Dim matches As MatchCollection
    Set matches = regex_handler.find_given_pattern(regex_handler.get_domnule_pattern)
    If matches.Count = 0 Then
        Set matches = regex_handler.find_given_pattern(regex_handler.get_doamna_pattern)
    End If
    
    regex_handler.change_matches_to matches, completeName
Else
    MsgBox "Numele nu este valid!"
End If

End Function

Function change_name_doamna(nume As String, prenume As String)
Dim isCorrect As Boolean
Dim completeName As String
completeName = "Doamna " + nume + " " + prenume
isCorrect = regex_handler.test_doamna(completeName)

If isCorrect Then
    Dim matches As MatchCollection
    Set matches = regex_handler.find_given_pattern(regex_handler.get_doamna_pattern)
    If matches.Count = 0 Then
        Set matches = regex_handler.find_given_pattern(regex_handler.get_domnule_pattern)
    End If
    
    regex_handler.change_matches_to matches, completeName
Else
    MsgBox "Numele nu este valid!"
End If

End Function


