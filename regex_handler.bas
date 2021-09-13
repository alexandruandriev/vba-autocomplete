Attribute VB_Name = "regex_handler"
'Testing Functions
'---------------------------
Function test_regex(text_pattern As String, Text As String) As Boolean

Dim objRegex As RegExp
Set objRegex = New RegExp
objRegex.pattern = text_pattern
test_regex = objRegex.Test(Text)

End Function

Function test_company(name As String) As Boolean

test_company = test_regex("Stimate\s+client(?!\w)", name)

End Function

Function test_client_code(client_code As String) As Boolean

test_client_code = test_regex("C[\d]{8}(?!\d)(?!\w)", client_code)

End Function

Function test_doamna(name As String) As Boolean

test_doamna = test_regex("Doamna[\s\w]+", name)

End Function

Function test_domnule(name As String) As Boolean

test_domnule = test_regex("Domnule[\s\w]+", name)

End Function

Function test_furnizor(name As String) As Boolean
test_furnizor = test_regex = "Enel\s+Energie\s+S.A\s+\\ Enel\s+Energie\s+Muntenia\s+S.A"
End Function




'End of Testing Functions
'---------------------------

'Get Functions
'---------------------------
Function get_company_pattern() As String
    get_company_pattern = "Stimate\s+client(?!\w)"
End Function


Function get_client_code_pattern() As String

    get_client_code_pattern = "C[\d]{8}(?!\d)(?!\w)"
    
End Function

Function get_doamna_pattern() As String

    get_doamna_pattern = "Stimata\s+Doamna[\s\w]+"
    
End Function

Function get_domnule_pattern() As String

    get_domnule_pattern = "Stimate\s+Domnule[\s\w]+"
    
End Function

Function get_furnizor_pattern() As String

    get_furnizor_pattern = "Enel\s+Energie\s+S.A\s+\\ Enel\s+Energie\s+Muntenia\s+S.A"
    
End Function

Function get_ee_pattern() As String

    get_ee_pattern = "Enel\s+Energie\s+S.A"
    
End Function

Function get_em_pattern() As String

    get_em_pattern = "Enel\s+Energie\s+Muntenia\s+S.A"
    
End Function
'End of get Functions
'---------------------------

Function find_given_pattern(pattern As String) As MatchCollection
'Returns all matches from a given Regex pattern
Dim objRegex As RegExp
Dim matches As MatchCollection
Dim fnd As Match

    With Selection
        .HomeKey wdStory
        .WholeStory
    End With
    
    Set objRegex = New RegExp

    With objRegex
        .pattern = pattern
        .Global = True
        .IgnoreCase = True
        Set matches = .Execute(Selection.Text)
    End With
    For Each fnd In matches
        Debug.Print ("Fnd in find_given_matches was : " + fnd)
        Next fnd
    Set find_given_pattern = matches
End Function

'Changes the matches you pass to anything you want
Sub change_matches_to(matches As MatchCollection, ToChange As String)


Dim fnd As Match

'Selects whole document
With Selection
        .HomeKey wdStory
        .WholeStory
End With
    
With Selection
        .HomeKey wdStory
        With .Find
            .ClearFormatting
            .Forward = True
            .Format = False
            .MatchCase = True
            For Each fnd In matches
                Debug.Print ("Fnd was2: " + fnd)
                Debug.Print ("To change is : " + ToChange)
                .Text = fnd
                .Execute
                With Selection
                    .Text = ToChange
                    .MoveRight wdCharacter
                End With
            Next fnd
        End With
        .HomeKey wdStory
End With
End Sub

