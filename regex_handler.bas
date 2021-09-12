Attribute VB_Name = "regex_handler"
'Testing Functions
'---------------------------
Function test_regex(text_pattern As String, Text As String) As Boolean

Dim objRegex As RegExp
Set objRegex = New RegExp
objRegex.pattern = text_pattern
test_regex = objRegex.Test(Text)

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

Function test_furnizor(denumire As String) As Boolean

test_furnizor = test_regex("Enel\sEnergie\sS.A[\s\w\\]+.A", denumire)

End Function



'End of Testing Functions
'---------------------------

'Get Functions
'---------------------------
Function get_client_code_pattern() As String

    get_client_code_pattern = "C[\d]{8}(?!\d)(?!\w)"
    
End Function

Function get_doamna_pattern() As String

    get_doamna_pattern = "Doamna[\s\w]+"
    
End Function

Function get_domnule_pattern() As String

    get_domnule_pattern = "Domnule[\s\w]+"
    
End Function

Function get_furnizor_pattern() As String

    get_furnizor_pattern = "Enel\sEnergie\sS.A[\s\w\\]+.A"
    
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

Sub change_matches_to(matches As MatchCollection, ToChange As String)
'Changes the matches you pass to anything you want

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

