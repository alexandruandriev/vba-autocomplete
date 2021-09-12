Attribute VB_Name = "safety_checks"
Function isEmpty(InputText As Variant) As Boolean
    If Trim(InputText & vbNullString) = vbNullString Then
        isEmpty = True
    Else
        isEmpty = False
    End If
    
End Function

Function isEmpty_list(coll As Object) As Boolean
    Dim item As Variant
    For Each item In coll
        If isEmpty(item) Then
            isEmpty_list = True
            Exit For
        Else
            isEmpty_list = False
        End If
    Next item
    
End Function

