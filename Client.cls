VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public first_name As String
Public last_name As String
Public client_code As String
Public isCompany As Boolean

'_c means it is a constructor local variable
Public Sub InitializeWithValues(client_code_c As String, first_name_c As String, Optional last_name_c As String, Optional isCompany_c As Boolean = False)
    client_code = client_code_c
    'Capitalezes first letter of the name
    first_name = StrConv(first_name_c, vbProperCase)
    last_name = StrConv(last_name_c, vbProperCase)
    isCompany = isCompany_c
    
End Sub

Function getName() As String
    Dim fullName As String
    If isCompany Then
        fullName = "Stimate client"
    Else
        fullName = last_name + " " + first_name
    End If
    getName = fullName
End Function
