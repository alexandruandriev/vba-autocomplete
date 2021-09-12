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

'_c means it is a constructor local variable
Public Sub InitializeWithValues(client_code_c As String, first_name_c As String, Optional last_name_c As String)
    client_code = client_code_c
    'Capitalezes first letter of the name
    first_name = StrConv(first_name_c, vbProperCase)
    last_name = StrConv(last_name_c, vbProperCase)
    
End Sub

Function checkClientCode() As String

End Function
