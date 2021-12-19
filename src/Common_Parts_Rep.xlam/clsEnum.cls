VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private dicAccdbDataType As Dictionary
Private Sub Class_Initialize()
    Call Initial_Valueset
End Sub
Private Sub Initial_Valueset()
    'dicAccDataType
    If dicAccdbDataType Is Nothing Then
        Set dicAccdbDataType = New Dictionary
    End If
'Public Enum ACCDB_Data_Type
'    [Text] = 1
'    [INTEGER] = 2
'    [BIT] = 3
'    [Boolean] = 3
'    [COUNTER] = 4
'    [AUTOINCREMENT] = 4
'    [Decimal] = 5
'End Enum
    dicAccdbDataType.Add ACCDB_Data_Type.Text, "TEXT"
    dicAccdbDataType.Add ACCDB_Data_Type.INTEGER, "INTEGER"
    dicAccdbDataType.Add ACCDB_Data_Type.BIT, "BIT"
    dicAccdbDataType.Add ACCDB_Data_Type.COUNTER, "COUNTER"
    dicAccdbDataType.Add ACCDB_Data_Type.Decimal, "DECIMAL"
End Sub
Property Get AccdbDataType(accdbType As ACCDB_Data_Type) As String
    AccdbDataType = dicAccdbDataType(accdbType)
    Exit Property
End Property
