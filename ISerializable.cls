VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISerializable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

#If USE_LATE_BINDING_IN_INTERFACES Then
    Public Sub Save(ByVal oWorksheet As Object)
        '
    End Sub
    
    Public Sub Load(ByVal oWorksheet As Object)
        '
    End Sub
#Else
    Public Sub Save(ByVal oWorksheet As Serializer)
        '
    End Sub
    
    Public Sub Load(ByVal oWorksheet As Serializer)
        '
    End Sub
#End If



