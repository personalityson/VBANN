VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sequential"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILayer
Implements ISerializable

#Const USE_LATE_BINDING_IN_INTERFACES = True

'State
Private m_oLayers As Collection
Private m_vLearnables As Variant

Public Property Get Learnables() As Variant
    Learnables = m_vLearnables
End Property

Public Sub Init()
    Clear
End Sub

Public Sub Add(ByVal oLayer As ILayer)
    Const PROCEDURE_NAME As String = "Sequential.Add"

    If oLayer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid ILayer object is required."
    End If
    m_oLayers.Add oLayer
    m_vLearnables = Combine(m_vLearnables, oLayer.Learnables)
End Sub

Public Function Forward(ByVal X As Tensor, _
                        Optional ByVal bIsTraining As Boolean) As Tensor
    Const PROCEDURE_NAME As String = "Sequential.Forward"
    Dim i As Long

    If X Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    For i = 1 To m_oLayers.Count
        Set X = m_oLayers(i).Forward(X, bIsTraining)
    Next i
    Set Forward = X
End Function

Public Function Backward(ByVal dLdY As Tensor) As Tensor
    Const PROCEDURE_NAME As String = "Sequential.Backward"
    Dim i As Long

    If dLdY Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Tensor object is required."
    End If
    For i = m_oLayers.Count To 1 Step -1
        Set dLdY = m_oLayers(i).Backward(dLdY)
    Next i
    Set Backward = dLdY
End Function

Private Sub Clear()
    Set m_oLayers = New Collection
    m_vLearnables = Array()
End Sub

Private Sub Save(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Save"
    Dim i As Long

    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    With oSerializer
        .WriteLong m_oLayers.Count
        For i = 1 To m_oLayers.Count
            .WriteObject m_oLayers(i)
        Next i
    End With
End Sub

Private Sub Load(ByVal oSerializer As Serializer)
    Const PROCEDURE_NAME As String = "Sequential.Load"
    Dim i As Long
    Dim lNumLayers As Long
    
    If oSerializer Is Nothing Then
        Err.Raise 5, PROCEDURE_NAME, "Valid Serializer object is required."
    End If
    Clear
    With oSerializer
        lNumLayers = .ReadLong()
        For i = 1 To lNumLayers
            Me.Add .ReadObject()
        Next i
    End With
End Sub

#If USE_LATE_BINDING_IN_INTERFACES Then
    Private Property Get ILayer_Learnables() As Variant
        ILayer_Learnables = Me.Learnables
    End Property
    
    Private Function ILayer_Forward(ByVal X As Object, _
                                    Optional ByVal bIsTraining As Boolean) As Object
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Object) As Object
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Object)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Object)
        Load oSerializer
    End Sub
#Else
    Private Property Get ILayer_Learnables() As Variant
        ILayer_Learnables = Me.Learnables
    End Property
    
    Private Function ILayer_Forward(ByVal X As Tensor, _
                                    Optional ByVal bIsTraining As Boolean) As Tensor
        Set ILayer_Forward = Me.Forward(X, bIsTraining)
    End Function
    
    Private Function ILayer_Backward(ByVal dLdY As Tensor) As Tensor
        Set ILayer_Backward = Me.Backward(dLdY)
    End Function
    
    Private Sub ISerializable_Save(ByVal oSerializer As Serializer)
        Save oSerializer
    End Sub
    
    Private Sub ISerializable_Load(ByVal oSerializer As Serializer)
        Load oSerializer
    End Sub
#End If

