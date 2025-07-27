VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const USE_LATE_BINDING_IN_INTERFACES = True

#If USE_LATE_BINDING_IN_INTERFACES Then
    Public Property Get LearningRate() As Double
        '
    End Property

    Public Property Let LearningRate(ByVal dblLearningRate As Double)
        '
    End Property
    
    Public Sub Update(ByVal vParameters As Variant)
        '
    End Sub
#Else
    Public Property Get LearningRate() As Double
        '
    End Property

    Public Property Let LearningRate(ByVal dblLearningRate As Double)
        '
    End Property

    Public Sub Update(ByVal W As Parameter)
        '
    End Sub
#End If
