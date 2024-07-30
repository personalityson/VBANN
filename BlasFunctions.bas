Attribute VB_Name = "BlasFunctions"
Option Explicit

Private Const OPENBLAS_DIRECTORY As String = "C:\Users\hello\OneDrive\Documents\VBANN"
Private Const OPENBLAS_LIBRARY_NAME As String = "libopenblas.dll"

Private m_bIsBlasAvailable As Boolean

Public Declare PtrSafe Sub daxpby Lib "libopenblas.dll" (ByRef n As Long, _
                                                         ByRef alpha As Double, _
                                                         ByVal X As LongPtr, _
                                                         ByRef incX As Long, _
                                                         ByRef beta As Double, _
                                                         ByVal Y As LongPtr, _
                                                         ByRef incY As Long)

Public Declare PtrSafe Sub dgemv Lib "libopenblas.dll" (ByVal trans As String, _
                                                        ByRef m As Long, _
                                                        ByRef n As Long, _
                                                        ByRef alpha As Double, _
                                                        ByVal A As LongPtr, _
                                                        ByRef ldA As Long, _
                                                        ByVal X As LongPtr, _
                                                        ByRef incX As Long, _
                                                        ByRef beta As Double, _
                                                        ByVal Y As LongPtr, _
                                                        ByRef incY As Long)

Public Declare PtrSafe Sub dgemm Lib "libopenblas.dll" (ByVal transA As String, _
                                                        ByVal transB As String, _
                                                        ByRef m As Long, _
                                                        ByRef n As Long, _
                                                        ByRef k As Long, _
                                                        ByRef alpha As Double, _
                                                        ByVal A As LongPtr, _
                                                        ByRef ldA As Long, _
                                                        ByVal B As LongPtr, _
                                                        ByRef ldB As Long, _
                                                        ByRef beta As Double, _
                                                        ByVal C As LongPtr, _
                                                        ByRef ldC As Long)

Public Property Get IsBlasAvailable() As Boolean
    IsBlasAvailable = m_bIsBlasAvailable
End Property

Public Sub VerifyOpenBlasLibrary()
    Dim sLibraryPath As String
    
    sLibraryPath = Fso.BuildPath(OPENBLAS_DIRECTORY, OPENBLAS_LIBRARY_NAME)
    If Fso.FileExists(sLibraryPath) Then
        ChDir OPENBLAS_DIRECTORY
        m_bIsBlasAvailable = True
    End If
End Sub
