VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fila"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Columnas"
Option Explicit

Public Key As String

Private mvarColumnas As Columnas


Public Property Get Columnas() As Columnas
Attribute Columnas.VB_UserMemId = 0
    If mvarColumnas Is Nothing Then
        Set mvarColumnas = New Columnas
    End If


    Set Columnas = mvarColumnas
End Property


Public Property Set Columnas(vData As Columnas)
    Set mvarColumnas = vData
End Property
Private Sub Class_Terminate()
    Set mvarColumnas = Nothing
End Sub



