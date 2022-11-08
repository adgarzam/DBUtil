VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Filas"
Option Explicit

Public Key As String

Private mvarFilas As Filas
'variables locales para almacenar los valores de las propiedades
Private mvarNombre As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarEstado As Integer 'copia local
Public Property Let Estado(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Estado = 5
    mvarEstado = vData
End Property


Public Property Get Estado() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Estado
    Estado = mvarEstado
End Property




Public Property Let Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property

Public Property Get Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property





Public Property Get Filas() As Filas
Attribute Filas.VB_UserMemId = 0
    If mvarFilas Is Nothing Then
        Set mvarFilas = New Filas
    End If


    Set Filas = mvarFilas
End Property


Public Property Set Filas(vData As Filas)
    Set mvarFilas = vData
End Property
Private Sub Class_Terminate()
    Set mvarFilas = Nothing
End Sub



