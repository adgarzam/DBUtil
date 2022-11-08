VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Filas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Fila"
Attribute VB_Ext_KEY = "Member0" ,"Fila"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private lngFila As Long
'variables locales para almacenar los valores de las propiedades
Private mvarBOF As Boolean 'copia local
Private mvarEOF As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades

Public Property Get FilaActual() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FilaActual
    FilaActual = lngFila
End Property

Public Property Let EOF(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.EOF = 5
    mvarEOF = vData
End Property

Public Property Get EOF() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.EOF
    EOF = mvarEOF
End Property

Public Property Let BOF(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.BOF = 5
    mvarBOF = vData
End Property

Public Property Get BOF() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.BOF
    BOF = mvarBOF
End Property

Public Sub MoverHacia(Fila As Long)
  If Fila >= (0) And Fila < mCol.Count Then
    lngFila = Fila
  End If
End Sub

Public Sub MoverUlt()
  lngFila = mCol.Count - 1
End Sub

Public Sub MoverPrim()
  lngFila = 0
End Sub

Public Sub MoverAnt()
  If lngFila > 0 Then
    lngFila = lngFila - 1
    Me.BOF = False
  Else
    lngFila = -1
    Me.BOF = True
  End If
End Sub

Public Sub MoverSig()
  If lngFila < mCol.Count Then
    lngFila = lngFila + 1
    Me.EOF = False
  Else
    lngFila = mCol.Count
    Me.EOF = True
  End If
End Sub

Public Sub Clear()
  ' Borra todas las filas
  Do While mCol.Count > 0
    mCol.Remove 0
    lngFila = -1
    Me.BOF = True
    Me.EOF = True
  Loop
End Sub

Public Sub Buscar(Criterio As String, Optional SaltarRegistros As Long, Optional DireccionBusqueda As ADODB.SearchDirectionEnum = adSearchForward, Optional Inicio As Variant)
  ' Sin establecer
End Sub

Public Function Add(Optional Clave As String) As Fila
    'crear un nuevo objeto
    Dim objNewMember As Fila
    Set objNewMember = New Fila


    'establecer las propiedades que se transfieren al método
    objNewMember.Key = Clave
    If Clave = "" Then
      mCol.Add objNewMember
    Else
      mCol.Add objNewMember, Clave
    End If

    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Fila
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
  On Error GoTo 0
End Property

Public Function ObtenerFila(Optional CualFila As Variant = "") As Fila
  On Error Resume Next
  If Not IsNull(CualFila) Then
    Set ObtenerFila = mCol(CualFila)
  Else
    Set ObtenerFila = mCol(lngFila)
  End If
  On Error GoTo 0
End Function

Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)

    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub
