VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Columna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Tabla"
Option Explicit

Public Key As String

'variables locales para almacenar los valores de las propiedades
Private mvarValor As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarTamanio As Long 'copia local
Private mvarAtributos As Long 'copia local
Private mvarFormatoDatos As Long 'copia local
Private mvarTamanioDefinido As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarEscalaNumerica As Byte 'copia local
Private mvarValorOriginal As Variant 'copia local
Private mvarPrecision As Byte 'copia local
Private mvarEstado As Long 'copia local
Private mvarTipo As DataTypeEnum 'copia local
Private mvarValorSubyacente As Variant 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarTabla As String 'copia local
Public Property Let Tabla(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tabla = 5
    mvarTabla = vData
End Property


Public Property Get Tabla() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    Tabla = mvarTabla
End Property




Public Property Let ValorSubyacente(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorSubyacente = 5
    mvarValorSubyacente = vData
End Property


Public Property Set ValorSubyacente(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.ValorSubyacente = Form1
    Set mvarValorSubyacente = vData
End Property


Public Property Get ValorSubyacente() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorSubyacente
    If IsObject(mvarValorSubyacente) Then
        Set ValorSubyacente = mvarValorSubyacente
    Else
        ValorSubyacente = mvarValorSubyacente
    End If
End Property



Public Property Let Tipo(ByVal vData As ADODB.DataTypeEnum)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Tipo = Form1
    mvarTipo = vData
End Property


Public Property Get Tipo() As ADODB.DataTypeEnum
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tipo
    Tipo = mvarTipo
End Property



Public Property Let Estado(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Estado = 5
    mvarEstado = vData
End Property


Public Property Get Estado() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Estado
    Estado = mvarEstado
End Property

Public Property Let Precision(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Precision = 5
    mvarPrecision = vData
End Property


Public Property Get Precision() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Precision
    Precision = mvarPrecision
End Property



Public Property Let ValorOriginal(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ValorOriginal = 5
    mvarValorOriginal = vData
End Property


Public Property Set ValorOriginal(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.ValorOriginal = Form1
    Set mvarValorOriginal = vData
End Property


Public Property Get ValorOriginal() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ValorOriginal
    If IsObject(mvarValorOriginal) Then
        Set ValorOriginal = mvarValorOriginal
    Else
        ValorOriginal = mvarValorOriginal
    End If
End Property



Public Property Let EscalaNumerica(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.EscalaNumerica = 5
    mvarEscalaNumerica = vData
End Property


Public Property Get EscalaNumerica() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.EscalaNumerica
    EscalaNumerica = mvarEscalaNumerica
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



Public Property Let TamanioDefinido(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.TamanioDefinido = 5
    mvarTamanioDefinido = vData
End Property


Public Property Get TamanioDefinido() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.TamanioDefinido
    TamanioDefinido = mvarTamanioDefinido
End Property



Public Property Let FormatoDatos(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FormatoDatos = 5
    mvarFormatoDatos = vData
End Property


Public Property Get FormatoDatos() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FormatoDatos
    FormatoDatos = mvarFormatoDatos
End Property



Public Property Let Atributos(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Atributos = 5
    mvarAtributos = vData
End Property


Public Property Get Atributos() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Atributos
    Atributos = mvarAtributos
End Property



Public Property Let Tamanio(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tamanio = 5
    mvarTamanio = vData
End Property


Public Property Get Tamanio() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tamanio
    Tamanio = mvarTamanio
End Property



Public Property Let Valor(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Valor = 5
    mvarValor = vData
    If (Me.Estado And ModoEstadoEnum.Modificar) = 0 Then
      Me.Estado = Me.Estado + ModoEstadoEnum.Modificar
    End If
End Property


Public Property Set Valor(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.Valor = Form1
    Set mvarValor = vData
    If (Me.Estado And ModoEstadoEnum.Modificar) = 0 Then
      Me.Estado = Me.Estado + ModoEstadoEnum.Modificar
    End If
End Property


Public Property Get Valor() As Variant
Attribute Valor.VB_UserMemId = 0
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Valor
    If IsObject(mvarValor) Then
        Set Valor = mvarValor
    Else
        Valor = mvarValor
    End If
End Property



