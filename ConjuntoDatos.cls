VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConjuntoDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Tablas"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarTablas As Tablas

Public Enum ModoEstadoEnum
  SinCambios = 0
  Insertar = 1
  Eliminar = 2
  Modificar = 4
End Enum

Public Property Get Tablas() As Tablas
    If mvarTablas Is Nothing Then
        Set mvarTablas = New Tablas
    End If

    Set Tablas = mvarTablas
End Property

Public Property Set Tablas(vData As Tablas)
    Set mvarTablas = vData
End Property
Private Sub Class_Terminate()
    Set mvarTablas = Nothing
End Sub

Public Function Llenar(Rs As ADODB.Recordset, Optional NombreTabla As String = "") As Tabla
  ' Var
  Dim tbl As Tabla
  Dim f As Fila
  Dim c As Columna
  Dim fld As ADODB.Field

  '
  If Not Rs.BOF And Not Rs.EOF Then
    Set tbl = Me.Tablas.Add(NombreTabla, ModoEstadoEnum.SinCambios)
    Rs.MoveFirst
    Do While Not Rs.EOF
      Set f = tbl.Filas.Add
      For Each fld In Rs.Fields
        Set c = f.Columnas.Add(fld.Name, fld.Value)
        c.Atributos = fld.Attributes
        c.EscalaNumerica = fld.NumericScale
        c.EscalaNumerica = fld.Status
        ' c.FormatoDatos = fld.DataFormat
        c.Precision = fld.Precision
        c.Tamanio = fld.ActualSize
        c.TamanioDefinido = fld.DefinedSize
        c.Tipo = fld.Type
        On Error Resume Next
        c.ValorOriginal = fld.OriginalValue
        c.ValorSubyacente = fld.UnderlyingValue
        c.Estado = ModoEstadoEnum.SinCambios
        On Error GoTo 0
        c.Tabla = Rs.Source
      Next fld
      Rs.MoveNext
    Loop
    
  End If

  Set Llenar = tbl
End Function

