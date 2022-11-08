VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Enum TipoDeValorPorAnalizarEnum
  Numerico
  Cadena
  Fecha
  Condicional
End Enum

Public Enum TipoDeConsultaDeEjecucionEnum
  Insercion
  Actualizacion
  Borrado
End Enum

Public Enum TipoTabla
  cpNinguna
  cpClavePrincipal
  cpCola
  cpSecuencia
End Enum

Private Type GUID
  data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

' Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID)
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Private WithEvents cnmCrono As Cronometro
Private cnx As ADODB.Connection
Private strTblClaves As String
Private strTblColas As String
Private strTblSecuencias As String
Private mvarUsuario As String 'copia local
Private mvarContrasenia As String 'copia local
Private mvarCadenaConexion As String 'copía local

Private Const COMPAS_DE_ESPERA As Integer = 200
'variables locales para almacenar los valores de las propiedades
Private mvarTag As String 'copia local
Public Property Let Tag(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tag = 5
    mvarTag = vData
End Property

Public Property Get Tag() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property

Private Function AbreCnx(Optional ForzarApertura As Boolean = False) As Boolean
  ' Var
  Dim bolCorrecto As Boolean
  
  If Len(cnx.ConnectionString) Then
    If (cnx.State = adStateClosed) Or (ForzarApertura) Then
      On Error Resume Next
      cnx.Close
      Err.Clear
      cnx.Open UserID:=mvarUsuario, Password:=mvarContrasenia
      bolCorrecto = (Err.Number = 0)
      Err.Clear
      On Error GoTo 0
    Else
      bolCorrecto = True
    End If
  Else
    bolCorrecto = False
  End If
  AbreCnx = bolCorrecto
End Function

Public Function AsignaDatosAFormulario(CualFrm As Object, tbl As Tabla) As Boolean
  ' Var
  Dim ctl As Control
  Dim strCampo As String
  Dim strTControl As String
  Dim strConvalidar As String
  Dim i As Integer
  Dim cProp As New PropVirtuales
  Dim X As Integer
  Dim cUt As New Utilerias
  Dim strCad As String
  
  For Each ctl In CualFrm.Controls
    strCampo = cProp.getPropiedad(ctl.Tag, "Campo")
    strConvalidar = cProp.getPropiedad(ctl.Tag, "Convalidar")
    If strCampo <> "" Then
      X = X + 1
      If Not (cProp.getPropiedad(ctl.Tag, "Asignar") = "No") Then
        strTControl = UCase(TypeName(ctl))
        Select Case strTControl
          Case "TEXTBOX"
            ctl.Text = "" & tbl(1)(strCampo).Valor
          Case "COMBOBOX", "LISTBOX"
            If cProp.getPropiedad(ctl.Tag, "Propiedad") = "Text" Then
              i = cUt.BuscaEnCombo(ctl, tbl(1)(strCampo).Valor, True)
            Else  ' De forma predeterminada, busca en ItemData
              i = cUt.BuscaEnCombo(ctl, tbl(1)(strCampo).Valor)
            End If
            
            If i > (-1) Then
              ctl.ListIndex = i
            End If
              
          Case "DTPICKER"
            ctl.Value = tbl(1)(strCampo).Valor
          Case "IMAGECOMBO"
            If ctl.ComboItems.Count > 0 Then
              strCad = Left(ctl.ComboItems(1).Key, 1)
              strCad = strCad & tbl(1)(strCampo).Valor
              Set ctl.SelectedItem = ctl.ComboItems(strCad)
              ctl.Text = ctl.SelectedItem.Text
            End If
          Case "CHECKBOX"
            ctl.Value = IIf(CInt(tbl(1)(strCampo).Valor) = 0, vbUnchecked, vbChecked)
          Case "LABEL"
            ctl.Caption = "" & tbl(1)(strCampo).Valor
        End Select
      End If
      
      If UCase(strConvalidar) = "SI" Or UCase(strConvalidar) = "SÍ" Then
        CualFrm.Convalida strCampo
      End If
      
    End If
  Next ctl
  
  AsignaDatosAFormulario = (X > 0)
End Function

Public Function BuscaFila(CualTabla As Tabla, QueBusca As Variant, BasadoEnQueCampo As Variant) As Fila
  ' Var
  Dim fil As Fila
  
  On Error Resume Next
  For Each fil In CualTabla.Filas
    If fil(BasadoEnQueCampo).Valor = QueBusca Then
      Exit For
    End If
  Next fil
  On Error GoTo 0
  
  If fil(BasadoEnQueCampo).Valor <> QueBusca Then
    Set fil = Nothing
  End If
  
  Set BuscaFila = fil
End Function

Private Function BuscaTabla(Tabla As String, Optional Crear As TipoTabla = cpNinguna) As Boolean
  ' Var
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim bolEncontrada As Boolean
  
  If Len(mvarCadenaConexion) > 0 Then
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
    
    ' Obtengo la tabla que me piden
    On Error Resume Next
    Set tbl = cat.Tables(Tabla)
    On Error GoTo 0
    If tbl Is Nothing Then
      If Crear <> cpNinguna Then
        Set tbl = New ADOX.Table
        If Crear = cpClavePrincipal Then
          With tbl
            .Name = strTblClaves
            .Columns.Append "Clave", adVarWChar, 30
            .Columns.Append "Valor", adInteger
          End With
        ElseIf Crear = cpCola Then
          With tbl
            .Name = strTblColas
            .Columns.Append "Clave", adVarWChar, 30
            .Columns.Append "Usuario", adVarWChar, 64
            .Columns.Append "Fecha", adDate
          End With
        ElseIf Crear = cpSecuencia Then
          With tbl
            .Name = strTblSecuencias
            .Columns.Append "Clave", adVarWChar, 30
            .Columns.Append "Base", adVarWChar, 128
            .Columns.Append "Valor", adInteger
          End With
        End If
        cat.Tables.Append tbl
        cat.Tables.Refresh
        CierraCnx
        AbreCnx True
        bolEncontrada = True
      End If
    Else
      bolEncontrada = True
    End If
    
    
    Set cat = Nothing
  End If
  
  BuscaTabla = bolEncontrada
End Function

Public Property Get CadenaConexion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CadenaConexion
    CadenaConexion = mvarCadenaConexion
End Property

Public Property Let CadenaConexion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CadenaConexion = 5
    If cnx.State = adStateOpen Then
      cnx.Close
    End If
    mvarCadenaConexion = vData
    cnx.ConnectionString = mvarCadenaConexion
End Property

Public Sub CerrarConexion()
  CierraCnx
End Sub

Private Sub CierraCnx()
  If cnx.State = adStateOpen Then
    cnx.Close
  End If
End Sub

Private Sub Class_Initialize()
  strTblClaves = "GpoClvPr"  ' Grupo de claves principales
  strTblColas = "ClaClvPr"   ' Cola de la clave principal
  strTblSecuencias = "GpoSecue" ' Grupo de secuencias
  Set cnx = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
  CierraCnx
  Set cnx = Nothing
End Sub

Public Property Get Conexion() As ADODB.Connection
  Set Conexion = cnx
End Property

Public Property Set Conexion(ByVal vNewValue As ADODB.Connection)
  If cnx.State = adStateOpen Then
    cnx.Close
  End If
  mvarCadenaConexion = vNewValue.ConnectionString
  cnx.ConnectionString = mvarCadenaConexion
End Property

Public Function ConsultaDeEjecucion(CualConsulta As String, Optional MantenerAbierto As Boolean = False) As Long
  ' Var
  Dim lngFilas As Long
  
  If AbreCnx() Then
    On Error Resume Next
    cnx.Execute CualConsulta, lngFilas
    On Error GoTo 0
    If Not MantenerAbierto Then
      CierraCnx
    End If
  End If
  ConsultaDeEjecucion = lngFilas
End Function

Public Function ConsultaDeResultados(CualConsulta As String, Optional MantenerAbierto As Boolean = False) As Tabla
  ' Var
  Dim cda As ConjuntoDatos
  Dim Rs As ADODB.Recordset
  Dim tbla As Tabla
  Dim bolError As Boolean
  
  If AbreCnx() Then
    On Error Resume Next
    Set Rs = cnx.Execute(CualConsulta)
    If Err.Number <> 0 Then
      bolError = True
    End If
    On Error GoTo 0
    If Not bolError Then
      Set cda = New ConjuntoDatos
      Set tbla = cda.Llenar(Rs)
      Rs.Close
    End If
    
    Set Rs = Nothing
    If Not MantenerAbierto Then
      CierraCnx
    End If
  End If
  
  Set ConsultaDeResultados = tbla
End Function

Public Property Get Contrasenia() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Contrasenia
    Contrasenia = mvarContrasenia
End Property

Public Property Let Contrasenia(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Contrasenia = 5
    mvarContrasenia = vData
End Property

Private Sub doSalirDeLaFormacion(CualClave As String, CualUsuario As String)
  ' Var
  Dim strSQL As String
  
  ' Saca al usuario de la formación
  strSQL = "DELETE FROM " & strTblColas & " WHERE Clave = '" & CualClave & "' AND Usuario ='" & CualUsuario & "';"
  cnx.Execute strSQL
End Sub

Public Function EjecutaCMDDeEjecucion(Tipo As TipoDeConsultaDeEjecucionEnum, CualTabla As String, Campos() As String, Valores() As String, Optional Condicion As String = "") As Long
  ' Var
  Dim cmd As ADODB.Command
  Dim par As ADODB.Parameter
  Dim strSQL As String
  Dim strParametro As String
  Dim i As Integer, j As Integer
  Dim lngRegistros As Integer
  Dim colCampos As Collection
  Dim cmp As Columna
    
  Set colCampos = ObtieneCamposDeTabla(CualTabla)
  strSQL = ObtieneSQLDeEjecucion(Tipo, CualTabla, Campos(), Valores(), Condicion, True)
  If strSQL <> "" Then
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnx.ConnectionString
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    cmd.Prepared = True
    j = UBound(Valores)
    For i = 0 To UBound(Valores)
      strParametro = "@" & Campos(i)
      Set cmp = colCampos(Campos(i))
      Select Case cmp.Tipo
        Case adLongVarChar, adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
          Set par = cmd.CreateParameter(strParametro, cmp.Tipo, adParamInput, cmp.Tamanio)
        Case Else
          Set par = cmd.CreateParameter(strParametro, cmp.Tipo, adParamInput)
      End Select
      
      par.Value = Valores(i)
      cmd.Parameters.Append par
    Next i
  
    ' Ejecuta el comando
    AbreCnx
    cmd.Execute lngRegistros, Valores()
    CierraCnx
    Set par = Nothing
    Set cmd = Nothing
    EjecutaCMDDeEjecucion = lngRegistros
    
  End If
  
End Function

Private Function EsperarTurno(CualCP As String, Usuario As String, Optional EsperaMaxima As Single = 5) As Boolean
  ' Var
  Dim bolContinuar As Boolean
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim strUsuario As String
  Dim intCuantaEspera As Integer
  Dim bolEsperaExitosa As Boolean
  
  strSQL = "SELECT Clave, Usuario, Fecha FROM " & strTblColas & " WHERE Clave = '" & CualCP & "' ORDER BY Fecha DESC;"
  Do
    Sleep COMPAS_DE_ESPERA
    Set Rs = cnx.Execute(strSQL)
    If StrComp("" & Rs!Usuario, Usuario) = 0 Then
      bolContinuar = True
      bolEsperaExitosa = True
    ElseIf DateDiff("s", CDate(Rs!Fecha), Now) > EsperaMaxima And intCuantaEspera < 1 Then
        strUsuario = Rs!Usuario
        doSalirDeLaFormacion CualCP, strUsuario
        intCuantaEspera = intCuantaEspera + 1
    ElseIf DateDiff("s", CDate(Rs!Fecha), Now) > EsperaMaxima And intCuantaEspera > 0 Then
      bolContinuar = True
    End If
    Rs.Close
  Loop Until bolContinuar
  Set Rs = Nothing
  EsperarTurno = bolEsperaExitosa
End Function

Private Function Formarse(NombreCP As String) As String
  ' Esta función devolverá el GUID correspondiente a este usuario
  '  para el lugar que le toca en la formación o cola de procesos.
  ' Var
  Dim strGUID As String
  Dim bolEstaLaTabla As Boolean
  Dim strSQL As String
  
  strGUID = ObtGUID()
  bolEstaLaTabla = BuscaTabla(strTblColas, cpCola)
  If bolEstaLaTabla Then
    strSQL = "INSERT INTO " & strTblColas & " (Clave, Usuario, Fecha) VALUES ('" & NombreCP & "', '" & strGUID & "', #" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "#);"
    cnx.Execute strSQL
    Formarse = strGUID
  End If
  
End Function

Public Function ObtGUID() As String
  '(c) 2002 David Garza
  
  Dim udtGUID As GUID
  
  If CoCreateGuid(udtGUID) = 0 Then
    ObtGUID = _
    String(8 - Len(Hex$(udtGUID.data1)), "0") & Hex$(udtGUID.data1) & _
    String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
    String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
    IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
    IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
    IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
    IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
    IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
    IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
    IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
    IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
  End If


End Function

Public Function ObtieneCamposDeTabla(CualTabla As String) As Collection
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim cl As ADOX.Column
  Dim colCampos As Collection
  Dim c As Columna
  
  ' Abro el catálogo de la base de datos
  If Len(mvarCadenaConexion) > 0 Then
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = mvarCadenaConexion 'cnx.ConnectionString
    
    ' Obtengo la tabla que me piden
    On Error Resume Next
    Set tbl = cat.Tables(CualTabla)
    On Error GoTo 0
    
    ' ¿Se encontró la tabla?
    If Not tbl Is Nothing Then
      Set colCampos = New Collection
      For Each cl In tbl.Columns
        Set c = New Columna
        c.Nombre = cl.Name
        c.Atributos = cl.Attributes
        c.EscalaNumerica = cl.NumericScale
        ' c.Estado = cl.Status
        ' c.FormatoDatos = cl.DataFormat
        c.Key = cl.Name
        c.Precision = cl.Precision
        c.Tabla = CualTabla
        c.Tamanio = cl.DefinedSize
        c.TamanioDefinido = cl.DefinedSize
        c.Tipo = cl.Type
        ' c.valor = cl.Value
        ' c.ValorOriginal = cl.OriginalValue
        ' c.ValorSubyacente = cl.UnderValue
        colCampos.Add c, c.Key
      Next cl
    End If  ' tbl
    
    Set cat = Nothing
  End If  ' cnx.ConnectionString
  
  Set ObtieneCamposDeTabla = colCampos

End Function

Private Function ObtieneConsulta(ByVal CualTabla As String, ByVal CualCadena As String) As String
  ' Var
  Dim strConsulta As String
  Dim Cads As Variant
  Dim strCampo As String
  Dim strValor As String
  Dim cCol As Columna
  Dim i As Integer
  Dim l As Integer
  Dim X As Integer
  Dim lngTipo As Long
  Dim colCampos As Collection
  Dim strCarIniFin As String
  Dim tvaTipoValor As Long
  Dim strError As String
  Dim bolCorrecto As Boolean
  Dim bolPrueba As Boolean
  
  If CualCadena <> "" Then
    Cads = Split(CualCadena, "|")
    l = UBound(Cads)
    
    Set colCampos = ObtieneCamposDeTabla(CualTabla)
        
    For i = 0 To l
      X = InStr(Cads(i), "=")
      If X > 0 Then
        strCampo = Left(Cads(i), X - 1)
        strValor = Mid(Cads(i), X + 1)
        strCampo = Trim(strCampo)
        strValor = Trim(strValor)
        Set cCol = Nothing
        On Error Resume Next
        Set cCol = colCampos(strCampo)
        On Error GoTo 0
        If Not cCol Is Nothing Then  ' Si existe el campo dado
          lngTipo = cCol.Tipo
          
          Select Case lngTipo
            Case adLongVarChar, adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
              strCarIniFin = "'"
              tvaTipoValor = Cadena
            Case adDate, adDBDate, adDBTime, adDBTimeStamp
              strCarIniFin = "'"
              tvaTipoValor = Fecha
            Case adBoolean
              strCarIniFin = ""
              tvaTipoValor = Condicional
            Case Else
              strCarIniFin = ""
              tvaTipoValor = Numerico
          End Select
          
          ' Verifica si el valor corresponde con el tipo de campo
          bolCorrecto = True
          Select Case tvaTipoValor
            Case Numerico
              If Not IsNumeric(strValor) Then
                bolCorrecto = False
                strError = strError & vbCrLf & strCampo & ": " & strValor & " no es numérico"
              End If
            Case Fecha
              If Not IsDate(strValor) Then
                bolCorrecto = False
                strError = strError & vbCrLf & strCampo & ": " & strValor & " no es una fecha"
              Else
                strValor = Format(CDate(strValor), "yyyy/mm/dd")
              End If
            Case Condicional
              On Error Resume Next
              bolPrueba = CBool(strValor)
              If Err.Number <> 0 Then
                bolCorrecto = False
                strError = strError & vbCrLf & strCampo & ": " & strValor & " no es condicional (booleano)"
              Else
                strValor = IIf(CBool(bolPrueba), 1, 0)
              End If
              On Error GoTo 0
            Case Else
          End Select
          
          ' Agrega, en su caso, la cadena.
          If bolCorrecto Then
            If Len(strConsulta) > 0 Then
              strConsulta = strConsulta & " AND "
            End If
            strConsulta = strConsulta & strCampo & "=" & strCarIniFin & strValor & strCarIniFin
          End If
        Else
          strError = strError & vbCrLf & "El campo " & strCampo & " indicado en la lista de condiciones, no existe."
        End If
      End If
    Next i
    
    If Len(strError) > 0 Then
      MsgBox "Error al generar la consulta final. Puede no dar los resultados esperados." & vbCrLf & strError, vbExclamation, "Error en ObtieneConsulta"
    End If
  End If
  
  ObtieneConsulta = strConsulta
End Function

Public Function ObtieneConsultaSQLDeSeleccion(CualTabla As String, CualesCamposBase As String, CualesValores As String, Optional OperadoresDeComparacion As String = "=", Optional NexoLogicoPredeterminado As String = "AND", Optional ObtenerCampos As String = "*", Optional Comodines As String = "") As String
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim cl As ADOX.Column
  Dim colCampos As Collection
  Dim varCampos As Variant
  Dim varValores As Variant
  Dim varOperadores As Variant
  Dim varCamposPorObtener As Variant
  Dim adtTiposCampos() As ADOX.DataTypeEnum
  Dim bolCorrecto As Boolean
  Dim i As Integer, b As Integer, a As Integer, l As Integer
  Dim strCmp As String
  Dim strError As String
  Dim strSQL As String
  Dim strCad As String
  Dim strCarIniFin As String
  Dim strOperador As String
  Dim tvaTipoValor As TipoDeValorPorAnalizarEnum
  Dim bolPrueba As Boolean
  Dim bolEsAccess As Boolean
  
  varCampos = Split(CualesCamposBase, ",")
  varValores = Split(CualesValores, ",")
  varCamposPorObtener = Split(ObtenerCampos, ",")
  varOperadores = Split(OperadoresDeComparacion, ",")
  If UBound(varCampos) > (-1) Then
    ReDim adtTiposCampos(0 To UBound(varCampos)) As ADOX.DataTypeEnum
  
    If UBound(varCampos) = UBound(varValores) Then
    
      If (UBound(varOperadores) = 0) Or (UBound(varOperadores) = UBound(varCampos)) Then  ' Hay tantos operadores como campos, o hay un operador predeterminado
        NexoLogicoPredeterminado = UCase(Trim(NexoLogicoPredeterminado))
      
        If NexoLogicoPredeterminado = "AND" Or NexoLogicoPredeterminado = "OR" Then
          NexoLogicoPredeterminado = " " & NexoLogicoPredeterminado & " "
          ' Abro el catálogo de la base de datos
          If Len(cnx.ConnectionString) > 0 Then
            Set cat = New ADOX.Catalog
            cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
            
            bolEsAccess = (InStr(1, cnx.ConnectionString, "Microsoft.Jet", vbTextCompare) > 0)
            If bolEsAccess And Len(Comodines) = 0 Then
              Comodines = "*?"
            ElseIf Not bolEsAccess And Len(Comodines) = 0 Then
              Comodines = "%_"
            End If
            
            ' Obtengo la tabla que me piden
            On Error Resume Next
            Set tbl = cat.Tables(CualTabla)
            If Err.Number <> 0 Then
              strError = "No existe la tabla " & CualTabla & "."
            End If
            On Error GoTo 0
            
            ' ¿Se encontró la tabla?
            If Not tbl Is Nothing Then
              
              ' ¿Están los campos por buscar?
              b = LBound(varCampos)
              a = UBound(varCampos)
              bolCorrecto = True
              On Error Resume Next
              For i = b To a
                varCampos(i) = Trim(varCampos(i))
                Set cl = tbl.Columns(varCampos(i))
                If Err.Number <> 0 Then
                  strError = strError & vbCrLf & varCampos(i)
                  Err.Clear
                  bolCorrecto = False
                Else
                  adtTiposCampos(i) = cl.Type
                End If
              Next i
              If Not bolCorrecto Then
                strError = "No existen los siguientes campos para hacer la búsqueda: " & vbCrLf & strError
              End If
              On Error GoTo 0
              
              If bolCorrecto Then
                
                ' ¿Están los campos por obtener?
                b = LBound(varCamposPorObtener)
                a = UBound(varCamposPorObtener)
                Set colCampos = Me.ObtieneCamposDeTabla(CualTabla)
                On Error Resume Next
                For i = b To a
                  varCamposPorObtener(i) = Trim(varCamposPorObtener(i))
                  If varCamposPorObtener(i) <> "*" Then
                    strCmp = colCampos(varCamposPorObtener(i)).Nombre
                    If Err.Number <> 0 Then
                      strError = strError & vbCrLf & varCamposPorObtener(i)
                      Err.Clear
                      bolCorrecto = False
                    Else
                      varCamposPorObtener(i) = strCmp
                    End If
                  End If
                Next i
                If Not bolCorrecto Then
                  strError = "No existen los siguientes campos de los que se pidió su valor: " & vbCrLf & strError
                End If
                On Error GoTo 0
                
                If bolCorrecto Then
                
                  ' Forma la cadena SQL
                  strSQL = "SELECT "
                  For i = b To a
                    strSQL = strSQL & varCamposPorObtener(i) & IIf(i < a, ", ", "")
                  Next i
                  
                  strSQL = strSQL & " FROM " & CualTabla & " WHERE "
                  b = LBound(varCampos)
                  a = UBound(varCampos)
                  For i = b To a
                    strCad = varCampos(i)
                    Select Case UCase(adtTiposCampos(i))
                      Case adLongVarChar, adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
                        strCarIniFin = "'"
                        tvaTipoValor = Cadena
                      Case adDate, adDBDate, adDBTime, adDBTimeStamp
                        If bolEsAccess Then
                          strCarIniFin = "#"
                        Else
                          strCarIniFin = "'"
                        End If
                        tvaTipoValor = Fecha
                      Case adBoolean
                        strCarIniFin = ""
                        tvaTipoValor = Condicional
                      Case Else
                        strCarIniFin = ""
                        tvaTipoValor = Numerico
                    End Select
                    
                    ' Verifica si el valor corresponde con el tipo de campo
                    Select Case tvaTipoValor
                      Case Numerico
                        If Not IsNumeric(varValores(i)) Then
                          bolCorrecto = False
                          strError = strError & vbCrLf & varCampos(i) & ": " & varValores(i) & " no es numérico"
                        End If
                      Case Fecha
                        If Not IsDate(varValores(i)) Then
                          bolCorrecto = False
                          strError = strError & vbCrLf & varCampos(i) & ": " & varValores(i) & " no es una fecha"
                        Else
                          varValores(i) = CDate(varValores(i))
                          varValores(i) = Format(varValores(i), "yyyy/mm/dd")
                        End If
                      Case Condicional
                        On Error Resume Next
                        bolPrueba = CBool(varValores(i))
                        If Err.Number <> 0 Then
                          bolCorrecto = False
                          strError = strError & vbCrLf & varCampos(i) & ": " & varValores(i) & " no es condicional (booleano)"
                        Else
                          If bolEsAccess Then
                            varValores(i) = IIf(bolPrueba, -1, 0)
                          Else
                            varValores(i) = IIf(bolPrueba, 1, 0)
                          End If
                        End If
                        On Error GoTo 0
                      Case Else
                      
                    End Select
                    
                    ' Verifica el tipo de comparación
                    If UBound(varOperadores) > 1 Then
                      strOperador = varOperadores(i)
                    Else
                      strOperador = varOperadores(0)
                    End If
                    
                    For l = 1 To Len(Comodines)
                      If InStr(varValores(i), Mid(Comodines, l, 1)) > 0 Then
                        strOperador = " LIKE "
                      End If
                    Next l
                    
                    strCad = strCad & strOperador & strCarIniFin & varValores(i) & strCarIniFin & IIf(i < a, NexoLogicoPredeterminado, ";")
                    strSQL = strSQL & strCad
                  Next i
                  
                  If Not bolCorrecto Then  ' Algún tipo de dato no correspondió
                    strSQL = ""
                    strError = "Los valores de los siguientes campos no son consistentes con el tipo de dato del campo:" & vbCrLf & strError
                  End If
                End If  ' bolCorrecto (los campos por obtener sí existen)
              End If  ' bolCorrecto (Los campos que son la base de la búsqueda sí existen)
            End If ' La tabla sí existe
          End If ' Sí hay asignada una cadena de conexión
        Else
          strError = "El nexo lógico predeterminado debe ser AND u OR"
        End If
      Else
        strError = "Si se establece más de un operador de comparación, deben ser la misma cantidad que los campos."
      End If
    ElseIf UBound(varCampos) > UBound(varValores) Then
      strError = "Hay más campos que valores en los campos base para la búsqueda."
    Else
      strError = "Hay más valores que campos en los campos base para la búsqueda."
    End If ' Es igual la cantidad de campos que de valores
  Else
    bolCorrecto = True
    If Len(cnx.ConnectionString) > 0 Then
      Set cat = New ADOX.Catalog
      cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
      
      ' Obtengo la tabla que me piden
      On Error Resume Next
      Set tbl = cat.Tables(CualTabla)
      If Err.Number <> 0 Then
        bolCorrecto = False
        strError = "No existe la tabla " & CualTabla & "."
      End If
      Err.Clear
      On Error GoTo 0
    Else
      bolCorrecto = False
      strError = "No se asignó una cadena de conexión"
    End If
    
    If bolCorrecto Then
    
      ' ¿Están los campos por obtener?
      b = LBound(varCamposPorObtener)
      a = UBound(varCamposPorObtener)
      Set colCampos = Me.ObtieneCamposDeTabla(CualTabla)
      On Error Resume Next
      For i = b To a
        varCamposPorObtener(i) = Trim(varCamposPorObtener(i))
        If varCamposPorObtener(i) <> "*" Then
          strCmp = colCampos(varCamposPorObtener(i)).Nombre
          If Err.Number <> 0 Then
            strError = strError & vbCrLf & varCamposPorObtener(i)
            Err.Clear
            bolCorrecto = False
          Else
            varCamposPorObtener(i) = strCmp
          End If
        End If
      Next i
    End If

    If bolCorrecto Then
      strSQL = "SELECT " & ObtenerCampos & " FROM " & CualTabla & ";"
    Else
      strSQL = ""
    End If

  End If
  
  If Len(strError) > 0 Then
    MsgBox strError, vbExclamation, "Notificación de error"
  End If
  
  ObtieneConsultaSQLDeSeleccion = strSQL
End Function

Public Function ObtieneCP(Tabla) As String
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim pk As ADOX.Key
  Dim strNombreCP As String
  Dim col As ADOX.Column
  
  ' Abro el catálogo de la base de datos
  If Len(mvarCadenaConexion) > 0 Then
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
    
    ' Obtengo la tabla que me piden
    On Error Resume Next
    Set tbl = cat.Tables(Tabla)
    On Error GoTo 0
    
    ' ¿Se encontró la tabla?
    strNombreCP = ""
    If Not tbl Is Nothing Then
      For Each pk In tbl.Keys
        If pk.Type = adKeyPrimary Then
          If pk.Columns.Count = 1 Then    ' Sólo los que tengan un campo en la clave principal
            Set col = tbl.Columns(pk.Columns(0).Name)
            'If col.Type = adInteger Then   ' Sólo los campos Integer (Long) y String
              strNombreCP = col.Name
              Exit For
            'End If
          End If
        End If
      Next pk
    End If
    
    Set cat = Nothing
  End If
  
  ObtieneCP = strNombreCP

End Function

Public Function ObtieneDatosDeFormulario(CualFrm As Object) As String()
  ' Var
  ReDim strCamposValores(0 To 0) As String
  Dim ctl As Control
  Dim strCampoValor As String
  Dim strTControl As String
  Dim i As Integer
  Dim cProp As New PropVirtuales
  
  For Each ctl In CualFrm.Controls
    strCampoValor = cProp.getPropiedad(ctl.Tag, "Campo")
    If strCampoValor <> "" Then
      If UCase(cProp.getPropiedad(ctl.Tag, "Guardar")) <> "NO" Then
        strTControl = UCase(TypeName(ctl))
        Select Case strTControl
          Case "TEXTBOX"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Trim(ctl.Text)
          Case "COMBOBOX", "LISTBOX"
            If ctl.ListIndex <> -1 Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.ItemData(ctl.ListIndex)
            Else
              strCampoValor = ""
            End If
          Case "DTPICKER"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Format(ctl.Value, "Short Date")
          Case "IMAGECOMBO"
            If Not ctl.SelectedItem Is Nothing Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.SelectedItem.Tag
            End If
          Case "CHECKBOX"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Value
          Case "LABEL"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Caption
        End Select
        i = i + 1
        If i > UBound(strCamposValores) Then
          ReDim Preserve strCamposValores(0 To i + 10) As String
        End If
        strCamposValores(i) = strCampoValor
      End If
    End If
  Next ctl
  
  ReDim Preserve strCamposValores(0 To i) As String
  
  ObtieneDatosDeFormulario = strCamposValores
End Function

Public Function ObtieneDatosDeFormularioCol(CualFrm As Object) As Collection
  ' Var
  Dim colCamposValores As New Collection
  Dim ctl As Control
  Dim strCampoValor As String
  Dim strTControl As String
  Dim cProp As New PropVirtuales
  
  For Each ctl In CualFrm.Controls
    strCampoValor = cProp.getPropiedad(ctl.Tag, "Campo")
    If strCampoValor <> "" Then
      If UCase(cProp.getPropiedad(ctl.Tag, "Guardar")) <> "NO" Then
        strTControl = UCase(TypeName(ctl))
        Select Case strTControl
          Case "TEXTBOX"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Trim(ctl.Text)
          Case "COMBOBOX", "LISTBOX"
            If ctl.ListIndex > -1 Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.ItemData(ctl.ListIndex)
            Else
              strCampoValor = ""
            End If
          Case "DTPICKER"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Format(ctl.Value, "mm/dd/yyyy")
          Case "IMAGECOMBO"
            If Not ctl.SelectedItem Is Nothing Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.SelectedItem.Tag
            End If
          Case "CHECKBOX"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Value
          Case "LABEL"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Caption
        End Select
        colCamposValores.Add strCampoValor
      End If
    End If
  Next ctl
  
  Set ObtieneDatosDeFormularioCol = colCamposValores
End Function

Public Function ObtieneDatosDeFormularioMat(CualFrm As Object, Campos() As String, Valores() As String) As Boolean
  ' Var
  Dim ctl As Control
  Dim strCampoValor As String
  Dim strTControl As String
  Dim i As Integer
  Dim cProp As New PropVirtuales
  
  i = -1
  For Each ctl In CualFrm.Controls
    strCampoValor = cProp.getPropiedad(ctl.Tag, "Campo")
    If strCampoValor <> "" Then
      If UCase(cProp.getPropiedad(ctl.Tag, "Guardar")) <> "NO" Then
        strTControl = UCase(TypeName(ctl))
        Select Case strTControl
          Case "TEXTBOX"
            If Trim(ctl.Text) <> "" Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Trim(ctl.Text)
            Else
              strCampoValor = ""
            End If
          Case "COMBOBOX", "LISTBOX"
            If ctl.ListIndex > (-1) Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.ItemData(ctl.ListIndex)
            Else
              strCampoValor = ""
            End If
          Case "DTPICKER"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & Format(ctl.Value, "Short Date")
          Case "IMAGECOMBO"
            If Not ctl.SelectedItem Is Nothing Then
              strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.SelectedItem.Tag
            End If
          Case "CHECKBOX"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Value
          Case "LABEL"
            strCampoValor = "Campo=" & strCampoValor & "|Valor=" & ctl.Caption
        End Select
        
        If strCampoValor <> "" Then
          i = i + 1
          If i > UBound(Campos) Then
            ReDim Preserve Campos(0 To i + 10) As String
            ReDim Preserve Valores(0 To i + 10) As String
          End If
          Campos(i) = cProp.getPropiedad(strCampoValor, "Campo")
          Valores(i) = cProp.getPropiedad(strCampoValor, "Valor")
        End If
      End If
    End If
  Next ctl
  
  If i > 0 Then
    ReDim Preserve Campos(0 To i) As String
    ReDim Preserve Valores(0 To i) As String
  End If
  
  ObtieneDatosDeFormularioMat = (i > 0)
End Function

Public Function ObtieneElID(Tabla As String, Campo As String, Optional UsarCola As Boolean = False, Optional CrearTablaClaves As Boolean = True, Optional MantenerAbierto As Boolean = False, Optional EsSeguro As Boolean = False) As Long
  ' Var
  Dim lngNuevoID() As Long
  
  lngNuevoID = ObtieneLosID(Tabla, Campo, 1, UsarCola, CrearTablaClaves, MantenerAbierto, EsSeguro)
  ObtieneElID = lngNuevoID(1)
End Function

  
Public Function ObtieneId(Tabla As String, Optional UsarCola As Boolean = False, Optional CrearTablaClaves As Boolean = True) As Long
  ' Var
  Dim lngNuevoID As Long
  Dim strNombreCP As String
  
  ' Obtenga el nombre del índice principal
  lngNuevoID = 0
  strNombreCP = ObtieneCP(Tabla)
  
  ' ¿Se encontró el nombre del índice?
  If Len(strNombreCP) <> 0 Then
    lngNuevoID = ObtieneElID(Tabla, strNombreCP, UsarCola, CrearTablaClaves)
  End If
  
  ObtieneId = lngNuevoID
End Function

Public Function ObtieneLosID(Tabla As String, Campo As String, Cuantos As Integer, Optional UsarCola As Boolean = False, Optional CrearTablaClaves As Boolean = True, Optional MantenerAbierto As Boolean = False, Optional EsSeguro As Boolean = False) As Long()
    ' Var
    Dim lngModif As Long
    Dim lngNuevoID As Long
    Dim Rs As ADODB.Recordset
    Dim strGUID As String
    Dim bolProceder As Boolean
    Dim cmdActualizar As ADODB.Command
    Dim cmdInsertar As ADODB.Command
    Dim cmdObtener As ADODB.Command
    Dim cmdVerificar As ADODB.Command
    Dim cmdActualizarCP As ADODB.Command
    Dim cmdObtenerMax As ADODB.Command
    Dim bolVerificar As Boolean
    Dim MatrizID() As Long
    Dim i As Integer
    Dim lngCualEsElMax As Long
    
  If Cuantos > 0 Then
    ' Abra la base de datos
    If AbreCnx() Then
      
      If CrearTablaClaves Then
        bolProceder = BuscaTabla(strTblClaves, cpClavePrincipal)
      ElseIf Not EsSeguro Then  ' Si es seguro, entonces se asume que la tabla de claves ya existe
        bolProceder = BuscaTabla(strTblClaves)
      Else
        bolProceder = True
      End If
      
      If bolProceder Then  ' Proceder, dado que ya existe la tabla de claves principales
      
        ' Fórmese, en su caso, en la fila
        If UsarCola Then
          strGUID = Formarse(Campo)
          bolProceder = EsperarTurno(Campo, strGUID)
        Else
          bolProceder = True
        End If
        
        If bolProceder Then  ' Proceder con la generación de la clave principal
        
          ' Genera las consultas precompiladas
          Set cmdActualizar = New ADODB.Command
          Set cmdInsertar = New ADODB.Command
          Set cmdObtener = New ADODB.Command
          Set cmdVerificar = New ADODB.Command
          Set cmdActualizarCP = New ADODB.Command
          Set cmdObtenerMax = New ADODB.Command
          ReDim MatrizID(1 To Cuantos) As Long
          
          With cmdActualizar
            Set .ActiveConnection = cnx
            .CommandText = "UPDATE " & strTblClaves & " SET Valor = Valor + 1 WHERE Clave = ?;"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 30)
            .Prepared = True
          End With
          
          With cmdInsertar
            Set .ActiveConnection = cnx
            .CommandText = "INSERT INTO " & strTblClaves & " (Clave, Valor) VALUES (?, ?);"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 30)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput)
            .Prepared = True
          End With
          
          With cmdObtener
            Set .ActiveConnection = cnx
            .CommandText = "SELECT Valor FROM " & strTblClaves & " WHERE Clave = ?;"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 30)
            .Prepared = True
          End With
          
          With cmdVerificar
            Set .ActiveConnection = cnx
            .CommandText = "UPDATE " & Tabla & " SET " & Campo & " = ? WHERE " & Campo & " = ?;"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adInteger, adParamInput)
            .Parameters.Append .CreateParameter(, adInteger, adParamInput)
            .Prepared = True
          End With
          
          With cmdActualizarCP
            Set .ActiveConnection = cnx
            .CommandText = "UPDATE " & strTblClaves & " SET Valor=? WHERE Clave=?;"
            .CommandType = adCmdText
            .Parameters.Append .CreateParameter(, adInteger, adParamInput)
            .Parameters.Append .CreateParameter(, adChar, adParamInput, 30)
            .Prepared = True
          End With
          
          With cmdObtenerMax
            Set .ActiveConnection = cnx
            .CommandText = "SELECT MAX(" & Campo & ") AS CPMax FROM " & Tabla & ";"
            .CommandType = adCmdText
            .Prepared = True
          End With
    
          ' Actualice el índice
          cmdActualizar.Parameters(0).Value = Campo
          
          ' strSql = "UPDATE " & strTblClaves & " SET Valor = Valor + 1 WHERE Clave = '" & strNombreCP & "';"
          ' cnx.Execute strSql, lngModif
          cmdActualizar.Execute lngModif
          
          If lngModif = 0 Then
            ' Si no encontró el índice, delo de alta.
            ' strSql = "INSERT INTO " & strTblClaves & " (Clave, Valor) VALUES ('" & strNombreCP & "', 1);"
            ' cnx.Execute strSql
            cmdInsertar.Parameters(0).Value = Campo
            cmdInsertar.Parameters(1).Value = 1
            cmdInsertar.Execute lngModif
          End If  ' lngModif
          
          ' Ahora, obtenga el nuevo valor de la clave principal
          ' strSql = "SELECT Valor FROM " & strTblClaves & " WHERE Clave = '" & strNombreCP & "';"
          ' Set Rs = cnx.Execute(strSql)
          cmdObtener.Parameters(0).Value = Campo
          cmdActualizarCP.Parameters(1).Value = Campo
          
          ' Determine si debe verificar el valor del ID de modo que no
          '  choque con un valor existente.
          If Cuantos > 1 Then
            ' Si la clave ya podría existir, entonces obtenga el número mayor que haya en el ID de la tabla
            ' strSql = "SELECT MAX(" & strNombreCP & ") AS CPMax FROM " & Tabla & ";"
            ' Set Rs = cnx.Execute(strSql)
            Set Rs = cmdObtenerMax.Execute()
            lngCualEsElMax = CLng(Val("" & Rs!CPMax))
            Rs.Close
            bolVerificar = (lngCualEsElMax >= lngNuevoID)
          Else
            bolVerificar = True
          End If
            
          For i = 1 To Cuantos
            Set Rs = cmdObtener.Execute()
            lngNuevoID = Rs!Valor
            Rs.Close
            
            ' Verifica si la clave ya existe
            ' strSql = "UPDATE " & Tabla & " SET " & Campo & " = " & lngNuevoID & " WHERE " & Campo & " = " & lngNuevoID & ";"
            ' cnx.Execute strSql, lngModif
            If bolVerificar Then
              cmdVerificar.Parameters(0).Value = lngNuevoID
              cmdVerificar.Parameters(1).Value = lngNuevoID
              cmdVerificar.Execute lngModif
              If lngModif <> 0 Then
                If lngCualEsElMax = 0 Then
                  Set Rs = cmdObtenerMax.Execute()
                  
                  ' Al número mayor, súmele uno
                  lngNuevoID = Rs!CPMax
                  lngNuevoID = lngNuevoID + 1
                  Rs.Close
                Else
                  lngNuevoID = lngCualEsElMax + 1
                End If
                
                ' Ahora, actualice el índice en la tabla de Claves, para asegurarnos que permanecerá este nuevo valor.
                cmdActualizarCP.Parameters(0).Value = lngNuevoID
                cmdActualizarCP.Execute lngModif
                ' strSql = "UPDATE " & strTblClaves & " SET Valor = " & lngNuevoID & " WHERE Clave = '" & strNombreCP & "';"
                ' cnx.Execute strSql, lngModif
                bolVerificar = False
              End If  ' lngModif
            End If
            MatrizID(i) = lngNuevoID
            
            If i < Cuantos Then
              cmdActualizar.Execute lngModif
            End If
          Next i
          Set Rs = Nothing

        End If  ' bolProceder (Si encontró la tabla)
          
      End If  ' bolProceder (Si se realiza el proceso)
    
      ' Saque la petición de la fila
      If UsarCola Then
        doSalirDeLaFormacion Campo, strGUID
      End If
      
      Set cmdActualizar = Nothing
      Set cmdInsertar = Nothing
      Set cmdObtener = Nothing
      Set cmdVerificar = Nothing
      Set cmdActualizarCP = Nothing
      Set cmdObtenerMax = Nothing
      
      If Not MantenerAbierto Then
        CierraCnx  ' Cierre la conexión
      End If
    End If
  Else
    ReDim MatrizID(1 To 1) As Long
  End If
    
    ObtieneLosID = MatrizID
End Function

Public Function ObtieneSecuencia(Tabla As String, Campo As String, Optional ValoresBase As String = "", Optional UsarCola As Boolean = False, Optional CrearTablaSecuencias As Boolean = True, Optional MantenerAbierto As Boolean = False, Optional IniciarAPartirDelUltimoDeLaTabla As Boolean = True) As Long
  ' Var
  Dim lngNuevaSecuencia As Long
  Dim lngModif As Long
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim strGUID As String
  Dim bolProceder As Boolean
  Dim strVB As String
  Dim strConsulta As String
  
  ' ¿Se encontró el nombre del índice?
  If Len(Campo) <> 0 Then
    
    ' Abra la base de datos
    If AbreCnx() Then
    
      ' Fórmese en la fila
      If UsarCola Then
        strGUID = Formarse(Campo)
        bolProceder = EsperarTurno(Campo, strGUID)
      Else
        bolProceder = True
      End If
    
      If bolProceder Then  ' Proceder con la búsqueda del índice
         
        If CrearTablaSecuencias Then
          bolProceder = BuscaTabla(strTblSecuencias, cpSecuencia)
        Else
          bolProceder = BuscaTabla(strTblSecuencias)
        End If
        
        If bolProceder Then
          ' Genera la cadena de valores base
          If Len(ValoresBase) > 0 Then
            strVB = QuitaEspacios(ValoresBase)
          Else
            strVB = ""
          End If
          ' Genera la condición de consulta
          strConsulta = ObtieneConsulta(Tabla, ValoresBase)
          
          ' Actualice el índice
          strSQL = "UPDATE " & strTblSecuencias & " SET Valor = Valor + 1 WHERE Clave = '" & Campo & "' AND Base='" & strVB & "';"
          cnx.Execute strSQL, lngModif
          
          If lngModif = 0 Then
            ' Si no encontró el índice, delo de alta.
            strSQL = "INSERT INTO " & strTblSecuencias & " (Clave, Base, Valor) VALUES ('" & Campo & "', '" & strVB & "', 1);"
            cnx.Execute strSQL
          
            If IniciarAPartirDelUltimoDeLaTabla Then
              GoSub ObtieneElUltimoDeLaTabla
            End If
          End If  ' lngModif
          
          ' Ahora, obtenga el nuevo valor de la clave principal
          strSQL = "SELECT Valor FROM " & strTblSecuencias & " WHERE Clave = '" & Campo & "' AND Base = '" & strVB & "';"
          Set Rs = cnx.Execute(strSQL)
          lngNuevaSecuencia = Rs!Valor
          Rs.Close
          Set Rs = Nothing
          
          
          ' Verifica si la clave ya existe
          strSQL = "UPDATE " & Tabla & " SET " & Campo & " = " & lngNuevaSecuencia & " WHERE " & Campo & " = " & lngNuevaSecuencia
          If Len(strConsulta) Then
            strSQL = strSQL & " AND " & strConsulta
          End If
          strSQL = strSQL & ";"
          cnx.Execute strSQL, lngModif
          If lngModif <> 0 Then
            GoSub ObtieneElUltimoDeLaTabla
          End If  ' lngModif
        End If  ' bolProceder (Si encontró la tabla)
          
      End If  ' bolProceder (Si se realiza el proceso)
    
      ' Saque la petición de la fila
      If UsarCola Then
        doSalirDeLaFormacion Campo, strGUID
      End If
      
      If Not MantenerAbierto Then
        CierraCnx  ' Cierre la conexión
      End If
    End If ' Abrir base
  End If  ' Campo
  
  If lngNuevaSecuencia <= 0 Then
    lngNuevaSecuencia = -1  ' Hubo un error al hacer la consulta
                            ' Tal vez la base de datos esté dañada.
  End If
  
  ObtieneSecuencia = lngNuevaSecuencia
  
  Exit Function
  
ObtieneElUltimoDeLaTabla:
  ' Si la clave ya existe, entonces obtenga el número mayor que haya en el ID de la tabla
  strSQL = "SELECT MAX(" & Campo & ") AS CPMax FROM " & Tabla
  If Len(strConsulta) Then
    strSQL = strSQL & " WHERE " & strConsulta
  End If
  strSQL = strSQL & ";"
  Set Rs = cnx.Execute(strSQL)
  
  ' Al número mayor, súmele uno
  lngNuevaSecuencia = 0 + Val("" & Rs!CPMax)
  lngNuevaSecuencia = lngNuevaSecuencia + 1
  Rs.Close
  Set Rs = Nothing
  
  ' Ahora, actualice el índice en la tabla de Claves, para asegurarnos que permanecerá este nuevo valor.
  strSQL = "UPDATE " & strTblSecuencias & " SET Valor = " & lngNuevaSecuencia & " WHERE Clave = '" & Campo & "' AND Base='" & strVB & "';"
  cnx.Execute strSQL, lngModif
Return
End Function

Public Function ObtieneSQLDeEjecucion(Tipo As TipoDeConsultaDeEjecucionEnum, CualTabla As String, Campos() As String, Valores() As String, Optional Condicion As String = "", Optional SoloEstructura As Boolean = False) As String
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim cl As ADOX.Column
  Dim adtTiposCampos() As ADOX.DataTypeEnum
  Dim bolCorrecto As Boolean
  Dim i As Integer, b As Integer, a As Integer
  Dim strError As String
  ' Dim strCampoValor As String
  Dim strSQL As String
  Dim strCad As String
  Dim strCad2 As String
  Dim strCarIniFin As String
  Dim tvaTipoValor As TipoDeValorPorAnalizarEnum
  Dim bolPrueba As Boolean
  Dim strCP As String
  
  ReDim adtTiposCampos(0 To UBound(Valores)) As ADOX.DataTypeEnum
  
  If (Tipo = Actualizacion Or Tipo = Borrado) And Condicion = "" Then
    bolCorrecto = False
  Else
    bolCorrecto = True
  End If
  
  If bolCorrecto Then
    If UBound(Campos) = UBound(Valores) Then
      ' Abro el catálogo de la base de datos
      If Len(cnx.ConnectionString) > 0 Then
        Set cat = New ADOX.Catalog
        cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
        
        ' Obtengo la tabla que me piden
        On Error Resume Next
        Set tbl = cat.Tables(CualTabla)
        If Err.Number <> 0 Then
          strError = "No existe la tabla " & CualTabla & "."
        End If
        On Error GoTo 0
        
        ' ¿Se encontró la tabla?
        If Not tbl Is Nothing Then
          
          If Tipo = Insercion Then
            ' Verifica si la clave principal es una cadena de 32 bytes. Si es así,
            '  evidentemente se trata de un GUID. Por ende, isertará en strCad el nombre
            '  de la clave principal y un GUID.
            strCP = Me.ObtieneCP(CualTabla)
            
            If strCP <> "" Then
              Set cl = tbl.Columns(strCP)
              If cl.Type = adChar Or cl.Type = adLongVarChar Or cl.Type = adLongVarWChar Or cl.Type = adVarChar Or cl.Type = adVarWChar Or cl.Type = adWChar Then
                If cl.DefinedSize = 32 Then
                  ReDim Preserve Campos(0 To UBound(Campos) + 1) As String
                  ReDim Preserve Valores(0 To UBound(Valores) + 1) As String
                  Campos(UBound(Campos)) = strCP
                  Valores(UBound(Valores)) = Me.ObtGUID()
                  ReDim adtTiposCampos(0 To UBound(Valores)) As ADOX.DataTypeEnum
                End If
              End If
            End If
          End If
          
          If Tipo = Actualizacion Or Tipo = Insercion Then
            ' ¿Están los campos por buscar?
            b = LBound(Campos)
            a = UBound(Campos)
            bolCorrecto = True
            On Error Resume Next
            For i = b To a
              Campos(i) = Trim(Campos(i))
              Set cl = tbl.Columns(Campos(i))
              If Err.Number <> 0 Then
                strError = strError & vbCrLf & Campos(i)
                Err.Clear
                bolCorrecto = False
              Else
                adtTiposCampos(i) = cl.Type
              End If
            Next i
            If Not bolCorrecto Then
              strError = "No existen los siguientes campos para hacer la consulta: " & vbCrLf & strError
            End If
            On Error GoTo 0
          Else
            bolCorrecto = True
          End If
          
          If bolCorrecto Then
              If Tipo = Actualizacion Or Tipo = Insercion Then
                ' Revisa si los datos corresponden con los campos
                b = LBound(Campos)
                a = UBound(Campos)
                For i = b To a
                  Select Case UCase(adtTiposCampos(i))
                    Case adLongVarChar, adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
                      tvaTipoValor = Cadena
                    Case adDate, adDBDate, adDBTime, adDBTimeStamp
                      tvaTipoValor = Fecha
                    Case adBoolean
                      tvaTipoValor = Condicional
                    Case Else
                      tvaTipoValor = Numerico
                  End Select
                  
                  ' Verifica si el valor corresponde con el tipo de campo
                  Select Case tvaTipoValor
                    Case Numerico
                      If Not IsNumeric(Valores(i)) Then
                        bolCorrecto = False
                        strError = strError & vbCrLf & Valores(i) & ": " & Valores(i) & " no es numérico"
                      End If
                    Case Fecha
                      If Not IsDate(Valores(i)) Then
                        bolCorrecto = False
                        strError = strError & vbCrLf & Valores(i) & ": " & Valores(i) & " no es una fecha"
                      Else
                        Valores(i) = CDate(Valores(i))
                        Valores(i) = Format(Valores(i), "yyyy/mm/dd")
                      End If
                    Case Condicional
                      On Error Resume Next
                      bolPrueba = CBool(Valores(i))
                      If Err.Number <> 0 Then
                        bolCorrecto = False
                        strError = strError & vbCrLf & Valores(i) & ": " & Valores(i) & " no es condicional (booleano)"
                      Else
                        Valores(i) = IIf(bolPrueba, 1, 0)
                      End If
                      On Error GoTo 0
                    Case Else
                    
                  End Select
                Next i
              Else
                bolCorrecto = True
              End If
              
              If bolCorrecto Then  ' Todos los tipos de datos correspondieron
                If Tipo = Actualizacion Or Tipo = Insercion Then
                  For i = b To a
                    Select Case UCase(adtTiposCampos(i))
                      Case adLongVarChar, adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
                        strCarIniFin = "'"
                        tvaTipoValor = Cadena
                        Valores(i) = Left(Valores(i), tbl.Columns(Campos(i)).DefinedSize)
                      Case adDate, adDBDate, adDBTime, adDBTimeStamp
                        strCarIniFin = "#"
                        tvaTipoValor = Fecha
                      Case adBoolean
                        strCarIniFin = ""
                        tvaTipoValor = Condicional
                      Case Else
                        strCarIniFin = ""
                        tvaTipoValor = Numerico
                    End Select
                    
                    If Tipo = Actualizacion Then
                      If SoloEstructura Then
                        strCad = strCad & Campos(i) & "=@" & Campos(i) & IIf(i < a, ", ", " ")
                      Else
                        strCad = strCad & Campos(i) & "=" & strCarIniFin & Valores(i) & strCarIniFin & IIf(i < a, ", ", " ")
                      End If
                    ElseIf Tipo = Insercion Then
                      strCad = strCad & Campos(i) & IIf(i < a, ", ", "")
                      If SoloEstructura Then
                        strCad2 = strCad2 & "@" & Campos(i) & IIf(i < a, ", ", "")
                      Else
                        strCad2 = strCad2 & strCarIniFin & Valores(i) & strCarIniFin & IIf(i < a, ", ", "")
                      End If
                    End If
                  Next i
                Else
                  bolCorrecto = True
                End If
                
                If Tipo = Actualizacion Then
                  strSQL = "UPDATE " & CualTabla & " SET " & strCad & " WHERE " & Condicion & ";"
                ElseIf Tipo = Insercion Then
                  strSQL = "INSERT INTO " & CualTabla & " (" & strCad & ") VALUES (" & strCad2 & ");"
                ElseIf Tipo = Borrado Then
                  strSQL = "DELETE FROM " & CualTabla & " WHERE " & Condicion & ";"
                End If
              Else
                strSQL = ""
                strError = "Los valores de los siguientes campos no son consistentes con el tipo de dato del campo:" & vbCrLf & strError
              End If
          End If  ' bolCorrecto (Los campos que son la base de la búsqueda sí existen)
        End If ' La tabla sí existe
      End If ' Sí hay asignada una cadena de conexión
    ElseIf UBound(Valores) > UBound(Valores) Then
      strError = "Hay más campos que valores en los campos base para la búsqueda."
    Else
      strError = "Hay más valores que campos en los campos base para la búsqueda."
    End If ' Es igual la cantidad de campos que de valores
  Else
    strError = "En un SQL de actualización, debe especificar una condición"
  End If ' No hay una condición y se trata de una consulta de actualización
  
  If Len(strError) > 0 Then
    MsgBox strError, vbExclamation, "Notificación de error"
  End If
  
  ObtieneSQLDeEjecucion = strSQL

End Function

Public Function ObtieneSQLDeEjecucionStr(Tipo As TipoDeConsultaDeEjecucionEnum, CualTabla As String, Campos As String, Valores As String, Optional Condicion As String = "", Optional SoloEstructura As Boolean = False, Optional Separador As String = ",") As String
  ' Var
  Dim varCampos As Variant
  Dim varValores As Variant
  ReDim strCampos(0 To 0) As String
  ReDim strValores(0 To 0) As String
  Dim strSQL As String
  Dim i As Integer, j As Integer
  
  If Tipo = Actualizacion Or Tipo = Insercion Then
    varCampos = Split(Campos, Separador)
    varValores = Split(Valores, Separador)
    If UBound(varCampos) = UBound(varValores) Then
      If Tipo = Actualizacion Or Tipo = Insercion Then
        j = UBound(varCampos)
        ReDim strCampos(0 To j) As String
        ReDim strValores(0 To j) As String
        
        For i = 0 To j
          strCampos(i) = Trim(CStr(varCampos(i)))
          strValores(i) = Trim(CStr(varValores(i)))
        Next i
        Erase varCampos
        Erase varValores
      End If
      
      strSQL = ObtieneSQLDeEjecucion(Tipo, CualTabla, strCampos(), strValores(), Condicion, SoloEstructura)
    Else
      MsgBox "No hay la misma cantidad de campos que de valores", vbExclamation, "Error"
      strSQL = ""
    End If
  Else
    strSQL = ObtieneSQLDeEjecucion(Tipo, CualTabla, strCampos(), strValores(), Condicion, SoloEstructura)
  End If
  
  ObtieneSQLDeEjecucionStr = strSQL
End Function

Public Function ObtieneTablas() As Collection
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim colTablas As Collection
  
  ' Abro el catálogo de la base de datos
  If Len(mvarCadenaConexion) > 0 Then
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = mvarCadenaConexion ' cnx.ConnectionString
    
    Set colTablas = New Collection
    For Each tbl In cat.Tables
      If UCase(tbl.Type) = "TABLE" Then
        colTablas.Add tbl.Name, UCase(tbl.Name)
      End If
    Next tbl
    
    Set cat = Nothing
  End If  ' cnx.ConnectionString
  
  Set ObtieneTablas = colTablas

End Function

Public Function ExisteTabla(CualTabla As String) As Boolean
  ' Var
  Dim colTablas As Collection
  Dim strTabla As String
  
  Set colTablas = Me.ObtieneTablas
  
  On Error Resume Next
  strTabla = ""
  strTabla = colTablas(UCase(CualTabla))
  On Error GoTo 0
  
  ExisteTabla = (strTabla <> "")
End Function

Public Function ObtieneValores(CualTabla As String, CualesCamposBase As String, CualesValores As String, Optional OperadoresDeComparacion As String = "=", Optional NexoLogicoPredeterminado As String = "AND", Optional ObtenerCampos As String = "*", Optional Comodines As String = "", Optional MantenerAbierto As Boolean = False) As Tabla
  ' Var
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim cd As ConjuntoDatos
  Dim tbla As Tabla
  
  strSQL = ObtieneConsultaSQLDeSeleccion(CualTabla, CualesCamposBase, CualesValores, OperadoresDeComparacion, NexoLogicoPredeterminado, ObtenerCampos, Comodines)
  If Len(strSQL) > 0 Then
    ' Ejecuta la instrucción
    AbreCnx
    Set Rs = cnx.Execute(strSQL)
    
    If Not Rs.BOF And Not Rs.EOF Then
      Set cd = New ConjuntoDatos
      Set tbla = cd.Llenar(Rs, CualTabla)
    End If
    Rs.Close
    Set Rs = Nothing
    If Not MantenerAbierto Then
      CierraCnx
    End If
  End If
  
  Set ObtieneValores = tbla

End Function

Private Function QuitaEspacios(ByVal CualCadena As String) As String
  ' Var
  Dim i As Integer
  
  Do
    i = InStr(CualCadena, " ")
    If i > 0 Then
      CualCadena = Left(CualCadena, i - 1) & Mid(CualCadena, i + 1)
    End If
  Loop Until i = 0
  
  If Right(CualCadena, 1) <> "|" Then
    CualCadena = CualCadena & "|"
  End If
  
  QuitaEspacios = CualCadena
End Function

Public Property Get TablaClaves() As String
  TablaClaves = strTblClaves
End Property

Public Property Let TablaClaves(ByVal vNewValue As String)
  strTblClaves = vNewValue
End Property

Public Property Get TablaColas() As String
  TablaColas = strTblColas
End Property

Public Property Let TablaColas(ByVal vNewValue As String)
  strTblColas = vNewValue
End Property

Public Property Get TablaSecuencias() As String
  TablaSecuencias = strTblSecuencias
End Property

Public Property Let TablaSecuencias(ByVal vNewValue As String)
  strTblSecuencias = vNewValue
End Property

Public Property Get Usuario() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Usuario
    Usuario = mvarUsuario
End Property

Public Property Let Usuario(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Usuario = 5
    mvarUsuario = vData
End Property


'---------------------------------------------------------------------------------------
' Procedimiento : GeneraBD
' Fecha y hora  : 02/09/2003 06:33
' Autor         : adgarza
' Propósito     :
'---------------------------------------------------------------------------------------
'
Public Function GeneraBD(EstructuraOrigen As String, BaseDestino As String) As Integer
  Dim intResult As Integer


  GeneraBD = intResult
End Function


'Purpose     :  Compact a JET (Access) database using ADO
'Inputs      :  strRutaBD                   The path to the database path eg. C:\nwind.mdb
'               [bolCifrarBD]              If True, encrypts the contents of the database
'Outputs     :  Returns zero if successful, else returns error code
'Author      :  Andrew Baker
'Date        :  15/01/2001 11:55
'Notes       :  Requires "Microsoft Jet and Replication Objects X.X library",
'               where (X.X is greater than or equal to 2.1)
'               Compacts the database by creating a temporary database with the extension .tmp then,
'               if the compaction is successful, it overwrites the original database.
'               Will not work if anyone else is connected to the database.
'Revisions   :
'Assumptions :

Function CompactarBaseDatosJET(strRutaBD As String, Optional bolCifrarBD As Boolean = False) As Long
    Dim oJRO As Object 'JRO.JetEngine

    On Error GoTo Falla
    
    If Len(Dir$(strRutaBD & ".tmp")) Then
        ' Suprime la actual base de datos temporal
        Kill strRutaBD & ".tmp"
    End If
    
    Set oJRO = CreateObject("JRO.JetEngine")
    
    If bolCifrarBD Then
        ' Compacta y cifra la base de datos
        oJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRutaBD, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRutaBD & ".tmp;Jet OLEDB:Encrypt Database=True"
    Else
        ' Compacta la base de datos
        oJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRutaBD, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRutaBD & ".tmp;Jet OLEDB:Engine Type=4"
    End If
    
    ' Suprime la base de datos actual
    VBA.Kill strRutaBD
    ' Cambia el nombre de la base de datos compactada
    Name strRutaBD & ".tmp" As strRutaBD
    Set oJRO = Nothing

    Exit Function

Falla:
    Debug.Print "No se pudo compactar la base de datos: " & Err.Description
    CompactarBaseDatosJET = Err.Number
    Set oJRO = Nothing
    On Error GoTo 0
End Function

Public Function ActualizarEstructura(CualBDModelo As String, CualBDResultante As String, Optional ContraseniaModelo As String = "", Optional ContraseniaResultante As String = "") As Boolean
    Dim objActualizaEstructura As New cUpdateADO  ' Defina al objeto
    Dim cnnBDResultante As New ADODB.Connection      ' Conexión a la base de datos de la aplicación
    Dim cnnBDModelo As New ADODB.Connection       ' Conexión con la base de datos origen (o modelo)
    Dim sTemp As String                     'used for results message
    Dim dStart As Date                      'process started
    Dim dEnd   As Date                      'process ended
    Dim iFN As Integer                      'open update log
    Dim sRec As String                      'read log
    Dim bAnyChanges As Boolean              '2003/05/30 Any changes
    Dim strNombreBaseResultante As String
    Dim mcsProvider As String
    
    mcsProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Screen.MousePointer = vbHourglass
    On Error GoTo DoUpdateEH                'in case it crashes
    dStart = Now
    Set cnnBDResultante = Nothing
    Set cnnBDModelo = Nothing
    
    ' ¿Existe la base de datos? Si no, la creamos.
    strNombreBaseResultante = getObtieneLaBase(CualBDResultante)
    If Dir(strNombreBaseResultante) = "" Then
      CreaBaseAccess strNombreBaseResultante
    End If
    
    ' Ahora, actualizaremos la estructura de acuerdo con el modelo
    With objActualizaEstructura
        If ContraseniaModelo <> "" Then
          cnnBDModelo.Open mcsProvider & Trim$(CualBDModelo) & ";Jet OLEDB:Database Password=" & ContraseniaModelo
        Else
          cnnBDModelo.Open mcsProvider & Trim$(CualBDModelo)
        End If
        
        If ContraseniaResultante <> "" Then
          cnnBDResultante.Open mcsProvider & Trim$(CualBDResultante) & ";Jet OLEDB:Database Password=" & ContraseniaResultante
        Else
          cnnBDResultante.Open mcsProvider & Trim$(CualBDResultante)
        End If
        
        If .UpdateTables(cnnBDResultante, cnnBDModelo) Then
            bAnyChanges = True
        End If
        If .UpdateIndexes() Then
            bAnyChanges = True
        End If
        If .UpdateStoredProceduresAndViews() Then
            bAnyChanges = True
        End If
        If .UpdateRelationships() Then
            bAnyChanges = True
        End If
                
        If bAnyChanges Then
            dEnd = Now
            sTemp = ""
            ' sTemp = sTemp & FormatSecondsDiff(dStart, dEnd) 'how long did it take.
            ' Me.Tag = sTemp
        Else
            Me.Tag = "No se hicieron cambios."
        End If

        ActualizarEstructura = objActualizaEstructura.bAnyErrors
    
    End With
    GoTo DoUpdateExit:
DoUpdateEH:
    Me.Tag = "Error: " & Err.Number & ", " & Err.Description & vbCrLf & vbCrLf & "Tal vez no sea una base de datos de Access."
DoUpdateExit:
    Set objActualizaEstructura = Nothing
    Screen.MousePointer = vbNormal          'reset mouse pointer
End Function

Public Function CreaBaseAccess(Optional CualBase As String = "", Optional ForzarCreacion As Boolean = False, Optional Contrasenia As String = "") As String
  ' Var
  Dim catBaseAccess As New ADOX.Catalog

  On Error Resume Next
  Err.Clear
  
  If CualBase = "" Or ForzarCreacion Then
    CualBase = GenerarCadenaConexion(CualBase)
    If CualBase <> "" Then
      ' Obtiene sólo la cadena con Data Source
      CualBase = getObtieneLaBase(CualBase)
    End If
  End If
  
  If CualBase <> "" Then
    If Contrasenia = "" Then
      catBaseAccess.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CualBase
    Else
      catBaseAccess.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CualBase & ";Jet OLEDB:Database Password=" & Contrasenia
    End If
  Else
    Err.Raise 75
  End If
  Set catBaseAccess = Nothing
  
  If Err.Number <> 0 Then
    CreaBaseAccess = ""
  Else
    CreaBaseAccess = CualBase
  End If
  On Error GoTo 0
End Function

' *************************************************
' * Programmer Name  : Leo Koach
' * Web Site         : http://www.vbonly.com
' * Date             : 8/7/2000
' **************************************************************
' * Comments         : display the ADO Connection Window.
' *(ADO DB Designer).  Pass the function a connection string
' * and the user will be able to edit the connection properties
' * via this window.  The function will then return the modified
' * connection string.
 '*
 '* References to Microsoft Active Data Objects and the
 '* the Microsoft OLE DB Service Component 1.0 Type Library
 '* are required.
' **************************************************************

Public Function GenerarCadenaConexion(ByVal CadenaConexion As String) As String
   Dim cnTmp            As ADODB.Connection
   Dim dlViewConnection As MSDASC.DataLinks

   On Error Resume Next

   If Not (Trim$(CadenaConexion) = "") Then
      Set cnTmp = New ADODB.Connection
      cnTmp.ConnectionString = CadenaConexion
      Set dlViewConnection = New MSDASC.DataLinks
      
      If dlViewConnection.PromptEdit(cnTmp) Then
         If Err.Number = 0 Then
           GenerarCadenaConexion = cnTmp.ConnectionString
         Else
           GenerarCadenaConexion = CadenaConexion
         End If
      Else
         GenerarCadenaConexion = CadenaConexion
      End If
      Set dlViewConnection = Nothing
      Set cnTmp = Nothing
   Else
      Set dlViewConnection = New MSDASC.DataLinks
      Set cnTmp = dlViewConnection.PromptNew
      If Err.Number = 0 Then
        GenerarCadenaConexion = cnTmp.ConnectionString
      Else
        GenerarCadenaConexion = ""
      End If
      Set dlViewConnection = Nothing
      Set cnTmp = Nothing
   End If
End Function

Private Function getObtieneLaBase(CualConexion As String) As String
  ' Var
  Dim i As Integer
  Dim l As Integer
  Dim strCualBase As String
  
  i = InStr(CualConexion, "Data Source")
  If i > 0 Then
    l = InStr(i, CualConexion, ";")
    If l = 0 Then
      l = Len(CualConexion) + 1
    End If
    i = InStr(i, CualConexion, "=")
    If i > 0 Then
      strCualBase = Mid(CualConexion, i + 1, l - (i + 1))
    Else
      strCualBase = ""
    End If
  Else
    strCualBase = ""
  End If

  getObtieneLaBase = strCualBase
End Function

Private Function FormatSecondsDiff(dStart As Date, dEnd As Date) As String
' Report the number of seconds that the process took.
    Dim dDiff As Date
    Dim fSeconds As Double
    dDiff = dEnd - dStart
    fSeconds = (24# * 60# * 60#) * dDiff
    FormatSecondsDiff = "  " & Format$(fSeconds, "###,##0.00") & vbTab & "seconds."
End Function

Public Function ObtieneOrdenTablasPorActualizar(CualBDModelo As String, Optional TablaBase As String = "OrdenTablas", Optional ContraseniaModelo As String = "", Optional UsuarioModelo As String = "") As Collection
  ' Var
  Dim colOrdenTablas As Collection
  Dim cnnBDModelo As New ADODB.Connection       ' Conexión con la base de datos origen (o modelo)
  Dim cat As ADOX.Catalog
  Dim tbl As ADOX.Table
  Dim strSQL As String
  Dim rsT As ADODB.Recordset
  
  ' Devuelve el orden en que deben actualizarse las tablas
  If Len(CualBDModelo) > 0 Then
    Set cat = New ADOX.Catalog
    If ContraseniaModelo Then
      If InStr(CualBDModelo, "Microsoft.Jet.OLEDB") Then
        cat.ActiveConnection = Trim$(CualBDModelo) & ";Jet OLEDB:Database Password=" & Contrasenia
      Else
        cat.ActiveConnection = Trim$(CualBDModelo)
      End If
    Else
      cat.ActiveConnection = Trim$(CualBDModelo)
    End If
    
    ' El orden de actualización se obtiene de una tabla que debe encontrarse
    '  en la base de datos modelo llamada OrdenTablas. Si no se encuentra esta
    '  tabla, intenta hacer la actualización en orden alfabético.
    
    ' Determina si la tabla OrdenTablas se encuentra
    On Error Resume Next
    Set tbl = cat.Tables(TablaBase)
    On Error GoTo 0
    If tbl Is Nothing Then
      For Each tbl In cat.Tables
        colOrdenTablas.Add tbl.Name
      Next tbl
      Set tbl = Nothing
    Else
      If ContraseniaModelo <> "" Then
        If InStr(CualBDModelo, "Microsoft.Jet.OLEDB") Then
          cnnBDModelo.Open Trim$(CualBDModelo) & ";Jet OLEDB:Database Password=" & Contrasenia
        Else
          cnnBDModelo.Open Trim$(CualBDModelo), UsuarioModelo, Contrasenia
        End If
      Else
        cnnBDModelo.Open Trim$(CualBDModelo)
      End If
      rsT.Open "SELECT * FROM " & TablaBase & " ORDER BY IDOrden;", cnnBDModelo
      
      If Not rsT.BOF And Not rsT.EOF Then
        Do While Not rsT.EOF
          colOrdenTablas.Add rsT.Fields("Tabla").Value
        Loop
      End If
      rsT.Close
      Set rsT = Nothing
    End If
  
    ' Libera recursos
    Set cat = Nothing
  End If
    
  Set ObtieneOrdenTablasPorActualizar = colOrdenTablas
End Function

Public Function ActualizarDatos(CualBDModelo As String, CualBDResultante As String, Optional ContraseniaModelo As String = "", Optional ContraseniaResultante As String = "") As Boolean
    ' Var
    Dim cnnBDResultante As New ADODB.Connection   ' Conexión a la base de datos de la aplicación
    Dim cnnBDModelo As New ADODB.Connection       ' Conexión con la base de datos origen (o modelo)
    Dim mcsProvider As String
    Dim rsR As Recordset
    Dim rsM As Recordset
    Dim rsT As Recordset
    Dim strSQL As String
    Dim fld As Field
    Dim strCvePrincipal As String
    
    mcsProvider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Screen.MousePointer = vbHourglass
    ' On Error GoTo DoUpdateEH                'in case it crashes
    Set cnnBDResultante = Nothing
    Set cnnBDModelo = Nothing
    
  
    ' Ahora, actualizaremos los datos con base en el modelo
    If ContraseniaModelo <> "" Then
      cnnBDModelo.Open mcsProvider & Trim$(CualBDModelo) & ";Jet OLEDB:Database Password=" & ContraseniaModelo
    Else
      cnnBDModelo.Open mcsProvider & Trim$(CualBDModelo)
    End If
        
    If ContraseniaResultante <> "" Then
      cnnBDResultante.Open mcsProvider & Trim$(CualBDResultante) & ";Jet OLEDB:Database Password=" & ContraseniaResultante
    Else
      cnnBDResultante.Open mcsProvider & Trim$(CualBDResultante)
    End If
    
    ' Obtiene la secuencia de tablas
    strSQL = "SELECT * FROM Tablas ORDER BY IDTabla;"
    Set rsT = cnnBDModelo.Execute(strSQL)
    
    ' Actualiza los datos
    If Not rsT.BOF And Not rsT.EOF Then
      rsT.MoveFirst
      Do While Not rsT.EOF
        strSQL = "SELECT * FROM " & rsT("Tabla").Value & ";"
        Set rsM = cnnBDModelo.Execute(strSQL)
        Set rsR = New ADODB.Recordset
        rsR.CursorType = adOpenKeyset
        rsR.LockType = adLockOptimistic
        rsR.Open strSQL, cnnBDResultante
        If Not rsM.BOF And Not rsM.EOF Then
          rsM.MoveFirst
          Do While Not rsM.EOF
            rsR.AddNew
            For Each fld In rsM.Fields
              rsR(fld.Name).Value = fld.Value
            Next fld
            On Error Resume Next
            rsR.Update
            If Err.Number = -2147217887 Then
              rsR.CancelUpdate
              Err.Clear
              
              ' Aquí tendría que actualizar los datos capturados
              'strCvePrincipal = ""
              'strCvePrincipal = Me.ObtieneCP(rsM("Tabla").Value)
              'strSQL = Me.ObtieneC
                
              'rsR.Find
              'For Each fld In rsM.Fields
              '  rsR(fld.Name).Value = fld.Value
              'Next fld
              'On Error Resume Next
            'rsR.Update
              
            ElseIf Err.Number <> 0 Then
              MsgBox Err.Description, vbExclamation, "Error " & Err.Number
              rsR.CancelUpdate
              Err.Clear
            End If
            On Error GoTo 0
            rsM.MoveNext
          Loop
          rsR.Close
          Set rsR = Nothing
        End If
        rsT.MoveNext
      Loop
    End If
    Screen.MousePointer = vbDefault
    
    ' He aquí algo que hay que modificar
    ActualizarDatos = True
End Function
