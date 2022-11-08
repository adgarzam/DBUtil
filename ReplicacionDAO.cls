VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReplicacionDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'Public Function VerificaBase(ByVal BaseDeDatos As String, ByVal Prototipo As String, ByRef OrdenTablas() As String) As Boolean
'  'Rutina principal de todo el modulo de verificación de estructura
'  'de una BaseDeDatos contra el prototipo seleccionado
'  On Error GoTo Error
'  Dim Resp As Integer
'  Dim Resp2 As Integer
'  Dim VersionPrototipo As Date
'  Dim VersionMayor As Long
'  Dim Db As DAO.Database
'  Dim rs As DAO.Recordset
'  VerificaBase = False
'  Resp = VerificaEstructura(BaseDeDatos, Prototipo, VersionPrototipo, VersionMayor)
'  If Resp = 2 Then Resp = 3
'  Select Case Resp
'    Case 4        '4 Error grave
'      VerificaBase = False
'    Case 1        '1 se verifico OK
'      VerificaBase = True
'    Case 2        '2 error corregible
'      Resp2 = MsgBox("La estructura de la BaseDeDatos no corresponde a la que está en uso. ¿Desea actualizarla?", vbYesNo + vbInformation, "Neodata")
'      If Resp2 = vbYes Then
'        Resp = ActualizaEstructura(BaseDeDatos, Prototipo)
'        If Resp Then
'          'una vez verificada la BaseDeDatos actualiza el numero de version
'          Set Db = OpenDatabase(BaseDeDatos, False, False)
'          Set rs = Db.OpenRecordset("version", dbOpenDynaset)
'          If rs.RecordCount = 0 Then
'            rs.AddNew
'          Else
'            rs.Edit
'          End If
'          rs!Version = VersionPrototipo
'          rs!VersionMayor = VersionMayor
'          rs.Update
'          rs.Close
'          ImportaParametros Db, Prototipo
'          Db.Close
'          '
'          MsgBox "La estructura se actualizó satisfactoriamente.", vbInformation, "Neodata"
'          VerificaBase = True
'        End If
'      End If
'    Case 3        '3 se necesita crear una BaseDeDatos nueva e importar los
'                  'datos a ella para corregirla
'      Resp2 = MsgBox("La estructura de la BaseDeDatos no corresponde a la que está en uso. ¿Desea actualizarla?", vbYesNo + vbInformation, "Neodata")
'      If Resp2 = vbYes Then
'        Resp = ImportaBaseDeDatos(BaseDeDatos, Prototipo, OrdenTablas())
'        If Resp Then
'          'una vez verificada la BaseDeDatos actualiza el numero de version
'          Set Db = OpenDatabase(BaseDeDatos, False, False)
'          Set rs = Db.OpenRecordset("version", dbOpenDynaset)
'          If rs.RecordCount = 0 Then
'            rs.AddNew
'          Else
'            rs.Edit
'          End If
'          rs!Version = VersionPrototipo
'          rs!VersionMayor = VersionMayor
'          rs.Update
'          rs.Close
'          ImportaParametros Db, Prototipo
'          Db.Close
'          '
'          MsgBox "La estructura se actualizó satisfactoriamente.", vbInformation, "Neodata"
'          VerificaBase = True
'        End If
'      End If
'  End Select
'
'  Exit Function
'Error:
'  VerificaBase = False
'End Function
'
'Public Function VerificaEstructura(ByVal Archivo As String, ByVal Prototipo As String, Optional ByRef VersionPrototipo As Date, Optional ByRef VersionMayor As Long) As Integer
'  On Error GoTo ErrorVerificaEstructura
'  '1 se verifico OK
'  '2 error corregible
'  '3 se necesita crear una BaseDeDatos nueva e importar los datos a ella para corregirla
'  '4 error grave
'  Dim Resp As Integer
'  Dim rs As DAO.Recordset
'  Dim Tabla  As DAO.TableDef
'  Dim Campo As DAO.Field
'  Dim Campo2 As DAO.Field
'  Dim Query As DAO.QueryDef
'  Dim Db As DAO.Database
'  Dim DbPrototipo As DAO.Database
'  Dim Existe As New ExisteArchivo
'  'establece el valor default
'  VerificaEstructura = 1
'  'verifica que existan los archivos
'  If Not Existe.ExisteArchivo(Archivo) Or Not Existe.ExisteArchivo(Prototipo) Then
'    'si no existe alguno de los archivos no hace nada
'    Exit Function
'  End If
'  'Establece el reloj de arena mientras dura el proceso
'  Screen.MousePointer = vbHourglass
'  'abre la base de datos a comparar contra la base prototipo y abre esta última
'  Set DbPrototipo = OpenDatabase(Prototipo, False, True)
'  Set Db = OpenDatabase(Archivo, False, True)
'  'verifica el campo Version
'  If Not ExisteTabla("version", DbPrototipo) Then
'    'si no existe la tabla version en la base de datos nuevo.neo
'    MsgBox "El no. de versión de la base de datos no existe en el prototipo. Comuníquese con el área de Asistencia técnica de Neodata.", vbCritical, "Neodata"
'    VerificaEstructura = 4
'    GoTo Salida
'  End If
'
'  Set rs = DbPrototipo.OpenRecordset("version", dbOpenDynaset)
'  VersionPrototipo = rs!Version
'  VersionMayor = rs!VersionMayor
'
'  If ExisteTabla("version", Db) Then
'    If ExisteCampo("versionmayor", "version", Db, Campo2) And ExisteCampo("version", "version", Db, Campo2) Then
'      Set rs = Db.OpenRecordset("version", dbOpenDynaset)
'      If rs.RecordCount > 0 Then
'        If rs!VersionMayor <> VersionMayor Then
'          VerificaEstructura = 3
'          GoTo Salida
'        End If
'        If rs!Version = VersionPrototipo Then
'          GoTo Salida
'        Else
'          VerificaEstructura = 2
'        End If
'      Else
'        VerificaEstructura = 3
'        GoTo Salida
'      End If
'    Else
'      VerificaEstructura = 3
'      GoTo Salida
'    End If
'  Else
'    VerificaEstructura = 3
'    GoTo Salida
'  End If
'
'
'  '1o. verifica que todos los querys de la base de datos existan y estén correctos
'  For Each Query In DbPrototipo.QueryDefs
'    If Not ExisteQuery(Query.Name, Db) Then
'      VerificaEstructura = 2
'    Else
'      If Query.Sql <> TextoSql(Query.Name, Db) Then
'        VerificaEstructura = 2
'      End If
'    End If
'  Next Query
'  'termina de verificar los querys creados en la base de datos
'
'  '2o. verifica las relaciones de la base de datos
'  Dim RelOrigen As DAO.Relation
'  Dim RelDestino As DAO.Relation
'  'Barre todas las relaciones de la base prototipo
'  'verificando que existan en la base destino
'  For Each RelOrigen In DbPrototipo.Relations
'    If Not ExisteRelacion(RelOrigen.Name, Db, RelDestino) Then
'      VerificaEstructura = 2
'    Else
'      'como existe verifica sus propiedades que sean iguales
'      If Not ChecaDetalleRelacion(RelOrigen, RelDestino) Then VerificaEstructura = 2
'    End If
'  Next RelOrigen
'  Set RelOrigen = Nothing
'  Set RelDestino = Nothing
'
'
'  '3o. verifica que todas las tablas de la base prototipo existan en la base a checar
'  For Each Tabla In DbPrototipo.TableDefs
'    If Left(LCase(Tabla.Name), 4) <> "msys" Then
'      If Not ExisteTabla(Tabla.Name, Db) Then
'        VerificaEstructura = 2
'      Else
'        '3o. verifica campo a campo que sean iguales
'        For Each Campo In Tabla.Fields
'          If Not ExisteCampo(Campo.Name, Tabla.Name, Db, Campo2) Then
'            VerificaEstructura = 2
'          Else
'            'si se llaman igual checa los atributos del campo que sean iguales
'            If Campo.Type <> Campo2.Type Then
'              VerificaEstructura = 3
'              GoTo Salida
'            End If
'            If Campo.Size <> Campo2.Size Then
'              VerificaEstructura = 3
'              GoTo Salida
'            End If
'            If Campo.Required <> Campo2.Required Then
'              VerificaEstructura = 3
'              GoTo Salida
'            End If
'            If Campo.DefaultValue <> Campo2.DefaultValue Then
'              VerificaEstructura = 3
'              GoTo Salida
'            End If
'            If Campo.AllowZeroLength <> Campo2.AllowZeroLength Then
'              VerificaEstructura = 3
'              GoTo Salida
'            End If
'          End If
'        Next Campo
'      End If
'    End If
'  Next Tabla
'Salida:
'  'Este segmento de codigo se utiliza para salir de la función
'  'restaurando los valores necesarios y destruyendo objetos
'  DbPrototipo.Close
'  Db.Close
'  Set DbPrototipo = Nothing
'  Set Db = Nothing
'  Screen.MousePointer = vbDefault
'  Exit Function
'ErrorVerificaEstructura:
'  Screen.MousePointer = vbDefault
'  VerificaEstructura = 3
'  Err.Clear
'  GoTo Salida
'End Function
''
'Public Function ImportaBaseDeDatos(ByVal Archivo As String, ByVal Prototipo As String, ByRef OrdenTablas() As String) As Boolean
'  'con esta rutina se crea una BaseDeDatos nueva con nombre x
'  'se importan los datos a ella
'  'y posteriormente se borra la anterior y se copia la nueva con el nombre anterior
'  On Error GoTo ErrorImportaBaseDeDatos
'  Dim i As Integer
'  Dim Respx As Integer
'  Dim rs As DAO.Recordset
'  Dim Rs2 As DAO.Recordset
'  Dim Tabla  As DAO.TableDef
'  Dim Campo As DAO.Field
'  Dim Campo2 As DAO.Field
'  Dim DbVieja As DAO.Database
'  Dim Dbnueva As DAO.Database
'  Dim Archivotmp As String      'genera un nombre de archivo temporal
'  'Establece el reloj de arena mientras dura el proceso
'  Screen.MousePointer = vbHourglass
'  ImportaBaseDeDatos = False
'  'copia la base de datos Nuevo.neo a un archivo temporal
'  Archivotmp = AsignaPas()
'  FileCopy Prototipo, Archivotmp ' Copy source to target.
'  Set DbVieja = OpenDatabase(Archivo, True, False)        'abre la base con estructura vieja
'  Set Dbnueva = OpenDatabase(Archivotmp, True, False) 'abre la base de datos temporal
'  'Primero borra todas las relaciones que existan para que no marque error
'  'borraRelaciones Dbnueva
'  'Copia todas las tablas de la base con estructura vieja a la BaseDeDatos nueva
'  i = 0
'  Dim LimpiaBase As clsLimpiaBase
'  Set LimpiaBase = New clsLimpiaBase
'  LimpiaBase.LimpiaBase Dbnueva
'  Set LimpiaBase = Nothing
'  For i = 1 To UBound(OrdenTablas())
'    If ExisteTabla(OrdenTablas(i), DbVieja) Then
'      'si la tabla existe en la BaseDeDatos con estructura vieja trata de importar datos
'        If Not ImportaTabla(Dbnueva, OrdenTablas(i), DbVieja.Name) Then
'          Set rs = DbVieja.OpenRecordset(OrdenTablas(i), dbOpenDynaset)
'          Set Rs2 = Dbnueva.OpenRecordset(OrdenTablas(i), dbOpenDynaset)
'          If rs.RecordCount > 0 Then
'            rs.MoveFirst
'            Do While Not rs.EOF
'              'procesa todos los campos de la tabla de la BaseDeDatos vieja
'              Rs2.AddNew
'              For Each Campo In rs.Fields
'                If ExisteCampo(Campo.Name, OrdenTablas(i), Dbnueva, Campo2) Then
'                  If Not IsNull(Campo.Value) Then Rs2.Fields(Campo.Name).Value = Campo.Value
'                End If
'              Next Campo
'              Rs2.Update
'              rs.MoveNext
'            Loop
'          End If
'          rs.Close
'          Rs2.Close
'        End If
'
'    End If
'  Next i
'
'  'copia de nuevo las relaciones que necesita
'  'DbVieja.Close
'  'Set DbVieja = OpenDatabase(Prototipo, False, True) 'abre la base con estructura prototipo
'  'VerificaRelaciones DbVieja, Dbnueva
'  DbVieja.Close
'  Dbnueva.Close
'  Set DbVieja = Nothing
'  Set Dbnueva = Nothing
'  'borra la BaseDeDatos vieja
'  Kill Archivo
'  'Archivo = Archivo & ".mdb"
'  Name Archivotmp As Archivo    ' Move and rename file.
'  ImportaBaseDeDatos = True
'  Screen.MousePointer = vbDefault
'  Exit Function
'ErrorImportaBaseDeDatos:
'  Screen.MousePointer = vbDefault
'  ImportaBaseDeDatos = False
'  Select Case Err.Number
'    Case 3163   'the field is too small to accept the amount of data
'      If Campo.Type = dbText Then
'        If Not IsNull(Campo.Value) Then
'          Rs2.Fields(Campo.Name).Value = Left(Campo.Value, Rs2.Fields(Campo.Name).Size)
'        End If
'      End If
'      Resume Next
'    Case 3421     'Data type conversion error
'      Resume Next
'    Case 3356
'      MsgBox "La base de datos se está utilizando y no puede ser actualizada.", vbInformation, "Neodata"
'      If Not DbVieja Is Nothing Then DbVieja.Close
'      If Not Dbnueva Is Nothing Then Dbnueva.Close
'      Set DbVieja = Nothing
'      Set Dbnueva = Nothing
'      Exit Function
'    Case Else
'      Respx = MsgBox("ImportaBaseDeDatos: Error no. " & Err.Number & "  " & Err.Description, vbAbortRetryIgnore, "Neodata")
'      If Respx = vbAbort Then
'        DbVieja.Close
'        Dbnueva.Close
'        Set DbVieja = Nothing
'        Set Dbnueva = Nothing
'        Exit Function
'      End If
'      If Respx = vbRetry Then
'        Resume
'      End If
'      If Respx = vbIgnore Then
'        Resume Next
'      End If
'  End Select
'End Function
'
'Private Function ImportaTabla(ByRef Dbnueva As DAO.Database, ByRef NombreTabla As String, ByRef DbVieja As String) As Boolean
'  'Con esta rutina se importan los datos de la tabla siempre y cuando
'  'Sea posible de lo contrario se genera un error
'  On Error GoTo ErrorImportaTabla
'  Dim Sql As String
'  Sql = " INSERT INTO " & NombreTabla & " SELECT * FROM [" & DbVieja & "]." & NombreTabla & ";"
'  Dbnueva.Execute Sql
'  ImportaTabla = True
'  Exit Function
'ErrorImportaTabla:
'  Err.Clear
'  ImportaTabla = False
'End Function
'
'Private Function AsignaPas() As String
'  'Esta rutina sirve para crear nombres
'  'unicos para archivos temporales
'  Dim Paso As Single
'  Dim paso2 As String
'  Paso = Timer
'  Paso = Paso * 100   'para quitarle los decimales
'  paso2 = "$" & Paso
'  AsignaPas = paso2
'End Function
'
'Private Function ActualizaEstructura(ByVal Archivo As String, ByVal Prototipo As String) As Boolean
'  On Error GoTo ErrorActualizaEstructura
'  Dim Resp As Integer
'  Dim Query As DAO.QueryDef
'  Dim QueryNuevo As DAO.QueryDef
'  Dim Tabla  As DAO.TableDef
'  Dim Campo As DAO.Field
'  Dim Campo2 As DAO.Field
'  Dim Db As DAO.Database
'  Dim DbNuevo As DAO.Database
'  'Establece el reloj de arena mientras dura el proceso
'  Screen.MousePointer = vbHourglass
'  ActualizaEstructura = False
'  'abre la base de datos a comparar contra la base nuevo.neo y abre esta última
'  Set DbNuevo = OpenDatabase(Prototipo, , True)
'  Set Db = OpenDatabase(Archivo, True)
'
'  '1o. verifica que todos los querys de la base de datos existan y estén correctos
'  For Each Query In DbNuevo.QueryDefs
'    If Not ExisteQuery(Query.Name, Db) Then
'      'agrega el query si no existe en la base de datos actual
'      Set QueryNuevo = Db.CreateQueryDef(Query.Name, Query.Sql)
'      'Db.QueryDefs.Append QueryNuevo
'    Else
'      'si el query ya existe verifica que su SQL sea el mismo
'      Set QueryNuevo = Db.QueryDefs(Query.Name)
'      If Query.Sql <> QueryNuevo.Sql Then
'        Db.QueryDefs(Query.Name).Sql = Query.Sql
'        'Db.QueryDefs.Refresh
'      End If
'    End If
'  Next Query
'  'termina de verificar los querys creados en la base de datos
'
'  '2o. verifica que todas las tablas de la base "nuevo.neo" existan en la base a checar
'  For Each Tabla In DbNuevo.TableDefs
'    If Left(LCase(Tabla.Name), 4) <> "msys" Then  'omite las tablas del sistema
'      If Not ExisteTabla(Tabla.Name, Db) Then
'        AgregaTabla Tabla, Db
'      Else
'        'verifica campo a campo que sean iguales
'        For Each Campo In Tabla.Fields
'          If Not ExisteCampo(Campo.Name, Tabla.Name, Db, Campo2) Then
'            AgregaCampo Campo, Db.TableDefs(Tabla.Name)
'          Else
'            'si se llaman igual checa los atributos del campo que sean iguales
'            'por el momento no hará nada si los atributos son diferentes
'            'por que no se sabe como resolver este problema
'          End If
'        Next Campo
'      End If
'    End If
'  Next Tabla
'
'  '3o. verifica las relaciones para agregar las que no existan
'  'y para borrar las que estén de más
'  VerificaRelaciones DbNuevo, Db
'  '
'  ActualizaEstructura = True
'  Screen.MousePointer = vbDefault
'  Exit Function
'ErrorActualizaEstructura:
'  Screen.MousePointer = vbDefault
'  ActualizaEstructura = False
'  Select Case Err.Number
'    Case 3356
'      MsgBox "La base de datos se está utilizando y no puede ser actualizada.", vbInformation, "Neodata"
'    Case Else
'      MsgBox "ActualizaEstructura.Error no.: " & Err.Number & " : " & Err.Description, , "Neodata"
'  End Select
'End Function
'
'Private Sub AgregaCampo(ByRef Campo As DAO.Field, ByRef Tabla As DAO.TableDef)
'  Dim Campo2 As DAO.Field
'  Set Campo2 = Tabla.CreateField(Campo.Name, Campo.Type, Campo.Size)
'  If Campo.Type = dbText Or Campo.Type = dbMemo Then
'    Campo2.AllowZeroLength = Campo.AllowZeroLength
'  End If
'  Campo2.DefaultValue = Campo.DefaultValue
'  Campo2.Required = Campo.Required
'  Tabla.Fields.Append Campo2
'End Sub
'
'Private Sub AgregaTabla(ByRef Tabla As DAO.TableDef, ByRef Db As DAO.Database)
'  Dim td As DAO.TableDef
'  Dim Campo As DAO.Field
'  'agrega la tabla
'  Set td = Db.CreateTableDef(Tabla.Name)
'    'agrega los campos
'    For Each Campo In Tabla.Fields
'      AgregaCampo Campo, td
'    Next Campo
'  Db.TableDefs.Append td
'End Sub
'
'Private Function ExisteCampo(ByVal NombreCampo As String, ByVal NombreTabla As String, ByRef Base As DAO.Database, ByRef Campo2 As DAO.Field) As Boolean
'  Dim Campo As DAO.Field
'  Dim Tabla As DAO.TableDef
'  ExisteCampo = False
'  For Each Tabla In Base.TableDefs
'    If LCase(NombreTabla) = LCase(Tabla.Name) Then
'      For Each Campo In Tabla.Fields
'        If LCase(NombreCampo) = LCase(Campo.Name) Then
'          Set Campo2 = Campo
'          ExisteCampo = True
'          Exit Function
'        End If
'      Next Campo
'    End If
'  Next Tabla
'End Function
'
'Private Function TextoSql(ByVal NombreQuery, ByRef Base As DAO.Database) As String
'  Dim Query As DAO.QueryDef
'  TextoSql = ""
'  For Each Query In Base.QueryDefs
'    If LCase(NombreQuery) = LCase(Query.Name) Then
'      TextoSql = Query.Sql
'      Exit Function
'    End If
'  Next Query
'End Function
'
'Private Function ExisteQuery(ByVal NombreQuery As String, ByRef Base As DAO.Database) As Boolean
'  Dim Query As DAO.QueryDef
'  ExisteQuery = False
'  For Each Query In Base.QueryDefs
'    If LCase(NombreQuery) = LCase(Query.Name) Then
'      ExisteQuery = True
'      Exit Function
'    End If
'  Next Query
'End Function
'
'Public Function ExisteTabla(ByVal NombreTabla As String, ByRef Base As DAO.Database) As Boolean
'  Dim Tabla As DAO.TableDef
'  ExisteTabla = False
'  For Each Tabla In Base.TableDefs
'    If LCase(NombreTabla) = LCase(Tabla.Name) Then
'      ExisteTabla = True
'      Exit Function
'    End If
'  Next Tabla
'End Function
'
'Private Sub borraRelaciones(ByRef Dbnueva As DAO.Database)
'  'borra todas las relaciones que existan para que no marque error
'  'al agregar datos indiscriminadamente
'  Dim Nombre As String
'  Dim i As Integer
'  'For Each Relacion In Dbnueva.Relations
'  For i = Dbnueva.Relations.Count To 1 Step -1
'    Nombre = Dbnueva.Relations(i - 1).Name
'    Dbnueva.Relations.Delete Nombre
'    Dbnueva.Relations.Refresh
'  Next i
'End Sub
'
'Private Sub Borraunarelacion(ByRef Base As DAO.Database, ByRef RelNombre As String)
'  Dim Nombre As String
'  Dim i As Integer
'  'For Each Relacion In Dbnueva.Relations
'  For i = Base.Relations.Count To 1 Step -1
'    Nombre = Base.Relations(i - 1).Name
'    If LCase(Nombre) = LCase(RelNombre) Then
'      Base.Relations.Delete Nombre
'      Base.Relations.Refresh
'      Exit Sub
'    End If
'  Next i
'End Sub
'
'Private Sub VerificaRelaciones(ByRef DbOrigen As DAO.Database, ByRef DbDestino As DAO.Database)
'  On Error GoTo Error
'  Dim RelOrigen As DAO.Relation
'  Dim RelDestino As DAO.Relation
'  Dim i As Integer
'  '1o. borra todas las relaciones de la base a modificar
'  For i = DbDestino.Relations.Count To 1 Step -1
'    DbDestino.Relations.Delete DbDestino.Relations(i - 1).Name
'  Next i
'  '
'  '2o. barre todas las relaciones de la base prototipo
'  'verificando que existan en la base destino
'  For Each RelOrigen In DbOrigen.Relations
'    If Not ExisteRelacion(RelOrigen.Name, DbDestino, RelDestino) Then
'      'como no existe la agrega tomando todas las propiedades de la relacion origen
'      CopiaRelacion DbDestino, RelOrigen
'    Else
'      'como existe verifica sus propiedades que sean iguales
'        If Not ChecaDetalleRelacion(RelOrigen, RelDestino) Then
'          '1o. borra la relacion
'          Borraunarelacion DbDestino, RelDestino.Name
'          '2o. copia la relacion
'          CopiaRelacion DbDestino, RelOrigen
'        End If
'    End If
'  Next RelOrigen
'  Exit Sub
'Error:
'  Select Case Err.Number
'    Case Else
'      MsgBox "VerificaRelaciones:Error no.: " & Err.Number & ":" & Err.Description
'  End Select
'End Sub
'
'Private Function ChecaDetalleRelacion(ByRef RelOrigen As DAO.Relation, ByRef RelDestino As DAO.Relation) As Boolean
'  'verifica sus propiedades que sean iguales
'  Dim CampoOrigen As DAO.Field
'  Dim CampoDestino As DAO.Field
'  On Error GoTo Error
'  ChecaDetalleRelacion = False
'  If RelOrigen.Attributes <> RelDestino.Attributes Then Exit Function
'  If RelOrigen.Table <> RelDestino.Table Then Exit Function
'  If RelOrigen.ForeignTable <> RelDestino.ForeignTable Then Exit Function
'  'checa primero que correspondan los campos de origen a destino
'  For Each CampoOrigen In RelOrigen.Fields
'    If CampoOrigen.ForeignName <> RelDestino.Fields(CampoOrigen.Name).ForeignName Then Exit Function
'  Next CampoOrigen
'  'checa de destino a origen
'  For Each CampoDestino In RelDestino.Fields
'    If CampoDestino.ForeignName <> RelOrigen.Fields(CampoDestino.Name).ForeignName Then Exit Function
'  Next CampoDestino
'  '
'  ChecaDetalleRelacion = True
'  Exit Function
'Error:
'  Err.Clear
'End Function
'
'Private Sub CopiaRelacion(ByRef Base As DAO.Database, ByVal RelOrigen As DAO.Relation)
'  On Error GoTo Error
'  Dim CampoOrigen As DAO.Field
'  Dim RelNuevo As DAO.Relation
'  Set RelNuevo = Base.CreateRelation()
'  RelNuevo.Name = RelOrigen.Name                    '"EmpleadosDepartamentos"
'  RelNuevo.Table = RelOrigen.Table                  'tdfNuevo.Name
'  RelNuevo.ForeignTable = RelOrigen.ForeignTable    'tdfEmpleados.Name
'  RelNuevo.Attributes = RelOrigen.Attributes        'dbRelationUpdateCascade
'  For Each CampoOrigen In RelOrigen.Fields
'    ' Crea el objeto Field para la colección Fields
'    ' del objeto Relation nuevo. Establece las
'    ' propiedades Name y ForeignName basadas en los
'    ' campos que se van a utilizar en la relación.
'    RelNuevo.Fields.Append RelNuevo.CreateField(CampoOrigen.Name)
'    RelNuevo.Fields(CampoOrigen.Name).ForeignName = CampoOrigen.ForeignName
'  Next CampoOrigen
'  Base.Relations.Append RelNuevo
'  Exit Sub
'Error:
'  Select Case Err.Number
'    Case Else
'      MsgBox "CopiaRelacion:Error no.: " & Err.Number & ":" & Err.Description
'  End Select
'End Sub
'
'Private Function ExisteRelacion(ByVal NombreRelacion As String, ByRef Base As DAO.Database, ByRef RelDestino As DAO.Relation) As Boolean
'  Dim Rel As DAO.Relation
'  ExisteRelacion = False
'  For Each Rel In Base.Relations
'    If LCase(NombreRelacion) = LCase(Rel.Name) Then
'      Set RelDestino = Rel
'      ExisteRelacion = True
'      Exit Function
'    End If
'  Next Rel
'End Function
'
'Private Sub ImportaParametros(ByRef Db As DAO.Database, ByVal Prototipo As String)
'  On Error GoTo Error
'  Dim DbNuevo As DAO.Database
'  Dim RsPara As DAO.Recordset
'  Dim Registros As Integer
'  Set DbNuevo = OpenDatabase(Prototipo, False, True)
'  Set RsPara = DbNuevo.OpenRecordset("SELECT * FROM  Parametros", dbOpenDynaset)
'  Do While Not RsPara.EOF
'    Db.Execute "INSERT INTO Parametros (parametro,valor) VALUES ('" & RsPara!Parametro & "','" & RsPara!Valor & "')"
'    Registros = Db.RecordsAffected
'    'If Registros <> 0 Then Stop
'    RsPara.MoveNext
'  Loop
'  Exit Sub
'Error:
'  Select Case Err.Number
'    Case Else
'      MsgBox "Error en NEOGENCV.Verifica.ImportaParametros: " & Err.Number & " - " & Err.Description
'      Resume Next
'  End Select
'End Sub
'
