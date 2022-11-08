VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Replicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Var
Private cUtO As DBUtil
Private cUtD As DBUtil

Private Enum IntegridadEnum
  Correcto = 0
  Incorrecto = 1
End Enum
  

Public Function Replicar(BaseOrigen As String, BaseDestino As String) As Integer
  ' Var
  Dim intRes As Integer
  
  Set cUtO = New DBUtil
  Set cUtD = New DBUtil
  
  ' Asigna las cadenas de la base origen y la base destino a los objetos DBUtil
  cUtO.CadenaConexion = BaseOrigen
  cUtD.CadenaConexion = BaseDestino
  
  cUtO.CadenaConexion = BaseOrigen
  cUtD.CadenaConexion = BaseDestino
  intRes = VerificaIntegridadBaseDeDatos()
  If intRes = Correcto Then
    intRes = ReplicaBase()
  End If
End Function

Public Function VerificaIntegridadBaseDeDatos() As Integer
  ' Var
  Dim tblO As Tabla
  Dim tblD As Tabla
  Dim tblsO As Tablas
  Dim tblsD As Tablas
  Dim strSQLO As String
  Dim strSQLD As String
  Dim intRes As Integer
  Dim intReconstruir As Integer
  
  ' Genera las conexiones a las bases de datos Origen y Destino
  strSQLO = cUtO.ObtieneConsultaSQLDeSeleccion("Version", "", "")
  strSQLD = cUtD.ObtieneConsultaSQLDeSeleccion("Version", "", "")
  Set tblO = cUtO.ConsultaDeResultados(strSQLO)
  Set tblD = cUtD.ConsultaDeResultados(strSQLD)
  
  ' Verifica si los valores son iguales
  If tblO(1).Columnas("Version").Valor <> tblD(1).Columnas("Version").Valor Then
    Set tblsO = cUtO.ObtieneTablas
    Set tblsD = cUtD.ObtieneTablas
    ' if tblso.Count = tbls
  End If
End Function

Private Function ReplicaBase()

End Function
