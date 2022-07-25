Option Explicit

Private OBJ_BusLogic             As CLS_Buslogic
Private Loc_ADO_Registro         As ADODB.Recordset 'copia local
Private mvarRegistro_Rs          As ADODB.Recordset

Private Loc_Codigo                     As Long
Private Loc_Descripcion                As String
Private Loc_Imputacion                 As String
Private Loc_TipoEntidad                As Integer
Private Loc_CodigoCTacte               As Long
Private Loc_Uen                        As String
Private Loc_Predeterminada             As String
Private Loc_Tra_Identificacion         As Long
Private Loc_Tra_Sucursal               As Integer

Public Property Let Codigo(ByVal vData As Variant)
   Loc_Codigo = vData
End Property

Public Property Get Codigo()
   Codigo = Loc_Codigo
End Property

Public Property Let Descripcion(ByVal vData As Variant)
   Loc_Descripcion = vData
End Property

Public Property Get Descripcion()
   Descripcion = Loc_Descripcion
End Property

Public Property Let Imputacion(ByVal vData As Variant)
   Loc_Imputacion = vData
End Property

Public Property Get Imputacion()
   Imputacion = Loc_Imputacion
End Property

Public Property Let TipoEntidad(ByVal vData As Variant)
   Loc_TipoEntidad = vData
End Property

Public Property Get TipoEntidad()
   TipoEntidad = Loc_TipoEntidad
End Property

Public Property Let CodigoCTacte(ByVal vData As Variant)
   Loc_CodigoCTacte = vData
End Property

Public Property Get CodigoCTacte()
   CodigoCTacte = Loc_CodigoCTacte
End Property

Public Property Let Uen(ByVal vData As Variant)
   Loc_Uen = vData
End Property

Public Property Get Uen()
   Uen = Loc_Uen
End Property

Public Property Let Predeterminada(ByVal vData As Variant)
   Loc_Predeterminada = vData
End Property

Public Property Get Predeterminada()
   Predeterminada = Loc_Predeterminada
End Property

Public Property Let Tra_Identificacion(ByVal vData As Variant)
   Loc_Tra_Identificacion = vData
End Property

Public Property Get Tra_Identificacion()
   Tra_Identificacion = Loc_Tra_Identificacion
End Property

Public Property Let Tra_Sucursal(ByVal vData As Variant)
   Loc_Tra_Sucursal = vData
End Property

Public Property Get Tra_Sucursal()
   Tra_Sucursal = Loc_Tra_Sucursal
End Property

Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)
   Set mvarRegistro_Rs = vData
End Property

Public Property Get Registro_RS() As ADODB.Recordset
   Set Registro_RS = Loc_ADO_Registro
End Property

'funcionalidad
Public Function Get_Agrupaciones(ByVal Codigo As String, ByVal Conectar As Boolean) As Boolean
   Get_Agrupaciones = Get_Agrupaciones_Local(Loc_ADO_Registro, Codigo, Conectar)
   If Get_Agrupaciones = True Then
      'cargo el objeto con el recordset temporario
      Loc_Codigo = Loc_ADO_Registro!AGP_CODIGO
      Loc_Descripcion = Loc_ADO_Registro!AGP_DESCRIPCION
      Loc_Imputacion = Loc_ADO_Registro!AGP_IMPUTACION
      Loc_TipoEntidad = Loc_ADO_Registro!AGP_TIPOENTIDAD
      Loc_CodigoCTacte = Loc_ADO_Registro!AGP_CODIGOCTACTE
      Loc_Uen = Loc_ADO_Registro!AGP_UEN
      Loc_Predeterminada = Loc_ADO_Registro!AGP_PREDETERMINADA
      Loc_Tra_Identificacion = Loc_ADO_Registro!Tra_Identificacion
      Loc_Tra_Sucursal = Loc_ADO_Registro!Tra_Sucursal
   End If
End Function

Public Function Get_Agrupaciones_Local(Registro_RS As ADODB.Recordset, ByVal Codigo As String, Optional Conectar As Boolean) As Boolean
   Dim Loc_RS         As ADODB.Recordset
   
   On Error GoTo ManejadorError
   Get_Agrupaciones_Local = False
   
   If Conectar Then
      Call OBJ_BusLogic.ConectarASQL_Aux(True)
   End If
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux
   GLO_ADO_Comando.CommandText = "Select * From Agrupaciones (NoLock) Where AGP_CODIGO = " & Codigo & " "
   GLO_ADO_Comando.CommandType = adCmdText

   Set Loc_RS = New ADODB.Recordset
   Loc_RS.CursorLocation = adUseClient
   Loc_RS.CursorType = adOpenStatic
   Loc_RS.LockType = adLockOptimistic
   Loc_RS.Open GLO_ADO_Comando
   
   If Not Loc_RS.EOF Then
     Get_Agrupaciones_Local = True
   End If
   
   Set Registro_RS = Loc_RS
   Set Loc_RS = Nothing
   
   If Conectar Then
      Call OBJ_BusLogic.DesconectarDeSql_Aux
   End If
   Exit Function
   
ManejadorError:
   Set Loc_RS = Nothing
   Set GLO_ADO_Conexion_Aux = Nothing
End Function

Private Sub Class_Initialize()
    Set OBJ_BusLogic = New CLS_Buslogic
End Sub

Private Sub Class_Terminate()
    Set OBJ_BusLogic = Nothing
End Sub

Public Function Alta_Agrupaciones() As Boolean
   
   Alta_Agrupaciones = False

   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Alta_Agrupaciones"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
   
   '@CODIGO           int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CODIGO", adInteger, adParamInput, , Codigo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@DESCRIPCION      varchar (50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("DESCRIPCION", adVarChar, adParamInput, 50, Mid(Me.Descripcion, 1, 50))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IMPUTACION       varchar (35),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IMPUTACION", adVarChar, adParamInput, 35, Mid(Me.Imputacion, 1, 35))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@TIPOENTIDAD      smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPOENTIDAD", adInteger, adParamInput, , Me.TipoEntidad)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@CODIGOCTACTE     int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CODIGOCTACTE", adInteger, adParamInput, , CodigoCTacte)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@UEN              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PREDETERMINADA   varchar (1),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PREDETERMINADA", adVarChar, adParamInput, 1, Mid(Me.Predeterminada, 1, 1))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IDENTIFICACION   int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IDENTIFICACION", adInteger, adParamInput, , GLO_Identificacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@SUCURSAL         smallint
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , GLO_CodigoUen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Alta_Agrupaciones = True
   Exit Function

End Function

Public Function Modi_Agrupaciones() As Boolean
   
   Modi_Agrupaciones = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Modi_Agrupaciones"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
   
   '@CODIGO           int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CODIGO", adInteger, adParamInput, , Codigo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@DESCRIPCION      varchar (50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("DESCRIPCION", adVarChar, adParamInput, 50, Mid(Me.Descripcion, 1, 50))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IMPUTACION       varchar (35),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IMPUTACION", adVarChar, adParamInput, 35, Mid(Me.Imputacion, 1, 35))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@TIPOENTIDAD      smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPOENTIDAD", adInteger, adParamInput, , Me.TipoEntidad)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@CODIGOCTACTE     int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CODIGOCTACTE", adInteger, adParamInput, , CodigoCTacte)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@UEN              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PREDETERMINADA   varchar (1),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PREDETERMINADA", adVarChar, adParamInput, 1, Mid(Me.Predeterminada, 1, 1))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IDENTIFICACION   int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IDENTIFICACION", adInteger, adParamInput, , GLO_Identificacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@SUCURSAL         smallint
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , GLO_CodigoUen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   

 
   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Modi_Agrupaciones = True
   Exit Function

End Function

Public Function Baja_Agrupaciones() As Boolean
   
   Baja_Agrupaciones = False
      
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Baja_Agrupaciones"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
   
   '@CODIGO           int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CODIGO", adInteger, adParamInput, , Codigo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
 
   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Baja_Agrupaciones = True
   Exit Function

End Function
