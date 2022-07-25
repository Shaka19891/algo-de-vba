Option Explicit

Private OBJ_BusLogic             As CLS_Buslogic
Private Loc_ADO_Registro         As ADODB.Recordset 'copia local
Private mvarRegistro_Rs          As ADODB.Recordset

'variables locales para almacenar los valores de las propiedades
Private Loc_Camara                  As Integer
Private Loc_Numero                  As Long
Private Loc_Fecha                   As Date
Private Loc_Renglon                 As Integer
Private Loc_Item                    As Integer
Private Loc_Porcentaje              As Double
Private Loc_BonificaRebaja          As Double
Private Loc_Porcentaje_Al_XX        As Double
Private Loc_Observacion             As String
Private Loc_Grado                   As Integer
Private Loc_BonificaRebajaCalculada As Double
Private Loc_GradoCalculado          As Integer
Private Loc_Tra_Operacion        As String
Private Loc_Tra_Identificacion   As Long
Private Loc_Tra_Sucursal         As Integer

Public Property Let Camara(ByVal vData As Variant)
   Loc_Camara = vData
End Property

Public Property Get Camara()
   Camara = Loc_Camara
End Property

Public Property Let Numero(ByVal vData As Variant)
   Loc_Numero = vData
End Property

Public Property Get Numero()
   Numero = Loc_Numero
End Property

Public Property Let Fecha(ByVal vData As Variant)
   Loc_Fecha = vData
End Property

Public Property Get Fecha()
   Fecha = Loc_Fecha
End Property

Public Property Let Renglon(ByVal vData As Variant)
   Loc_Renglon = vData
End Property

Public Property Get Renglon()
   Renglon = Loc_Renglon
End Property

Public Property Let Item(ByVal vData As Variant)
   Loc_Item = vData
End Property

Public Property Get Item()
   Item = Loc_Item
End Property

Public Property Let Porcentaje(ByVal vData As Variant)
   Loc_Porcentaje = vData
End Property

Public Property Get Porcentaje()
   Porcentaje = Loc_Porcentaje
End Property

Public Property Let BonificaRebaja(ByVal vData As Variant)
   Loc_BonificaRebaja = vData
End Property

Public Property Get BonificaRebaja()
   BonificaRebaja = Loc_BonificaRebaja
End Property

Public Property Let Porcentaje_Al_XX(ByVal vData As Variant)
   Loc_Porcentaje_Al_XX = vData
End Property

Public Property Get Porcentaje_Al_XX()
   Porcentaje_Al_XX = Loc_Porcentaje_Al_XX
End Property

Public Property Let Observacion(ByVal vData As Variant)
   Loc_Observacion = vData
End Property

Public Property Get Observacion()
   Observacion = Loc_Observacion
End Property

Public Property Let Grado(ByVal vData As Variant)
   Loc_Grado = vData
End Property

Public Property Get Grado()
   Grado = Loc_Grado
End Property

Public Property Let BonificaRebajaCalculada(ByVal vData As Variant)
   Loc_BonificaRebajaCalculada = vData
End Property

Public Property Get BonificaRebajaCalculada()
   BonificaRebajaCalculada = Loc_BonificaRebajaCalculada
End Property

Public Property Let GradoCalculado(ByVal vData As Variant)
   Loc_GradoCalculado = vData
End Property

Public Property Get GradoCalculado()
   GradoCalculado = Loc_GradoCalculado
End Property

Public Property Let Tra_Operacion(ByVal vData As Variant)
   Loc_Tra_Operacion = vData
End Property

Public Property Get Tra_Operacion()
   Tra_Operacion = Loc_Tra_Operacion
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
Public Function Get_AnalisisVentaRenglon(ByVal Camara As Integer, ByVal Numero As Long, ByVal Fecha As String, ByVal Renglon As Integer, ByVal Item As Integer, ByVal Conectar As Boolean) As Boolean
   Get_AnalisisVentaRenglon = Get_AnalisisVentaRenglon_Local(Loc_ADO_Registro, Camara, Numero, Fecha, Renglon, Item, Conectar)
   If Get_AnalisisVentaRenglon = True Then
      'cargo el objeto con el recordset temporario
      Loc_Camara = Loc_ADO_Registro!AVR_CAMARA
      Loc_Numero = Loc_ADO_Registro!AVR_NUMERO
      Loc_Fecha = Loc_ADO_Registro!AVR_FECHA
      Loc_Renglon = Loc_ADO_Registro!AVR_RENGLON
      Loc_Item = Loc_ADO_Registro!AVR_ITEM
      Loc_Porcentaje = Loc_ADO_Registro!AVR_PORCENTAJE
      Loc_BonificaRebaja = Loc_ADO_Registro!AVR_BONIFICAREBAJA
      Loc_Porcentaje_Al_XX = Loc_ADO_Registro!AVR_PORCENTAJE_AL_XX
      Loc_Observacion = Loc_ADO_Registro!AVR_OBSERVACION
      Loc_Grado = Loc_ADO_Registro!AVR_GRADO
      Loc_BonificaRebajaCalculada = Loc_ADO_Registro!AVR_BONIFICAREBAJACALCULADA
      Loc_GradoCalculado = Loc_ADO_Registro!AVR_GRADOCALCULADO
      Loc_Tra_Operacion = Loc_ADO_Registro!Tra_Operacion
      Loc_Tra_Identificacion = Loc_ADO_Registro!Tra_Operacion
      Loc_Tra_Sucursal = Loc_ADO_Registro!Tra_Identificacion
   End If
End Function

Public Function Get_AnalisisVentaRenglon_Local(Registro_RS As ADODB.Recordset, ByVal Camara As Integer, ByVal Numero As Long, ByVal Fecha As String, ByVal Renglon As Integer, ByVal Item As Integer, Optional Conectar As Boolean) As Boolean
   Dim Loc_RS         As ADODB.Recordset
   
   On Error GoTo ManejadorError
   Get_AnalisisVentaRenglon_Local = False

   If Conectar Then
      Call OBJ_BusLogic.ConectarASQL_Aux(True)
   End If
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux
   
   GLO_ADO_Comando.CommandText = "Select * From ANALISISVENTARENGLON (Nolock) Where AVR_CAMARA = " & Camara & " and AVR_NUMERO = " & Numero & " and AVR_FECHA = '" & Fecha & "' and AVR_RENGLON = " & Renglon & " and AVR_ITEM = " & Item & " "
   GLO_ADO_Comando.CommandType = adCmdText

   Set Loc_RS = New ADODB.Recordset
   Loc_RS.CursorLocation = adUseClient
   Loc_RS.CursorType = adOpenStatic
   Loc_RS.LockType = adLockOptimistic
   Loc_RS.Open GLO_ADO_Comando
   
   If Not Loc_RS.EOF Then
     Get_AnalisisVentaRenglon_Local = True
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

Public Function Alta_AnalisisVentaRenglon() As Boolean
On Error GoTo ManejadorError

   Alta_AnalisisVentaRenglon = False
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   GLO_ADO_Comando.CommandText = "Alta_AnalisisVentaRenglon"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
   
   '@CAMARA              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CAMARA", adInteger, adParamInput, , Me.Camara)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@NUMERO              int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("NUMERO", adInteger, adParamInput, , Me.Numero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@FECHA               varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHA", adVarChar, adParamInput, 8, Format(Me.Fecha, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@RENGLON                   smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RENGLON", adInteger, adParamInput, , Me.Renglon)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@ITEM                      smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ITEM", adInteger, adParamInput, , Me.Item)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PORCENTAJE                numeric (14,4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PORCENTAJE", adDouble, adParamInput, 1, Me.Porcentaje)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@BONIFICAREBAJA            numeric (14,4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("BONIFICAREBAJA", adDouble, adParamInput, , Me.BonificaRebaja)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PORCENTAJE_AL_XX          numeric(14, 4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PORCENTAJE_AL_XX", adDouble, adParamInput, , Me.Porcentaje_Al_XX)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@OBSERVACION               varchar(30),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("OBSERVACION", adVarChar, adParamInput, 30, Me.Observacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@GRADO                     smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("GRADO", adInteger, adParamInput, , Me.Grado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@BONIFICAREBAJACALCULADA   numeric(14, 4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("BONIFICAREBAJACALCULADA", adDouble, adParamInput, , Me.BonificaRebajaCalculada)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@GRADOCALCULADO            smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("GRADOCALCULADO", adInteger, adParamInput, , Me.GradoCalculado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@OPERACION           varchar (1),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("OPERACION", adVarChar, adParamInput, 1, Me.Tra_Operacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IDENTIFICACION      int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IDENTIFICACION", adInteger, adParamInput, , Me.Tra_Identificacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@SUCURSAL            smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute

   Set GLO_ADO_Parametro = Nothing
   Alta_AnalisisVentaRenglon = True
   Exit Function

ManejadorError:
End Function

Public Function Modi_AnalisisVentaRenglon() As Boolean
On Error GoTo ManejadorError
   
   Modi_AnalisisVentaRenglon = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Modi_AnalisisVentaRenglon"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
   
   '@CAMARA              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CAMARA", adInteger, adParamInput, , Me.Camara)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@NUMERO              int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("NUMERO", adInteger, adParamInput, , Me.Numero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@FECHA               varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHA", adVarChar, adParamInput, 8, Format(Me.Fecha, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@RENGLON                   smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RENGLON", adInteger, adParamInput, , Me.Renglon)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@ITEM                      smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ITEM", adInteger, adParamInput, , Me.Item)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PORCENTAJE                numeric (14,4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PORCENTAJE", adDouble, adParamInput, 1, Me.Porcentaje)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@BONIFICAREBAJA            numeric (14,4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("BONIFICAREBAJA", adDouble, adParamInput, , Me.BonificaRebaja)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@PORCENTAJE_AL_XX          numeric(14, 4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("PORCENTAJE_AL_XX", adDouble, adParamInput, , Me.Porcentaje_Al_XX)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@OBSERVACION               varchar(30),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("OBSERVACION", adVarChar, adParamInput, 30, Me.Observacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@GRADO                     smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("GRADO", adInteger, adParamInput, , Me.Grado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@BONIFICAREBAJACALCULADA   numeric(14, 4),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("BONIFICAREBAJACALCULADA", adDouble, adParamInput, , Me.BonificaRebajaCalculada)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@GRADOCALCULADO            smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("GRADOCALCULADO", adInteger, adParamInput, , Me.GradoCalculado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@OPERACION           varchar (1),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("OPERACION", adVarChar, adParamInput, 1, Me.Tra_Operacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@IDENTIFICACION      int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IDENTIFICACION", adInteger, adParamInput, , Me.Tra_Identificacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   '@SUCURSAL            smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   GLO_ADO_Comando.Execute
   
   Set GLO_ADO_Parametro = Nothing
   Modi_AnalisisVentaRenglon = True
   Exit Function

ManejadorError:
End Function

Private Sub Class_Initialize()
    Set OBJ_BusLogic = New CLS_Buslogic
End Sub

Private Sub Class_Terminate()
    Set OBJ_BusLogic = Nothing
End Sub

