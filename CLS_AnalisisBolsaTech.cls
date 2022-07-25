Option Explicit
Private OBJ_BusLogic     As CLS_Buslogic
Private Loc_ADO_Registro   As ADODB.Recordset
Private mvarRegistro_Rs    As ADODB.Recordset

'variables locales para almacenar los valores de las propiedades
Private Loc_CartaPorteSucursal         As Integer
Private Loc_CartaPorteNumero           As Long
Private Loc_TipoAnalisis               As String
Private Loc_Tecnologia                 As Integer
Private Loc_ResultadoCualitativo       As String
Private Loc_ResultadoCuantitativo      As Integer
Private Loc_FechaAnalisis              As Date
Private Loc_FechaPresentacion          As Date
Private Loc_FacturaSucursal            As Integer
Private Loc_FacturaNumero              As Long
Private Loc_FacturaFecha               As Date
Private Loc_Tra_Operacion              As String
Private Loc_Tra_Identificacion         As Long
Private Loc_Tra_Sucursal               As Integer
Private Loc_TitularCuit                As String
Private Loc_Tipo                       As String

Public Property Let CartaPorteSucursal(ByVal vData As Variant)
   Loc_CartaPorteSucursal = vData
End Property

Public Property Get CartaPorteSucursal()
   CartaPorteSucursal = Loc_CartaPorteSucursal
End Property

Public Property Let CartaPorteNumero(ByVal vData As Variant)
   Loc_CartaPorteNumero = vData
End Property

Public Property Get CartaPorteNumero()
   CartaPorteNumero = Loc_CartaPorteNumero
End Property

Public Property Let TipoAnalisis(ByVal vData As Variant)
   Loc_TipoAnalisis = vData
End Property

Public Property Get TipoAnalisis()
   TipoAnalisis = Loc_TipoAnalisis
End Property

Public Property Let Tecnologia(ByVal vData As Variant)
   Loc_Tecnologia = vData
End Property

Public Property Get Tecnologia()
   Tecnologia = Loc_Tecnologia
End Property

Public Property Let ResultadoCualitativo(ByVal vData As Variant)
   Loc_ResultadoCualitativo = vData
End Property

Public Property Get ResultadoCualitativo()
   ResultadoCualitativo = Loc_ResultadoCualitativo
End Property

Public Property Let ResultadoCuantitativo(ByVal vData As Variant)
   Loc_ResultadoCuantitativo = vData
End Property

Public Property Get ResultadoCuantitativo()
   ResultadoCuantitativo = Loc_ResultadoCuantitativo
End Property

Public Property Let FechaAnalisis(ByVal vData As Variant)
   Loc_FechaAnalisis = vData
End Property

Public Property Get FechaAnalisis()
   FechaAnalisis = Loc_FechaAnalisis
End Property

Public Property Let FechaPresentacion(ByVal vData As Variant)
   Loc_FechaPresentacion = vData
End Property

Public Property Get FechaPresentacion()
   FechaPresentacion = Loc_FechaPresentacion
End Property

Public Property Let FacturaSucursal(ByVal vData As Variant)
   Loc_FacturaSucursal = vData
End Property

Public Property Get FacturaSucursal()
   FacturaSucursal = Loc_FacturaSucursal
End Property

Public Property Let FacturaNumero(ByVal vData As Variant)
   Loc_FacturaNumero = vData
End Property

Public Property Get FacturaNumero()
   FacturaNumero = Loc_FacturaNumero
End Property

Public Property Let FacturaFecha(ByVal vData As Variant)
   Loc_FacturaFecha = vData
End Property

Public Property Get FacturaFecha()
   FacturaFecha = Loc_FacturaFecha
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

Public Property Let TitularCUIT(ByVal vData As Variant)
   Loc_TitularCuit = vData
End Property

Public Property Get TitularCUIT()
   TitularCUIT = Loc_TitularCuit
End Property

Public Property Let Tipo(ByVal vData As Variant)
   Loc_Tipo = vData
End Property

Public Property Get Tipo()
   Tipo = Loc_Tipo
End Property

Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)
   Set mvarRegistro_Rs = vData
End Property

Public Property Get Registro_RS() As ADODB.Recordset
   Set Registro_RS = Loc_ADO_Registro
End Property

'funcionalidad
Public Function Get_AnalisisBolsaTech(ByVal TitularCUIT As String, ByVal Tipo As String, ByVal CartaPorteSucursal As Integer, ByVal CartaPorteNumero As Long, ByVal Conectar As Boolean) As Boolean
   Get_AnalisisBolsaTech = Get_AnalisisBolsaTech_Local(Loc_ADO_Registro, TitularCUIT, Tipo, CartaPorteSucursal, CartaPorteNumero, Conectar)
   If Get_AnalisisBolsaTech = True Then
      Loc_CartaPorteSucursal = Loc_ADO_Registro!ABO_CARTAPORTESUCURSAL
      Loc_CartaPorteNumero = Loc_ADO_Registro!ABO_CARTAPORTENUMERO
      Loc_TipoAnalisis = Loc_ADO_Registro!ABO_TIPOANALISIS
      Loc_Tecnologia = Loc_ADO_Registro!ABO_TECNOLOGIA
      Loc_ResultadoCualitativo = Loc_ADO_Registro!ABO_RESULTADOCUALITATIVO
      Loc_ResultadoCuantitativo = Loc_ADO_Registro!ABO_RESULTADOCUANTITATIVO
      Loc_FechaAnalisis = Loc_ADO_Registro!ABO_FECHAANALISIS
      Loc_FechaPresentacion = Loc_ADO_Registro!ABO_FECHAPRESENTACION
      Loc_FacturaSucursal = Loc_ADO_Registro!ABO_FACTURASUCURSAL
      Loc_FacturaNumero = Loc_ADO_Registro!ABO_FACTURANUMERO
      Loc_FacturaFecha = Loc_ADO_Registro!ABO_FACTURAFECHA
      Loc_Tra_Operacion = Loc_ADO_Registro!Tra_Operacion
      Loc_Tra_Identificacion = Loc_ADO_Registro!Tra_Identificacion
      Loc_Tra_Sucursal = Loc_ADO_Registro!Tra_Sucursal
      Loc_TitularCuit = Loc_ADO_Registro!ABO_TITULARCUIT
      Loc_Tipo = Loc_ADO_Registro!ABO_TIPO
   End If
End Function

Public Function Get_AnalisisBolsaTech_Local(Registro_RS As ADODB.Recordset, _
                          Optional ByVal TitularCUIT As String, _
                          Optional ByVal Tipo As String, _
                          Optional ByVal CartaPorteSucursal As Integer, _
                          Optional ByVal CartaPorteNumero As Long, _
                          Optional Conectar As Boolean) As Boolean
   Dim Loc_RS         As ADODB.Recordset
   Dim sErrDesc       As String
   Dim lErrNo         As Long
   
   On Error GoTo ManejadorError
   
   Get_AnalisisBolsaTech_Local = False
   
   If Conectar Then
      Call OBJ_BusLogic.ConectarASQL_Aux(True)
   End If
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux
   
   GLO_ADO_Comando.CommandText = "Select * From ANALISISBOLSATECH (Nolock) Where ABO_TITULARCUIT = '" & TitularCUIT & "' and ABO_TIPO = '" & Tipo & "' and ABO_CARTAPORTESUCURSAL = " & CartaPorteSucursal & " and ABO_CARTAPORTENUMERO = " & CartaPorteNumero
   
   GLO_ADO_Comando.CommandType = adCmdText
   
   Set Loc_RS = New ADODB.Recordset
   Loc_RS.CursorLocation = adUseClient
   Loc_RS.CursorType = adOpenStatic
   Loc_RS.LockType = adLockOptimistic
   Loc_RS.Open GLO_ADO_Comando
   
   If Not Loc_RS.EOF Then
     Get_AnalisisBolsaTech_Local = True
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
   GLO_Mensaje = MsgBox("No se Puede Leer el Registro de Análisis Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")
End Function

Public Function Alta_AnalisisBolsaTech() As Boolean
On Error GoTo ManejadorError
   
   Alta_AnalisisBolsaTech = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Alta_AnalisisBolsaTech"
   GLO_ADO_Comando.CommandType = adCmdStoredProc

   '@CARTAPORTESUCURSAL     smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTESUCURSAL", adInteger, adParamInput, , Me.CartaPorteSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CARTAPORTENUMERO       int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTENUMERO", adInteger, adParamInput, , Me.CartaPorteNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPOANALISIS           varchar(20),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPOANALISIS", adVarChar, adParamInput, 20, Mid(Me.TipoAnalisis, 1, 20))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TECNOLOGIA             smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TECNOLOGIA", adInteger, adParamInput, , Me.Tecnologia)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RESULTADOCUALITATIVO   varchar(10),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RESULTADOCUALITATIVO", adVarChar, adParamInput, 10, Mid(Me.ResultadoCualitativo, 1, 20))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RESULTADOCUANTITATIVO  smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RESULTADOCUANTITATIVO", adInteger, adParamInput, , Me.ResultadoCuantitativo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHAANALISIS          varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHAANALISIS", adVarChar, adParamInput, 8, Format(Me.FechaAnalisis, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHAPRESENTACION      varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHAPRESENTACION", adVarChar, adParamInput, 8, Format(Me.FechaPresentacion, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURASUCURSAL          smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURASUCURSAL", adInteger, adParamInput, , Me.FacturaSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURANUMERO            int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURANUMERO", adInteger, adParamInput, , Me.FacturaNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURAFECHA          varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURAFECHA", adVarChar, adParamInput, 8, Format(Me.FacturaFecha, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@IDENTIFICACION            int output
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("IDENTIFICACION", adInteger, adParamInput, , Me.Tra_Identificacion)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@SUCURSAL                 smallint
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TITULARCUIT            varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TITULARCUIT", adVarChar, adParamInput, 11, Me.TitularCUIT)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPO                   varchar(3)
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPO", adVarChar, adParamInput, 3, Me.Tipo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute
   
   Set GLO_ADO_Parametro = Nothing
   Alta_AnalisisBolsaTech = True
   Exit Function

ManejadorError:
End Function

Public Function Modi_AnalisisBolsaTech() As Boolean
On Error GoTo ManejadorError
   
   Modi_AnalisisBolsaTech = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Modi_AnalisisBolsaTech"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
    
   '@CARTAPORTESUCURSAL     smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTESUCURSAL", adInteger, adParamInput, , Me.CartaPorteSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CARTAPORTENUMERO       int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTENUMERO", adInteger, adParamInput, , Me.CartaPorteNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPOANALISIS           varchar(20),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPOANALISIS", adVarChar, adParamInput, 20, Mid(Me.TipoAnalisis, 1, 20))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TECNOLOGIA             smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TECNOLOGIA", adInteger, adParamInput, , Me.Tecnologia)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RESULTADOCUALITATIVO   varchar(10),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RESULTADOCUALITATIVO", adVarChar, adParamInput, 10, Mid(Me.ResultadoCualitativo, 1, 20))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RESULTADOCUANTITATIVO  smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RESULTADOCUANTITATIVO", adInteger, adParamInput, , Me.ResultadoCuantitativo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHAANALISIS          varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHAANALISIS", adVarChar, adParamInput, 8, Format(Me.FechaAnalisis, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHAPRESENTACION      varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHAPRESENTACION", adVarChar, adParamInput, 8, Format(Me.FechaPresentacion, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURASUCURSAL          smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURASUCURSAL", adInteger, adParamInput, , Me.FacturaSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURANUMERO            int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURANUMERO", adInteger, adParamInput, , Me.FacturaNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURAFECHA          varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURAFECHA", adVarChar, adParamInput, 8, Format(Me.FacturaFecha, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@SUCURSAL                 smallint
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TITULARCUIT            varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TITULARCUIT", adVarChar, adParamInput, 11, Me.TitularCUIT)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPO                   varchar(3)
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPO", adVarChar, adParamInput, 3, Me.Tipo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Modi_AnalisisBolsaTech = True
   Exit Function

ManejadorError:
End Function

Public Function Modi_AnalisisBolsaTechFactura() As Boolean
On Error GoTo ManejadorError
   
   Modi_AnalisisBolsaTechFactura = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Modi_AnalisisBolsaTechFactura"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
    
   '@FACTURASUCURSAL          smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURASUCURSAL", adInteger, adParamInput, , Me.FacturaSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURANUMERO            int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURANUMERO", adInteger, adParamInput, , Me.FacturaNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FACTURAFECHA          varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FACTURAFECHA", adVarChar, adParamInput, 8, Format(Me.FacturaFecha, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@ORDENCOMPRA          varchar(20),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ORDENCOMPRA", adVarChar, adParamInput, 20, GLO_BolsaTechOC)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@SUCURSAL                 smallint
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TITULARCUIT               varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TITULARCUIT", adVarChar, adParamInput, 11, Me.TitularCUIT)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPO               varchar(3),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPO", adVarChar, adParamInput, 3, Me.Tipo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Modi_AnalisisBolsaTechFactura = True
   Exit Function

ManejadorError:
End Function

Public Function Baja_AnalisisBolsaTech() As Boolean
On Error GoTo ManejadorError
   
   Baja_AnalisisBolsaTech = False
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Baja_AnalisisBolsaTech"
   GLO_ADO_Comando.CommandType = adCmdStoredProc

   '@CARTAPORTESUCURSAL     smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTESUCURSAL", adInteger, adParamInput, , Me.CartaPorteSucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CARTAPORTENUMERO       int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CARTAPORTENUMERO", adInteger, adParamInput, , Me.CartaPorteNumero)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@SUCURSAL               int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("SUCURSAL", adInteger, adParamInput, , Me.Tra_Sucursal)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TITULARCUIT               varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TITULARCUIT", adVarChar, adParamInput, 11, Me.TitularCUIT)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@TIPO               varchar(3),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("TIPO", adVarChar, adParamInput, 3, Me.Tipo)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   
   GLO_ADO_Comando.Execute
   
   Set GLO_ADO_Parametro = Nothing
   Baja_AnalisisBolsaTech = True
   Exit Function

ManejadorError:
End Function

Private Sub Class_Initialize()
    Set OBJ_BusLogic = New CLS_Buslogic
End Sub

Private Sub Class_Terminate()
    Set OBJ_BusLogic = Nothing
End Sub
