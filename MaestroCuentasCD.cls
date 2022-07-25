Option Explicit

Private Obj_buslogic As CLS_Buslogic
Private Loc_ADO_Registro As ADODB.Recordset 'copia local
Private mvarRegistro_Rs As ADODB.Recordset

'variables locales para almacenar los valores de las propiedades
Private Loc_Id                As Long
Private Loc_Uen               As Integer
Private Loc_CuitFirmante      As String
Private Loc_Apellido          As String
Private Loc_Nombre            As String
Private Loc_RazonSocial       As String
Private Loc_CuitEmpresa       As String
Private Loc_Mail              As String
Private Loc_Celular           As String
Private Loc_FechaCarga        As Date
Private Loc_Estado            As String

Public Property Let Id(ByVal vData As Variant)
   Loc_Id = vData
End Property

Public Property Get Id()
   Id = Loc_Id
End Property

Public Property Let Uen(ByVal vData As Variant)
   Loc_Uen = vData
End Property

Public Property Get Uen()
   Uen = Loc_Uen
End Property

Public Property Let CuitFirmante(ByVal vData As Variant)
   Loc_CuitFirmante = vData
End Property

Public Property Get CuitFirmante()
   CuitFirmante = Loc_CuitFirmante
End Property

Public Property Let Apellido(ByVal vData As Variant)
   Loc_Apellido = vData
End Property

Public Property Get Apellido()
   Apellido = Loc_Apellido
End Property

Public Property Let Nombre(ByVal vData As Variant)
   Loc_Nombre = vData
End Property

Public Property Get Nombre()
   Nombre = Loc_Nombre
End Property

Public Property Let RazonSocial(ByVal vData As Variant)
   Loc_RazonSocial = vData
End Property

Public Property Get RazonSocial()
   RazonSocial = Loc_RazonSocial
End Property

Public Property Let CuitEmpresa(ByVal vData As Variant)
   Loc_CuitEmpresa = vData
End Property

Public Property Get CuitEmpresa()
   CuitEmpresa = Loc_CuitEmpresa
End Property

Public Property Let Mail(ByVal vData As Variant)
   Loc_Mail = vData
End Property

Public Property Get Mail()
   Mail = Loc_Mail
End Property

Public Property Let Celular(ByVal vData As Variant)
   Loc_Celular = vData
End Property

Public Property Get Celular()
   Celular = Loc_Celular
End Property

Public Property Let FechaCarga(ByVal vData As Variant)
   Loc_FechaCarga = vData
End Property

Public Property Get FechaCarga()
   FechaCarga = Loc_FechaCarga
End Property

Public Property Let Estado(ByVal vData As Variant)
   Loc_Estado = vData
End Property

Public Property Get Estado()
   Estado = Loc_Estado
End Property

Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)
Set mvarRegistro_Rs = vData
End Property

Public Property Get Registro_RS() As ADODB.Recordset
    Set Registro_RS = Loc_ADO_Registro
End Property

'funcionalidad
Public Function Get_MaestroCuentasCD(ByVal Id As Long, ByVal Conectar As Boolean) As Boolean
   Get_MaestroCuentasCD = Get_MaestroCuentasCD_Local(Loc_ADO_Registro, Id, Conectar)
   If Get_MaestroCuentasCD = True Then
      'cargo el objeto con el recordset temporario
      Loc_Id = Loc_ADO_Registro!CCD_ID
      Loc_Uen = Loc_ADO_Registro!CCD_UEN
      Loc_CuitFirmante = Loc_ADO_Registro!CCD_CUITFIRMANTE
      Loc_Apellido = Loc_ADO_Registro!CCD_APELLIDO
      Loc_Nombre = Loc_ADO_Registro!CCD_NOMBRE
      Loc_RazonSocial = Loc_ADO_Registro!CCD_RAZONSOCIAL
      Loc_CuitEmpresa = Loc_ADO_Registro!CCD_CUITEMPRESA
      Loc_Mail = Loc_ADO_Registro!CCD_MAIL
      Loc_Celular = Loc_ADO_Registro!CCD_CELULAR
      Loc_FechaCarga = Loc_ADO_Registro!CCD_FECHACARGA
      Loc_Estado = Loc_ADO_Registro!CCD_ESTADO
   End If
End Function

Public Function Get_MaestroCuentasCD_Local(Registro_RS As ADODB.Recordset, ByVal Id As Long, Optional Conectar As Boolean) As Boolean
   Dim Loc_RS         As ADODB.Recordset
   Dim sErrDesc       As String
   Dim lErrNo         As Long
   
   On Error GoTo ManejadorError
   
   Get_MaestroCuentasCD_Local = False
   
   If Conectar Then
      Call Obj_buslogic.ConectarASQL_Aux(True)
   End If
   
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux
   
   GLO_ADO_Comando.CommandText = "Select * From MAESTROCUENTASCD (Nolock) Where CCD_ID = " & Id & " "
   
   GLO_ADO_Comando.CommandType = adCmdText
   
   Set Loc_RS = New ADODB.Recordset
   Loc_RS.CursorLocation = adUseClient
   Loc_RS.CursorType = adOpenStatic
   Loc_RS.LockType = adLockOptimistic
   Loc_RS.Open GLO_ADO_Comando
   
   If Not Loc_RS.EOF Then
     Get_MaestroCuentasCD_Local = True
   End If
   
   Set Registro_RS = Loc_RS
   Set Loc_RS = Nothing
   
   If Conectar Then
      Call Obj_buslogic.DesconectarDeSql_Aux
   End If
   
   Exit Function
   
ManejadorError:
   
   Set Loc_RS = Nothing
   Set GLO_ADO_Conexion_Aux = Nothing
   GLO_Mensaje = MsgBox("No se Puede Leer el Registro de MaestroCuentasCD Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")

End Function

Public Function Alta_MaestroCuentasCD() As Boolean
   Alta_MaestroCuentasCD = False
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   GLO_ADO_Comando.CommandText = "Alta_MaestroCuentasCD"
   GLO_ADO_Comando.CommandType = adCmdStoredProc

   '@ID               int,

   '@UEN              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CUITFIRMANTE     varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CUITFIRMANTE", adVarChar, adParamInput, 11, Me.CuitFirmante)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@APELLIDO         varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("APELLIDO", adVarChar, adParamInput, 50, Me.Apellido)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@NOMBRE           varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("NOMBRE", adVarChar, adParamInput, 50, Me.Nombre)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RAZONSOCIAL      varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RAZONSOCIAL", adVarChar, adParamInput, 50, Me.RazonSocial)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CUITEMPRESA      varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CUITEMPRESA", adVarChar, adParamInput, 11, Me.CuitEmpresa)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@MAIL             varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("MAIL", adVarChar, adParamInput, 50, Me.Mail)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CELULAR          varchar(20),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CELULAR", adVarChar, adParamInput, 20, Me.Celular)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHACARGA       varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHACARGA", adVarChar, adParamInput, 8, Format(Me.FechaCarga, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@ESTADO           varchar(1)
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ESTADO", adVarChar, adParamInput, 1, Me.Estado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Alta_MaestroCuentasCD = True
End Function

Public Function Baja_MaestroCuentasCD() As Boolean
   Baja_MaestroCuentasCD = False
   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   
   GLO_ADO_Comando.CommandText = "Baja_MaestroCuentasCD"
   GLO_ADO_Comando.CommandType = adCmdStoredProc
    
   '@ID               int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ID", adInteger, adParamInput, , Me.Id)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute

   Set GLO_ADO_Parametro = Nothing
   Baja_MaestroCuentasCD = True
End Function

Public Function Modi_MaestroCuentasCD() As Boolean
   Modi_MaestroCuentasCD = False

   Set GLO_ADO_Comando = New ADODB.Command
   GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
   GLO_ADO_Comando.CommandText = "Modi_MaestroCuentasCD"
   GLO_ADO_Comando.CommandType = adCmdStoredProc

   '@ID               int,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ID", adInteger, adParamInput, , Me.Id)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@UEN              smallint,
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CUITFIRMANTE     varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CUITFIRMANTE", adVarChar, adParamInput, 11, Me.CuitFirmante)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@APELLIDO         varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("APELLIDO", adVarChar, adParamInput, 50, Me.Apellido)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@NOMBRE           varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("NOMBRE", adVarChar, adParamInput, 50, Me.Nombre)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@RAZONSOCIAL      varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("RAZONSOCIAL", adVarChar, adParamInput, 50, Me.RazonSocial)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CUITEMPRESA      varchar(11),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CUITEMPRESA", adVarChar, adParamInput, 11, Me.CuitEmpresa)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@MAIL             varchar(50),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("MAIL", adVarChar, adParamInput, 50, Me.Mail)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@CELULAR          varchar(20),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("CELULAR", adVarChar, adParamInput, 20, Me.Celular)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@FECHACARGA       varchar(8),
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHACARGA", adVarChar, adParamInput, 8, Format(Me.FechaCarga, "YYYYMMDD"))
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
   '@ESTADO           varchar(1)
   Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ESTADO", adVarChar, adParamInput, 1, Me.Estado)
   GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

   GLO_ADO_Comando.Execute
   Set GLO_ADO_Parametro = Nothing
   Modi_MaestroCuentasCD = True
End Function

Private Sub Class_Initialize()
    Set Obj_buslogic = New CLS_Buslogic
End Sub

Private Sub Class_Terminate()
    Set Obj_buslogic = Nothing
End Sub