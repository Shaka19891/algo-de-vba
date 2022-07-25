Option Explicit

Private OBJ_BusLogic     As CLS_Buslogic
Private Loc_ADO_Registro     As ADODB.Recordset
Private mvarRegistro_Rs     As ADODB.Recordset

Private Loc_Id       As Long
Private Loc_Uen       As Integer
Private Loc_Apellido       As String
Private Loc_FechaCarga       As Date


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

Public Property Let Apellido(ByVal vData As Variant)
     Loc_Apellido = vData
End Property

Public Property Get Apellido()
     Apellido = Loc_Apellido
End Property

Public Property Let FechaCarga(ByVal vData As Variant)
     Loc_FechaCarga = vData
End Property

Public Property Get FechaCarga()
     FechaCarga = Loc_FechaCarga
End Property

Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)
     Set mvarRegistro_Rs = vData
End Property

Public Property Get Registro_RS() As ADODB.Recordset
     Set Registro_RS = Loc_ADO_Registro
End Property


Public Function Get_TablaCD(ByVal Id As Long, ByVal Conectar As Boolean) As Boolean
     Get_TablaCD = Get_TablaCD_Local(Loc_ADO_Registro, Id, Conectar)
     If Get_TablaCD = True Then
          Loc_Id = Loc_ADO_Registro!TTD_ID
          Loc_Uen = Loc_ADO_Registro!TTD_UEN
          Loc_Apellido = Loc_ADO_Registro!TTD_APELLIDO
          Loc_FechaCarga = Loc_ADO_Registro!TTD_FECHA
     End If
End Function

Public Function Get_TablaCD_Local(Registro_RS As ADODB.Recordset, ByVal Id As Long, Optional Conectar As Boolean) As Boolean
     Dim Loc_RS As ADODB.Recordset
     Dim sErrDesc As String
     Dim lErrNo As Long

     On Error GoTo ManejadorError

     Get_TablaCD_Local = False

     If Conectar Then
        Call Obj_buslogic.ConectarASQL_Aux(True)
     End If

     Set GLO_ADO_Comando = New ADODB.Command
     GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux

     GLO_ADO_Comando.CommandText = "Select * From TABLACD (Nolock) Where TTD_ID = " & Id & " "

     GLO_ADO_Comando.CommandType = adCmdText

     Set Loc_RS = New ADODB.Recordset
     Loc_RS.CursorLocation = adUseClient
     Loc_RS.CursorType = adOpenStatic
     Loc_RS.LockType = adLockOptimistic
     Loc_RS.Open GLO_ADO_Comando

     If Not Loc_RS.EOF Then
          Get_TablaCD_Local = True
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
     GLO_Mensaje = MsgBox("No se Puede Leer el Registro de TablaCD Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")

End Function


Public Function Alta_TablaCD() As Boolean
     Alta_TablaCD = False
     Set GLO_ADO_Comando = New ADODB.Command
     GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
     GLO_ADO_Comando.CommandText = "Alta_TablaCD"
     GLO_ADO_Comando.CommandType = adCmdStoredProc

'     @UEN               smallint,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'     @APELLIDO               varchar,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("APELLIDO", adVarChar, adParamInput, 50, Me.Apellido)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'     @FECHA               varchar,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHA", adVarChar, adParamInput, 8, Format(Me.FechaCarga, "YYYYMMDD"))
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

     GLO_ADO_Comando.Execute
     Set GLO_ADO_Parametro = Nothing
     Alta_TablaCD = True
End Function

Public Function Baja_TablaCD() As Boolean
     Baja_TablaCD = False
     Set GLO_ADO_Comando = New ADODB.Command
     GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
     GLO_ADO_Comando.CommandText = "Baja_TablaCD"
     GLO_ADO_Comando.CommandType = adCmdStoredProc

'     @ID               int,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ID", adInteger, adParamInput, , Me.Id)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

     GLO_ADO_Comando.Execute
     Set GLO_ADO_Parametro = Nothing
     Baja_TablaCD = True
End Function

Public Function Modi_TablaCD() As Boolean
     Modi_TablaCD = False
     Set GLO_ADO_Comando = New ADODB.Command
     GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion
     GLO_ADO_Comando.CommandText = "Modi_TablaCD"
     GLO_ADO_Comando.CommandType = adCmdStoredProc

'     @UEN               smallint,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("UEN", adInteger, adParamInput, , Me.Uen)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'     @APELLIDO               varchar,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("APELLIDO", adVarChar, adParamInput, 50, Me.Apellido)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'     @FECHA               varchar,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("FECHA", adVarChar, adParamInput, 8, Format(Me.FechaCarga, "YYYYMMDD"))
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro
'     @ID               int,
     Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("ID", adInteger, adParamInput, , Me.Id)
     GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro

     GLO_ADO_Comando.Execute
     Set GLO_ADO_Parametro = Nothing
     Modi_TablaCD = True
End Function
