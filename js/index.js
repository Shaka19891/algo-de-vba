let lista_Atributos = document.querySelector("#atributos");

document.querySelector("#generar-atributo").addEventListener("click",agregarAtributo);
document.querySelector("#btn-codigo").addEventListener("click",generarCodigo);

function agregarAtributo(){
    let item = document.createElement("li");
    item.innerHTML= `<div>
                        <label>Nombre Atrib: </label>
                        <input class="medium" type="text" placeholder="Nombre Atributo">
                    </div>
                    <div>
                        <label>Tipo Atrib: </label>
                        <input class="medium" type="text" placeholder="Tipo Atributo">
                    </div>
                    <div>
                        <label>Nombre col BBDD: </label>
                        <input class="medium" type="text" placeholder="Columna BBDD">
                    </div>
                    <div>
                        <label>Tipo dato BBDD:</label>
                        <input class="medium" type="text" placeholder="Tipo BBDD">
                        <input class="num" type="number" placeholder="cant varchar(n)">
                    </div>
                    <div>
                        <label>Key: </label>
                        <input type="checkbox">
                    </div>`;

                    lista_Atributos.appendChild(item);
}
                
function generarCodigo(){
    
    let arreglo_de_atributos= [];
    let nombre_de_clase = document.querySelector("#nombre-clase").value;
    let prefijo = document.querySelector("#prefijo-tabla").value;

    let items = lista_Atributos.children;
    for (const iterator of items) {
        let variable={};
        variable.nombre = iterator.firstElementChild.lastElementChild.value;
        variable.tipo = iterator.firstElementChild.nextElementSibling.lastElementChild.value;
        variable.columna = iterator.firstElementChild.nextElementSibling.nextElementSibling.lastElementChild.value;
        variable.tipoc = iterator.firstElementChild.nextElementSibling.nextElementSibling.nextElementSibling.firstElementChild.nextElementSibling.value;
        variable.cant = iterator.firstElementChild.nextElementSibling.nextElementSibling.nextElementSibling.lastElementChild.value;
        variable.clave = iterator.lastElementChild.lastElementChild.checked;
        arreglo_de_atributos.push(variable);
    }

    let parrafo_variables = document.querySelector("#variables");
    let parrafo_SetsLets = document.querySelector("#gets-sets");
    let parrafo_Get = document.querySelector("#get");
    let parrafo_AltaBajaModi = document.querySelector("#alta-baja-modi");

    parrafo_variables.innerHTML = "";
    parrafo_SetsLets.innerHTML = "";
    parrafo_Get.innerHTML = "";
    parrafo_AltaBajaModi.innerHTML = "";
    
    parrafo_variables.innerHTML+= ` Option Explicit <br><br>
                                    Private OBJ_BusLogic  &nbsp;&nbsp;&nbsp;   As CLS_Buslogic <br>
                                    Private Loc_ADO_Registro  &nbsp;&nbsp;&nbsp;  As ADODB.Recordset <br>
                                    Private mvarRegistro_Rs  &nbsp;&nbsp;&nbsp;   As ADODB.Recordset <br><br>`;

    let texto1_get= `Public Function Get_${nombre_de_clase}(ByVal`;
    let texto2_get= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase} = Get_${nombre_de_clase}_Local(Loc_ADO_Registro`;
    let texto3_get= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Get_${nombre_de_clase} = True Then<br>`;
    
    let texto1_get_local= `Public Function Get_${nombre_de_clase}_Local(Registro_RS As ADODB.Recordset`;
    let texto2_get_local= ``;
    
    let pre_Alta= ` Public Function Alta_${nombre_de_clase}() As Boolean<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Alta_${nombre_de_clase} = False<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Comando = New ADODB.Command<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandText = "Alta_${nombre_de_clase}"<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandType = adCmdStoredProc<br><br>`
    ;

    let pre_Modi= ` Public Function Modi_${nombre_de_clase}() As Boolean<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Modi_${nombre_de_clase} = False<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Comando = New ADODB.Command<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandText = "Modi_${nombre_de_clase}"<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandType = adCmdStoredProc<br><br>`
    ;

    let pre_Baja= ` Public Function Baja_${nombre_de_clase}() As Boolean<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Baja_${nombre_de_clase} = False<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Comando = New ADODB.Command<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandText = "Baja_${nombre_de_clase}"<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandType = adCmdStoredProc<br><br>`
    ;

    let post_Alta= `<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.Execute<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Parametro = Nothing<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Alta_${nombre_de_clase} = True<br>
                    End Function<br><br>`
    ;

    let post_Modi= `<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.Execute<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Parametro = Nothing<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Modi_${nombre_de_clase} = True<br>
                    End Function<br><br>`
    ;

    let post_Baja= `<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.Execute<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Parametro = Nothing<br>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Baja_${nombre_de_clase} = True<br>
                    End Function<br><br>`
    ;

    let clave_AltaBajaModi = "";
    let noClave_AltaBajaModi = "";

    let primeraVez= true;

    for (const atrib of arreglo_de_atributos) {
        parrafo_variables.innerHTML+=  `Private Loc_${atrib.nombre} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; As ${atrib.tipo}<br>`;

        parrafo_SetsLets.innerHTML+=   `Public Property Let ${atrib.nombre}(ByVal vData As Variant)<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_${atrib.nombre} = vData <br>
                                        End Property <br><br>
                                        Public Property Get ${atrib.nombre}()<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${atrib.nombre} = Loc_${atrib.nombre}<br>
                                        End Property<br><br>`
        ;

        if (atrib.clave == true){
            if (primeraVez){
                primeraVez = false;
                texto1_get+= ` ${atrib.nombre} As ${atrib.tipo}`;
                texto2_get_local+=`${prefijo}_${atrib.columna} = " & ${atrib.nombre} & "`;
            }
            else{
                texto1_get+= `, ByVal ${atrib.nombre} As ${atrib.tipo}`;
                texto2_get_local+=` AND ${prefijo}_${atrib.columna} = " & ${atrib.nombre} & "`;
            }
            texto2_get+= `, ${atrib.nombre}`;
            texto1_get_local+= `, ByVal ${atrib.nombre} As ${atrib.tipo}`;

            clave_AltaBajaModi+= `  '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;@${atrib.columna}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${atrib.tipoc},<br>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("${atrib.columna}", `
            
            switch (atrib.tipoc) {
                case 'int':
                    clave_AltaBajaModi+= `adInteger, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                case 'smallint':
                    clave_AltaBajaModi+= `adInteger, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                case 'varchar':
                    if(atrib.columna.includes('FECHA'))
                        clave_AltaBajaModi+= `adVarChar, adParamInput, ${atrib.cant}, Format(Me.${atrib.nombre}, "YYYYMMDD"))<br>`
                    else
                        clave_AltaBajaModi+= `adVarChar, adParamInput, ${atrib.cant}, Me.${atrib.nombre})<br>`
                    break;
                case 'numeric':
                    clave_AltaBajaModi+= `adDouble, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                default:
                    console.log('entro al default con : '+ atrib.tipoc);
                }

            clave_AltaBajaModi+= `  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro<br>`;
        }
        else{
            noClave_AltaBajaModi+= `  '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;@${atrib.columna}&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${atrib.tipoc},<br>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Parametro = GLO_ADO_Comando.CreateParameter("${atrib.columna}", `

            switch (atrib.tipoc) {
                case 'int':
                    noClave_AltaBajaModi+= `adInteger, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                case 'smallint':
                    noClave_AltaBajaModi+= `adInteger, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                case 'varchar':
                    if(atrib.columna.includes('FECHA'))
                        noClave_AltaBajaModi+= `adVarChar, adParamInput, ${atrib.cant}, Format(Me.${atrib.nombre}, "YYYYMMDD"))<br>`
                    else
                        noClave_AltaBajaModi+= `adVarChar, adParamInput, ${atrib.cant}, Me.${atrib.nombre})<br>`
                    break;
                case 'numeric':
                    clave_AltaBajaModi+= `adDouble, adParamInput, , Me.${atrib.nombre})<br>`
                    break;
                default:
                    console.log('entro al default con : '+ atrib.tipoc);
            }

            noClave_AltaBajaModi+= `  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.Parameters.Append GLO_ADO_Parametro<br>`;
        }

        texto3_get+= `&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_${atrib.nombre} = Loc_ADO_Registro!${prefijo}_${atrib.columna}<br>`;
        
    }
    
    texto1_get+=`, ByVal Conectar As Boolean) As Boolean<br>`;
    texto2_get+=`, Conectar)<br>`;
    texto3_get+=`&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br>
                 End Function<br><br>`;

    texto2_get_local+=` "<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandType = adCmdText<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = New ADODB.Recordset<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.CursorLocation = adUseClient<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.CursorType = adOpenStatic<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.LockType = adLockOptimistic<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_RS.Open GLO_ADO_Comando<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Not Loc_RS.EOF Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase}_Local = True<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Registro_RS = Loc_RS<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = Nothing<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Conectar Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Call Obj_buslogic.DesconectarDeSql_Aux<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Exit Function<br><br>
                        ManejadorError:<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Loc_RS = Nothing<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Conexion_Aux = Nothing<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_Mensaje = MsgBox("No se Puede Leer el Registro de ${nombre_de_clase} Seleccionado" & vbCr & "Error: " & CStr(Err.Number) & "-" & Err.Description, vbOKOnly + vbCritical, "Advertencia")<br><br>
                        End Function<br><br>`;

    texto1_get_local+=  `, Optional Conectar As Boolean) As Boolean<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim Loc_RS         As ADODB.Recordset<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim sErrDesc       As String<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim lErrNo         As Long<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;On Error GoTo ManejadorError<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Get_${nombre_de_clase}_Local = False<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If Conectar Then<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Call Obj_buslogic.ConectarASQL_Aux(True)<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set GLO_ADO_Comando = New ADODB.Command<br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.ActiveConnection = GLO_ADO_Conexion_Aux<br><br>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GLO_ADO_Comando.CommandText = "Select * From ${nombre_de_clase.toUpperCase()} (Nolock) Where `;

    parrafo_variables.innerHTML+=  `<br>`;
    parrafo_SetsLets.innerHTML +=  `Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)<br>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set mvarRegistro_Rs = vData <br>
                                    End Property <br><br>
                                    Public Property Get Registro_RS() As ADODB.Recordset<br>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Registro_RS = Loc_ADO_Registro<br>
                                    End Property<br><br>`;

    parrafo_Get.innerHTML += texto1_get;
    parrafo_Get.innerHTML += texto2_get;
    parrafo_Get.innerHTML += texto3_get;
    parrafo_Get.innerHTML += texto1_get_local;
    parrafo_Get.innerHTML += texto2_get_local;
    
    parrafo_AltaBajaModi.innerHTML += pre_Alta;
    parrafo_AltaBajaModi.innerHTML += noClave_AltaBajaModi;
    parrafo_AltaBajaModi.innerHTML += post_Alta;
    
    parrafo_AltaBajaModi.innerHTML += pre_Baja;
    parrafo_AltaBajaModi.innerHTML += clave_AltaBajaModi;
    parrafo_AltaBajaModi.innerHTML += post_Baja;
    
    parrafo_AltaBajaModi.innerHTML += pre_Modi;
    parrafo_AltaBajaModi.innerHTML += clave_AltaBajaModi;
    parrafo_AltaBajaModi.innerHTML += noClave_AltaBajaModi;
    parrafo_AltaBajaModi.innerHTML += post_Modi;
}