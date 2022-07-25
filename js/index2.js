let lista_Atributos = document.querySelector("#atributos");

document.querySelector("#generar-atributo").addEventListener("click",agregarAtributo);
document.querySelector("#btn-codigo").addEventListener("click",generarCodigo);

function agregarAtributo(){
    let item = document.createElement("li");
    item.innerHTML= `<div>
                        <label>Nombre de Atributo: </label>
                        <input type="text" placeholder="nombre">
                    </div>
                    <div>
                    <label>Tipo de Atributo: </label>
                    <input type="text" placeholder="tipo">
                    </div>`;

                    lista_Atributos.appendChild(item);
}
                
function generarCodigo(){
    
    let arreglo_de_atributos= [];
    
    let items = lista_Atributos.children;
    for (const iterator of items) {
        let variable={};
        variable.nombre = iterator.firstElementChild.lastElementChild.value;
        variable.tipo = iterator.lastElementChild.lastElementChild.value;
        arreglo_de_atributos.push(variable);
    }

    let parrafo_variables = document.querySelector("#variables");
    let parrafo_SetsLets = document.querySelector("#gets-sets");

    parrafo_variables.innerHTML = "";
    parrafo_SetsLets.innerHTML = "";

    parrafo_variables.innerHTML+= `Option Explicit <br><br>
    Private OBJ_BusLogic  &nbsp;&nbsp;&nbsp;   As CLS_Buslogic <br>
    Private Loc_ADO_Registro  &nbsp;&nbsp;&nbsp;  As ADODB.Recordset <br>
    Private mvarRegistro_Rs  &nbsp;&nbsp;&nbsp;   As ADODB.Recordset <br><br>`;
    
    for (const atrib of arreglo_de_atributos) {
        parrafo_variables.innerHTML+=  `Private Loc_${atrib.nombre} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; As ${atrib.tipo}<br>`;
        parrafo_SetsLets.innerHTML+=   `Public Property Let ${atrib.nombre}(ByVal vData As Variant)<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Loc_${atrib.nombre} = vData <br>
                                        End Property <br><br>
                                        Public Property Get ${atrib.nombre}()<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${atrib.nombre} = Loc_${atrib.nombre}<br>
                                        End Property<br><br>`;
    }
    parrafo_variables.innerHTML+=  `<br>`;
    parrafo_SetsLets.innerHTML+=    `Public Property Let Registro_RS(ByVal vData As ADODB.Recordset)<br>
                                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set mvarRegistro_Rs = vData <br>
                                     End Property <br><br>
                                     Public Property Get Registro_RS() As ADODB.Recordset<br>
                                     &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Set Registro_RS = Loc_ADO_Registro<br>
                                     End Property<br><br>`;
}