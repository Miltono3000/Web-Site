async function obtenerDatos() {
    const respuesta = await fetch('http://192.168.112.29:3000/autofill');
    return await respuesta.json();
}
//AutoComplete the first five fields according to the ID
async function completarDatos() {
    const IDvalue = document.getElementById('id').value;
    const completar = await obtenerDatos();

    const register = completar.find(reg => reg.id == IDvalue);

    document.getElementById('nombre').value = register ? register.nombre : '';
    document.getElementById('area').value = register ? register.area : '';
    document.getElementById('fecha').value = register ? register.fecha : '';
    document.getElementById('proyecto').value = register ? register.proyecto : '';
    document.getElementById('descripcion').value = register ? register.descripcion : '';
}

async function enviarDatos() {
    let ID_Proyecto =  document.getElementById('id').value;
    let Responsable = document.getElementById('nombre').value;
    let Area = document.getElementById('area').value;
    let Fecha_Proyecto = document.getElementById('fecha').value;
    let Proyecto = document.getElementById('proyecto').value;
    let Descripcion = document.getElementById('descripcion').value;
    let Fecha_Inicio = document.getElementById('FI').value;
    let Fecha_Termino = document.getElementById('FT').value;
    let Prioridad = document.getElementById('Prioridad').value;
    let Areas_Involucradas = document.getElementById('areainvolucrada').value;
    let Avance = document.getElementById('avance').value;
    let Proximos_pasos = document.getElementById('proximopaso').value;
    let Impedimentos = document.getElementById('Impedimentos').value;
    let Observaciones = document.getElementById('observacion').value;

    function comparar(Fecha_Proyecto, Fecha_Inicio) {
        const date1 = Date.parse(Fecha_Proyecto.split('/').reverse().join('-')); //Change the format to YYYY-MM-DD
        const date2 = Date.parse(Fecha_Inicio.split('/').reverse().join('-'));
        return date2 < date1;
    }

    if(Avance > 100){
        alert('Digite un número valido'); //The field "Avance" can not be higher than 100
        return;
    }else{
        if(Responsable === ""||Area === ""||Proyecto === ""||Descripcion === "") {
            alert('ID invalido o no registrado'); //Need a registered ID
            event.preventDefault();
        }else { //The date of the project can not be before the date of the progress
            if(comparar(Fecha_Proyecto,Fecha_Inicio)){
                alert('La fecha de Inicio del avance no puede ser antes que la Fecha de Inicio del proyecto');
                event.preventDefault();
            }else{//The end date can not be before the start date
                if(Fecha_Termino < Fecha_Inicio){
                    alert('Fechas de Termino no puede estar antes a la Fecha de Inicio');
                    event.preventDefault();
                }else{ //No fiels empty
                    if(ID_Proyecto === ""||Fecha_Inicio === ""||Fecha_Termino === ""||Prioridad === ""||Areas_Involucradas === ""||
                        Avance === ""||Proximos_pasos === ""||Impedimentos === ""||Observaciones === ""){
                        alert('Completa los campos faltantes');
                        event.preventDefault();
                    }else{
                    let response = await fetch('http://192.168.112.29:3000/api/datos', {
                        method: 'POST', 
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({ID_Proyecto, Responsable, Area, Fecha_Proyecto, Proyecto, Descripcion, Fecha_Inicio, Fecha_Termino,
                                            Prioridad, Areas_Involucradas, Avance, Proximos_pasos, Impedimentos, Observaciones
                        })  
                    }); //Send the data to the excel.
                    if(response.ok) {
                        alert('Datos enviados correctamente');
                        location.reload();
                        let descargaResonse = await fetch('http://192.168.112.29:3000/api/download', {
                            method: 'GET'
                        });
                    }else {
                        alert('Error al enviar datos');
                    }
                    }  
                }
            }
        }
    }
}      