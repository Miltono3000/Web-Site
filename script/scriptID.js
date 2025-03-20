//Get all the data from a sheet
async function obtenerDatos() {
    const respuesta = await fetch('http://192.168.112.29:80/fill');
    return await respuesta.json();
}
async function SelFill() {
    try {
        const respuesta = await fetch('http://192.168.112.29:80/FillSel');
        if (!respuesta.ok) {
            throw new Error(`Error en la respuesta: ${respuesta.status} ${respuesta.statusText}`);
        }
        const data = await respuesta.json();
        return data;
    } catch (error) {
        console.error("Error en SelFill()", error);
        throw error;
    }
}
//Add the option to add a new project
function AddProject() {
    const selProject = document.getElementById('Proyect');
    const NewProject = document.getElementById('NewProyect');

    if(selProject.value === "Add"){
        NewProject.style.display = 'block';
        NewProject.focus();
    }else{
        NewProject.style.display = 'none';
    }
}
//Update the select options according to the name
async function UpdateSelect() {
    const FillData = document.getElementById('Responsable').value;
    const selProject = document.getElementById('Proyect');

    try {
        let completar = await SelFill();
        console.log("Datos obtenidos:", completar);

        const proyectosFiltrados = completar
            .filter(reg => reg.Responsable === FillData)
            .map(reg => reg.Proyecto);

        selProject.innerHTML = `
            <option value="">Seleccione un Proyecto</option>
            <option value="Add">Agregar Proyecto</option>
        `;

        proyectosFiltrados.forEach(proyecto => {
            const option = document.createElement('option');
            option.value = proyecto;
            option.textContent = proyecto;
            selProject.appendChild(option);
        });

    } catch (error) {
        console.error("Error en UpdateSelect()", error);
        alert("No se pudieron cargar los proyectos.");
    }
}
//Update the field "Descripcion" according to the differents projects.
async function UpdateDesc() {
    let FillDesc;
    let SelDesc = document.getElementById('Proyect').value;
    if(SelDesc === 'Add'){
        FillDesc = document.getElementById('NewProyect').value;
    }else{
        FillDesc = document.getElementById('Proyect').value;
    }
    let completar = await obtenerDatos();

    const register = completar.find(reg => reg.Proyect == FillDesc);
    document.getElementById('Description').value = register ? register.Description : '';
}

async function GenerateID() {
    event.preventDefault();

    let Responsable = document.getElementById('Responsable').value;
    let Area = document.getElementById('area').value;
    let Fecha_Inicio = document.getElementById('FI').value;
    let Proyecto = document.getElementById('Proyect').value;
    let Descripcion = document.getElementById('Description').value;

    if(Proyecto === "Add"){
        Proyecto = document.getElementById('NewProyect').value;
        if(Proyecto === ""){
            alert('Ingrese el nombre del Nuevo Proyecto');
            return;
        }
    }

    if(Responsable === ""||Area === ""||Fecha_Inicio === ""||Proyecto === ""||Descripcion === ""){
        alert('Completa los campos faltantes');
        return; 
    }
else{
    try {
        let response = await fetch('http://192.168.112.29:80/guardar', {
            method: 'POST',
            headers: {
                'Content-type': 'application/json'
            },
            body: JSON.stringify({Responsable, Area, Fecha_Inicio, Proyecto, Descripcion})
        });
    
        if(!response.ok){
            throw new Error('Error en el servidor');
        }
        //Sow the ID to the user
        let data = await response.json();
        let id = data.IDalert;
        alert(`Su ID es: ${id}`);
        location.reload();
    } catch (error) {
        console.error('Error:', error);
        alert(`Hubo un error al generar el id: ${error}`);
    }
}
}