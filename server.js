const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const app2 = express();

const wb = new ExcelJS.Workbook();

const sheetNames = ['Registros', 'Progreso'];
sheetNames.forEach(sheetName => {
    let worksheet = wb.addWorksheet(sheetName);
    worksheet.state = 'visible';
});
const hoja2 = wb.getWorksheet('Progreso');
hoja2.mergeCells('K1:L1');  

app2.use(cors());
app2.use(bodyParser.json());
app2.use(express.static(path.join(__dirname)));

app2.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'templates', 'ID.html'));
});

let datos2 = [];

app2.get('/fill', async (req, res) => {
    const sheets = new ExcelJS.Workbook();
    await sheets.xlsx.readFile(path.join(__dirname, 'Proyectos.xlsx'));
    const sheet1 = sheets.getWorksheet('Registros');

    const datafill = [];
    sheet1.eachRow((row, rowNumber) => {
        if(rowNumber > 1){
            datafill.push({
                Responsable: row.getCell(2).value,
                Proyect: row.getCell(6).value,
                Description: row.getCell(7).value
            });
        }
    });
    res.json(datafill);
})

app2.get('/FillSel', (req, res) =>{
    const filePath = path.join(__dirname, 'Proyectos.xlsx');
    const Workbook = xlsx.readFile(filePath);
    const sheetName = Workbook.SheetNames[0];
    const sheet = Workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
    res.json(data);
})

app2.post('/guardar', async (req, res) => {
    const { Responsable, Area, Fecha_Inicio, Proyecto, Descripcion} = req.body;

    datos2.push({...req.body});

    console.log('Datos recibidos:', datos2); // Verifica los datos en la consola

    const worksheet = wb.getWorksheet('Registros');

    worksheet.columns = [
        { header: 'Num', key: 'Num', width: 5},                     //A 1
        { header: 'ID Generado', key: 'ID_Generado', width: 20},    //B 2
        { header: 'Responsable', key: 'Responsable', width: 40},    //C 3
        { header: 'Área', key: 'Area', width: 20},                  //D 4
        { header: 'Fecha Inicio', key: 'Fecha_Inicio', width: 14},  //E 5
        { header: 'Proyecto', key: 'Proyecto', width: 30},          //F 6
        { header: 'Descripción', key: 'Descripcion', width: 90}     //G 7
    ];

    let rowsearch = 0;
    const nuevoDato = datos2[datos2.length - 1];
    var ID_Valor = ``;
    const FI = new Date(nuevoDato.Fecha_Inicio);
    const year = FI.getFullYear();
    
    worksheet.eachRow({includeEmpty: false}, (row, rowNumber) => {
        const dateE = new Date(row.getCell(5).value);
        const yearE = dateE.getFullYear();
        if(row.getCell(3).value === nuevoDato.Responsable && row.getCell(4).value === nuevoDato.Area 
            && row.getCell(6).value === nuevoDato.Proyecto && yearE === year){
            rowsearch = rowNumber;
        }
    });

    var IDalert;
    var count = 1;
    var No_Proyecto;
    if(rowsearch !== 0){
        IDalert = worksheet.getCell(`B${rowsearch}`).value;
    }else{ 
        worksheet.eachRow({includeEmpty: false}, (row, rowNumber) => {
            const WSFI = new Date(row.getCell(5).value);
            const WSY = WSFI.getFullYear();
            if(row.getCell(4).value === nuevoDato.Area && WSY === year){
                count++;
            }
        });
        No_Proyecto = count;
        ID_Valor = `${nuevoDato.Area}${String(No_Proyecto).padStart(3, '0')}${year}`;
   
        worksheet.addRow({
            ID_Generado: ID_Valor,
            Responsable: nuevoDato.Responsable,
            Area: nuevoDato.Area,
            Fecha_Inicio: FI,
            Proyecto: nuevoDato.Proyecto,
            Descripcion: nuevoDato.Descripcion
        }); 
        IDalert = ID_Valor;
    }

    worksheet.eachRow({includeEmpty: false}, (row, rowNumber) => {
        row.getCell(1).value = rowNumber - 1;
    });

    //A1 se queda como 'Num'
    worksheet.getCell('A1').value = "Num";

    for(let i = 1; i < 8; i++){
        worksheet.getColumn(i).font = {
            name: 'Noto Sans', 
            size: 10
        }
        worksheet.getColumn(i).border = {
            top: { style: 'thin'},
            left: { style: 'thin' },
            bottom: { style: 'thin'},
            right: { style: 'thin'}
        };
        worksheet.getColumn(i).alignment = {
            vertical: 'middle', 
            horizontal: 'center',
            wrapText: true
        };
    }

    const row = worksheet.getRow(1);
    for(let i = 1;i < 8; i++){
        row.getCell(i).font = { 
            name: 'Noto Sans', 
            size: 10, 
            bold: true,
            color: { argb: 'FFFFFF'}
        };
        row.getCell(i).alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        row.getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '793d4c'}
        };
        row.getCell(i).border = {
            top: { style: 'thin'},
            left: { style: 'thin' },
            right: { style: 'thin'},
            color: { argb: 'AEABAB'}
        };
    }
    const filePath = path.join(__dirname, 'Proyectos.xlsx');

    try {
        await wb.xlsx.writeFile(filePath);
        console.log('Archivo Excel generado correctamente en:', filePath);

        // Verifica que el archivo existe antes de intentar enviarlo
        if (!fs.existsSync(filePath)) {
            throw new Error('El archivo no se generó correctamente');
        }
        res.json({IDalert});
    } catch (err) {
        console.error('Error al guardar el archivo Excel:', err);
    }
});

// Iniciar el servidor
const PORT2 = 80;
app2.listen(PORT2, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT2}`);
});


//Otra pagina
const app1 = express();
const PORT = 3000;

app1.use(cors());
app1.use(bodyParser.json());
app1.use(express.static(path.join(__dirname)));

app1.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'templates', 'index.html'));
});

app1.get('/autofill', async (req, res) => {
    function formatoFecha(date){
        if(!date){
            return '';
        }
        const d = new Date(date);
        const day = String(d.getDate() + 1).padStart(2,'0');
        const month = String(d.getMonth() + 1).padStart(2,'0');
        const year = d.getFullYear();
        return `${day}/${month}/${year}`;
    }
    const sheets = new ExcelJS.Workbook();
    await sheets.xlsx.readFile(path.join(__dirname, 'Proyectos.xlsx'));
    const sheet1 = sheets.getWorksheet(1);

    const datfill = [];
    sheet1.eachRow((row, rowNumber) => {
        if(rowNumber > 1){
            datfill.push({
                id: row.getCell(2).value,
                nombre: row.getCell(3).value,
                area: row.getCell(4).value,
                fecha: formatoFecha(row.getCell(5).value),
                proyecto: row.getCell(6).value,
                descripcion: row.getCell(7).value
            });
        }
    });
    res.json(datfill);
})

let datos = [];

app1.post('/api/datos', (req, res) => {
    const { ID_Proyecto, Responsable, Area, Fecha_Proyecto, Proyecto, Descripcion, Fecha_Inicio, Fecha_Termino, Prioridad,
        Areas_Involucradas, Avance, Proximos_pasos, Impedimentos, Observaciones } = req.body;

    const Porcentaje = parseFloat(Avance) / 100;

    // Agrega los datos
    datos.push({...req.body, Avance: Porcentaje});

    console.log('Datos recibidos:', datos); // Verifica los datos en la consola
    res.status(200).send('Datos recibidos');
});

app1.get('/api/download', async (req, res) => {
    // Verifica si hay datos para exportar
    if (datos.length === 0) {
        return res.status(400).send('No hay datos para exportar');
    }
    //Todo se hara sobre la hoja de Progreso
    const ws = wb.getWorksheet('Progreso');
    // Definir los encabezados de las columnas
    ws.columns = [
        { header: 'No', key: 'No', width: 8},                                   //A     1
        { header: 'ID Proyecto', key: 'ID_Proyecto', width: 30},                //B     2
        { header: 'Responsable', key: 'Responsable', width: 40},                //C     3
        { header: 'Área', key: 'Area', width: 20},                              //D     4
        { header: 'Proyecto', key: 'Proyecto', width: 25},                      //E     5
        { header: 'Descripción', key: 'Descripcion', width: 70},                //F     6
        { header: 'Fecha Inicio', key: 'Fecha_Inicio', width: 14},              //G     7
        { header: 'Fecha Término', key: 'Fecha_Termino', width: 14},            //H     8
        { header: 'Prioridad', key: 'Prioridad', width: 10},                    //I     9
        { header: 'Areas Involucradas', key: 'Areas_Involucradas', width: 22},  //J     10
        { header: 'Avance', key: 'Avance', width: 15},                          //K     11
        { header: 'Avance', key: 'Icono', width: 5},                            //L     12
        { header: 'Proximos Pasos', key: 'Proximos_pasos', width: 50},          //M     13
        { header: 'Impedimentos', key: 'Impedimentos', width: 50},              //N     14
        { header: 'Observaciones', key: 'Observaciones', width: 50},            //O     15
        { header: 'Fecha del Proyecto', key: 'Fecha_Proyecto', width: 19}       //P     16
    ];
    
    // Agregar los datos a la hoja
    datos.forEach((item) => {
        const FechaI = new Date(item.Fecha_Inicio);
        const FechaT = new Date(item.Fecha_Termino);
        let rowDel = 0;
        ws.eachRow({includeEmpty: false}, (row, rowNumber) => {
            if(row.getCell(2).value === item.ID_Proyecto){
                rowDel = rowNumber;
            }
        });

        if(rowDel !== 0){
            const Num = ws.getCell(`A${rowDel}`).value;
            ws.spliceRows(rowDel, 1);
            ws.addRow({
                No: Num,
                ID_Proyecto: item.ID_Proyecto,
                Responsable: item.Responsable,
                Area: item.Area,
                Proyecto: item.Proyecto,
                Descripcion: item.Descripcion,
                Fecha_Inicio: FechaI,
                Fecha_Termino: FechaT,
                Prioridad: item.Prioridad,
                Areas_Involucradas: item.Areas_Involucradas,
                Avance: item.Avance,
                Icono: item.Avance,
                Proximos_pasos: item.Proximos_pasos,
                Impedimentos: item.Impedimentos,
                Observaciones: item.Observaciones,
                Fecha_Proyecto: item.Fecha_Proyecto
            });
        }else{
        ws.addRow({
            ID_Proyecto: item.ID_Proyecto,
            Responsable: item.Responsable,
            Area: item.Area,
            Proyecto: item.Proyecto,
            Descripcion: item.Descripcion,
            Fecha_Inicio: FechaI,
            Fecha_Termino: FechaT,
            Prioridad: item.Prioridad,
            Areas_Involucradas: item.Areas_Involucradas,
            Avance: item.Avance,
            Icono: item.Avance,
            Proximos_pasos: item.Proximos_pasos,
            Impedimentos: item.Impedimentos,
            Observaciones: item.Observaciones,
            Fecha_Proyecto: item.Fecha_Proyecto
        });
    }
    //Agregar Secuencia de registros al excel
    ws.eachRow({includeEmpty: false, startRow: 2}, (row, rowNumber) => {
        row.getCell(1).value = rowNumber - 1;
    });
    //A1 se queda como 'No'
    ws.getCell('A1').value = "No";
    });
    //Se le agrega el formato de porcentaje a la columna de avance (K)
    ws.getColumn(11).numFmt = '0%';
    //Formato de Semaforo a la columna L. Rojo 0-49, Amarillo 50-75, Verde 76-100
    const minValor = 0.5;
    const maxValor = 0.75;
    ws.addConditionalFormatting({
        ref: 'L1:L1048570',
        rules: [
            {
                type: 'iconSet',
                iconSet: '3TrafficLights1',
                cfvo: [{type: 'num', value: 0}, {type: 'num', value: minValor}, {type: 'num', value: maxValor}],
                showValue: false
            }
        ]
    })
    //Barra de datos en la columna K
    const minValue = 0.01;
    const maxValue = 1;
    ws.addConditionalFormatting({
        ref: 'K1:K1048570',
        rules: [
            {
                type: 'dataBar',
                cfvo: [{type: 'num', value: minValue}, {type: 'num', value: maxValue}],
                color: {argb: "66ff4e"},
                direction: 'leftToRight'
            }
        ]
    });
    //Formato a cada columna
    for(let i = 1; i < 17; i++){
        ws.getColumn(i).font = {
            name: 'Noto Sans', 
            size: 10
        }
        ws.getColumn(i).border = {
            top: { style: 'thin'},
            left: { style: 'thin' },
            bottom: { style: 'thin'},
            right: { style: 'thin'}
        };
        ws.getColumn(i).alignment = {
            vertical: 'middle', 
            horizontal: 'center',
            wrapText: true
        };
    }
    //Encabezado (fila 1) Con tipo de letra especifico, tamaño, negritas, color de letra
    const row = ws.getRow(1);
    for(let i = 1;i < 17; i++){
        row.getCell(i).font = { 
            name: 'Noto Sans', 
            size: 10, 
            bold: true,
            color: { argb: 'FFFFFF'}
        };
        row.getCell(i).alignment = {
            vertical: 'middle',
            horizontal: 'center'
        };
        row.getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '793d4c'}
        };
        row.getCell(i).border = {
            top: { style: 'thin'},
            left: { style: 'thin' },
            right: { style: 'thin'},
            color: { argb: 'AEABAB'}
        };
    }
    // Guardar el archivo Excel
    const filePath = path.join(__dirname, 'Proyectos.xlsx');

    try {
        await wb.xlsx.writeFile(filePath);
        console.log('Archivo Excel generado correctamente en:', filePath);
        // Verifica que el archivo existe antes de intentar enviarlo
        if (!fs.existsSync(filePath)) {
            throw new Error('El archivo no se generó correctamente');
        }
        // Enviar el archivo para descargar
        res.download(filePath, 'datos.xlsx', (err) => {
            if (err) {
                console.error('Error al descargar el archivo:', err);
                if (!res.headersSent) {
                    res.status(500).send('Error al descargar el archivo');
                }
            } else {
                console.log('Archivo enviado correctamente');
            }
        });
    } catch (err) {
        console.error('Error al guardar el archivo Excel:', err);
        if (!res.headersSent) {
            res.status(500).send('Error al generar el archivo Excel');
        }
    }
});

app1.listen(PORT, '0.0.0.0', () => {
    console.log(`Servidor 2 corriendo en http://localhost:${PORT}`);
});

//Pagina para la parte visual
const app3 = express();
const PORT3 = 4000;

app3.use(cors());
app3.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'templates', 'Tabla.html'));
});

app3.use(express.static(path.join(__dirname)));

app3.get("/excel", (req, res) => {
    const filePath = path.join(__dirname, "Proyectos.xlsx");
    res.sendFile(filePath);
});

app3.listen(PORT3, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT3}`);
});