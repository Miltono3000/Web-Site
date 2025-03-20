document.addEventListener("DOMContentLoaded", function() {
    const table = document.getElementById("AvanceTabla");

    const columnasAMostrar = [1, 2, 6, 10, 11, 12, 13]; //Different columns that include the table
    //heading titles
    const encabezados = ['ID', 'Responsable', 'Fecha Termino', 'Avance', 'Avance Semaforo', 'Proximos Pasos', 'Impedimentos'];

    function DateHTML(dates) {
        const excele = new Date(1900,0,1);
        const utcDays = Math.floor(dates - 2); 
        const utcValue = utcDays * 86400000; 
        const date = new Date(excele.getTime() + utcValue);
        return date.toLocaleDateString();
    }

    function ShowData() {
        fetch("http://192.168.112.29:4000/excel")
            .then(response => response.arrayBuffer())
            .then(data => {//Set the sheet to use
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[1];
                const worksheet = workbook.Sheets[sheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                table.innerHTML = "";

                //Add headers
                const trEncabezados = document.createElement("tr");
                encabezados.forEach(texto => {
                    const th = document.createElement("th");
                    th.textContent = texto;
                    trEncabezados.appendChild(th);
                });
                table.appendChild(trEncabezados);

                const filasvalidas = json.filter((row, index) => {
                    return row.some(cell => cell !== null && cell !== '');
                });

                filasvalidas.forEach((row, index) => {
                    if (index === 0) return; 
                    const tr = document.createElement("tr");

                    columnasAMostrar.forEach((colIndex, i) => {
                        const td = document.createElement("td");
                        let cellValue = row[colIndex] || "";

                        switch (i) {
                            case 2: //Set the date form to a column
                                cellValue = DateHTML(cellValue);
                                break;
                            case 3:
                                cellValue =  `${cellValue *100}%`;
                                break;
                            case 4: //Set the traffic light form to a column
                                function Semaforo(valor) {
                                    if (valor <= 0.49){
                                        return "🔴";
                                    }
                                    if (valor <= 0.74){
                                        return "🟠";
                                    } 
                                    if (valor <= 0.99){
                                        return "🟡";
                                    }
                                    if (valor = 1){
                                        return "🟢";
                                    }
                                }
                                cellValue = Semaforo(cellValue);
                                break;
                        }

                        td.textContent = cellValue;
                        tr.appendChild(td);
                    });

                    table.appendChild(tr);
                });
            })
            .catch(error => {
                console.error("Error al cargar el archivo Excel:", error);
            });
    }
    ShowData();//Show the data of the columns

    setInterval(ShowData, 5000);//Automatically update the page (add or remove)
});