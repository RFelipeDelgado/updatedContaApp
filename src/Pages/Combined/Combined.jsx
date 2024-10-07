/* eslint-disable react/prop-types */
import { useState } from 'react';
import { DataGrid } from '@mui/x-data-grid';
import { Box, Button, Typography, Modal, TextField } from '@mui/material';
import { useNavigate } from 'react-router-dom';
import { Workbook } from 'exceljs';
import { saveAs } from 'file-saver';

const Combined = ({ combinedData }) => {
    const [userInput, setUserInput] = useState({
        year: '2024',
        name: 'Nombre',
        initialBalance: '1000',
        rutUser: '9.999.999-K'
    });
    // Estado para controlar la apertura del modal (popup dojnde se ingresa userData)
    const [open, setOpen] = useState(false);
    // Funciones para abrir y cerrar el modal
    const handleOpen = () => setOpen(true);
    const handleClose = () => setOpen(false);

    const navigate = useNavigate();

    // const columns = [
    //     { field: 'correlativo', headerName: 'Número correlativo', width: 130 },
    //     { field: 'folio', headerName: 'N° de documento', width: 130 },
    //     { field: 'tipoDocumento', headerName: 'Tipo Documento', width: 130 },
    //     { field: 'rutEmisor', headerName: 'RUT EMISOR', width: 130 },
    //     { field: 'fechaOperacion', headerName: 'FECHA OPERACIÓN', width: 130 },
    //     { field: 'montoNeto', headerName: 'MONTO NETO', width: 130 },
    //     { field: 'iva', headerName: 'IVA', width: 130 },
    //     { field: 'montoOperacionesExentas', headerName: 'MONTO OPERACIONES EXENTAS O NO GRAVADAS CON IVA', width: 250 },
    //     { field: 'montoTotal', headerName: 'MONTO TOTAL', width: 130 },
    //     { field: 'montoPercibido', headerName: 'MONTO PERCIBIDO O PAGADO', width: 200 },
    //     { field: 'glosaOperacion', headerName: 'GLOSA DE OPERACIÓN', width: 200 },
    //     { field: 'operacionEntidadRelacionada', headerName: 'OPERACIÓN CON ENTIDAD RELACIONADA', width: 250 },
    //     { field: 'percepcionOperacionDevengada', headerName: 'PERCEPCIÓN O PAGO PROVIENE DE OPERACIÓN DEVENGADA CON ANTERIORIDAD AL INGRESO AL RÉGIMEN SIMPLIFICADO O AL 31.12.2014', width: 600 },
    //     { field: 'operacionPagoPlazo', headerName: 'OPERACIÓN PACTADA CON PAGO A PLAZO', width: 250 },
    //     { field: 'fechaExigibilidadPago', headerName: 'FECHA DE EXIGIBILIDAD DEL PAGO', width: 250 },
    //     { field: 'montoIngreso', headerName: 'MONTO INGRESO', width: 130 },
    //     { field: 'montoEgreso', headerName: 'MONTO EGRESO', width: 130 },
    //     { field: 'saldo', headerName: 'SALDO', width: 130 },
    //     { field: 'montoTotalCompras', headerName: 'MONTO TOTAL COMPRAS', width: 150 },
    //     { field: 'montoTotalVentas', headerName: 'MONTO TOTAL VENTAS', width: 150 },
    // ];

    const columns = [
        { field: 'correlativo', headerName: 'Número correlativo', width: 130 },
        { field: 'tipoOperacion', headerName: 'TIPO OPERACIÓN (FLUJO INGRESO = 1; FLUJO EGRESO = 2)', width: 130 },
        { field: 'folio', headerName: 'N° de documento', width: 130 },
        { field: 'tipoDocumento', headerName: 'Tipo Documento', width: 130 },
        { field: 'rutEmisor', headerName: 'RUT EMISOR', width: 130 },
        { field: 'fechaOperacion', headerName: 'FECHA OPERACIÓN', width: 130 },
        { field: 'glosaOperacion', headerName: 'GLOSA DE OPERACIÓN', width: 200 },
        { field: 'glosaOperacion2', headerName: '0', width: 200 },
        { field: 'montoTotal', headerName: 'MONTO TOTAL', width: 130 },
        { field: 'montoNeto', headerName: 'MONTO NETO', width: 130 }
    ];

    const isValidDate = (date) => {
        return date instanceof Date && !isNaN(date);
    };

    const excelDateToJSDate = (serial) => {
        const utc_days = Math.floor(serial - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);

        const fractional_day = serial - Math.floor(serial) + 0.0000001;
        let total_seconds = Math.floor(86400 * fractional_day);

        const seconds = total_seconds % 60;
        total_seconds -= seconds;

        const hours = Math.floor(total_seconds / (60 * 60));
        const minutes = Math.floor(total_seconds / 60) % 60;

        return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
    };

    const formatDate = (date) => {
        if (!date) return '';

        if (!isNaN(date)) {
            const jsDate = excelDateToJSDate(parseFloat(date));
            if (isValidDate(jsDate)) {
                const day = String(jsDate.getDate()).padStart(2, '0');
                const month = String(jsDate.getMonth() + 1).padStart(2, '0');
                const year = jsDate.getFullYear();
                return `${day}/${month}/${year}`;
            }
        }

        const dateParts = date.split('/');
        if (dateParts.length === 3) {
            const day = dateParts[0];
            const month = dateParts[1];
            const year = dateParts[2];
            return `${day}/${month}/${year}`;
        }

        return date;
    };

    const processCombinedData = (combinedData) => {
        // Verificar que combinedData tenga las propiedades correctas y que sean arrays
        if (!Array.isArray(combinedData.enero) || !Array.isArray(combinedData.febrero) || !Array.isArray(combinedData.marzo) || !Array.isArray(combinedData.abril) || !Array.isArray(combinedData.mayo) || !Array.isArray(combinedData.junio) || !Array.isArray(combinedData.julio) || !Array.isArray(combinedData.agosto) || !Array.isArray(combinedData.septiembre) || !Array.isArray(combinedData.octubre) || !Array.isArray(combinedData.noviembre) || !Array.isArray(combinedData.diciembre)) {
            console.error('combinedData debe tener las propiedades los meses del año como arrays.');
            return {
                eneroSetRows: [],
                febreroSetRows: [],
                marzoSetRows: [],
                abrilSetRows: [],
                mayoSetRows: [],
                junioSetRows: [],
                julioSetRows: [],
                agostoSetRows: [],
                septiembreSetRows: [],
                octubreSetRows: [],
                noviembreSetRows: [],
                diciembreSetRows: [],
            };
        }

        const processSet = (dataSet) => {
            return dataSet
                .filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
        };

        // Procesar ambos conjuntos de datos en combinedData
        const eneroSetRows = processSet(combinedData.enero);
        const febreroSetRows = processSet(combinedData.febrero);
        const marzoSetRows = processSet(combinedData.marzo);
        const abrilSetRows = processSet(combinedData.abril);
        const mayoSetRows = processSet(combinedData.mayo);
        const junioSetRows = processSet(combinedData.junio);
        const julioSetRows = processSet(combinedData.julio);
        const agostoSetRows = processSet(combinedData.agosto);
        const septiembreSetRows = processSet(combinedData.septiembre);
        const octubreSetRows = processSet(combinedData.octubre);
        const noviembreSetRows = processSet(combinedData.noviembre);
        const diciembreSetRows = processSet(combinedData.diciembre);

        return {
            eneroSetRows,
            febreroSetRows,
            marzoSetRows,
            abrilSetRows,
            mayoSetRows,
            junioSetRows,
            julioSetRows,
            agostoSetRows,
            septiembreSetRows,
            octubreSetRows,
            noviembreSetRows,
            diciembreSetRows
        };
    };


    const { eneroSetRows,
        febreroSetRows,
        marzoSetRows,
        abrilSetRows,
        mayoSetRows,
        junioSetRows,
        julioSetRows,
        agostoSetRows,
        septiembreSetRows,
        octubreSetRows,
        noviembreSetRows,
        diciembreSetRows } = processCombinedData(combinedData);

    const handleBackButton = () => {
        navigate('/display');
    };

    const handleInputChange = (e) => {
        const { id, value } = e.target;
        setUserInput((prevState) => ({
            ...prevState,
            [id]: id === 'initialBalance' ? parseFloat(value) : value
        }));
    };



    const handleExport = async () => {
        const { year, rutUser, name, initialBalance } = userInput;
        setOpen(false)
        const workbook = new Workbook();
        const worksheet1 = workbook.addWorksheet('Enero');
        // worksheet1.views = [
        //   { state: 'frozen', ySplit: 11 } // Inmoviliza las primeras 11 filas
        // ];

        // Crear segunda hoja (Planilla 2)
        const worksheet2 = workbook.addWorksheet('Febrero');

        const worksheet3 = workbook.addWorksheet('Marzo');

        const worksheet4 = workbook.addWorksheet('Abril');

        const worksheet5 = workbook.addWorksheet('Mayo');

        const worksheet6 = workbook.addWorksheet('Junio');

        const worksheet7 = workbook.addWorksheet('Julio');

        const worksheet8 = workbook.addWorksheet('Agosto');

        const worksheet9 = workbook.addWorksheet('Septiembre');

        const worksheet10 = workbook.addWorksheet('Octubre');

        const worksheet11 = workbook.addWorksheet('Noviembre');

        const worksheet12 = workbook.addWorksheet('Diciembre');

        const mesesExcel = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

        const applyFormating = (worksheet, data, saldoInicial, i, mes) => {
            const ws_data = [];

            //   for (let i = 0; i < 3; i++) {
            //     ws_data.push([]);
            //   }

            ws_data.push(columns.map(col => col.headerName));

            // Filtra los arrays y da formato a las fechas de los sets de datos por mes
            const filteredDataArray1 = eneroSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray2 = febreroSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray3 = marzoSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray4 = abrilSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray5 = mayoSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray6 = junioSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray7 = julioSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray8 = agostoSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray9 = septiembreSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray10 = octubreSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray11 = noviembreSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));
            const filteredDataArray12 = diciembreSetRows.filter(row => Object.values(row).some(cell => cell !== null && cell !== ''))
                .map((row, index) => ({
                    id: index,
                    ...row,
                    fechaOperacion: formatDate(row.fechaOperacion),
                }));

            // Combina los arrays filtrados en un solo array
            const filteredData = [filteredDataArray1, filteredDataArray2, filteredDataArray3, filteredDataArray4, filteredDataArray5, filteredDataArray6, filteredDataArray7, filteredDataArray8, filteredDataArray9, filteredDataArray10, filteredDataArray11, filteredDataArray12];
            console.log(filteredData)


            ws_data.forEach((row, index) => {
                worksheet.addRow(row);
                if (index === 0) {
                    worksheet.mergeCells('C2:M2');
                    worksheet.getCell('C2').font = { bold: true, size: 16 };
                    worksheet.getCell('C2').alignment = { horizontal: 'center' };
                }
            });

            worksheet.getCell('C2').value = 'ANEXO 3. LIBRO DE CAJA CONTRIBUYENTES ACOGIDOS AL RÉGIMEN DEL ARTÍCULO 14 LETRA D) DEL N°3 Y N°8 LETRA (a) DE LA LEY SOBRE IMPUESTO A LA RENTA'
            worksheet.getCell('C2').alignment = {
                vertical: 'middle', // Alinea el contenido verticalmente en el medio
                horizontal: 'center', // Alinea el contenido horizontalmente en el centro
                wrapText: true // Ajusta el texto si es necesario para ocupar toda la celda
            };
            worksheet.getCell('C2').font = { bold: true, size: 11 }



            const borderStyle = { style: 'thin' };
            const borderStyleOutside = { style: 'thick' };




            worksheet.getCell('M4').border = { right: borderStyle }
            worksheet.getCell('M5').border = { right: borderStyle }
            worksheet.getCell('M6').border = { right: borderStyle }
            worksheet.getCell('M7').border = { right: borderStyle }
            worksheet.getCell('M8').border = { right: borderStyle }
            worksheet.getCell('M9').border = { right: borderStyle }
            worksheet.getCell('M10').border = { right: borderStyle }
            worksheet.getCell('M11').border = { right: borderStyle }
            worksheet.getCell('M12').border = { right: borderStyle }

            // Mantener las primeras 12 filas intactas (Formato nuevo)

            worksheet.mergeCells('C8:D8');
            worksheet.mergeCells('D4:E4');
            worksheet.mergeCells('D6:E6');
            worksheet.mergeCells('E8:L8');

            worksheet.getCell('E8').value = name
            worksheet.getCell('E8').alignment = { horizontal: 'center' };
            worksheet.getCell('E8').font = { bold: true, size: 12 }
            worksheet.getCell('E8').border = {
                left: borderStyle,
                top: borderStyle,
                bottom: borderStyle,
                right: borderStyle
            }


            //Parte superior
            worksheet.getCell('B2').border = {
                top: borderStyleOutside,
                left: borderStyleOutside
            }
            // worksheet.mergeCells('C2:M2');

            worksheet.getCell('C2:M2').border = {
                top: borderStyleOutside,
                right: borderStyleOutside
            }
            //borde izq
            for (let i = 3; i <= 12; i++) {
                worksheet.getCell(`B${i}`).border = {
                    left: borderStyleOutside
                }
            }
            //borde der
            for (let i = 3; i <= 12; i++) {
                worksheet.getCell(`M${i}`).border = {
                    right: borderStyleOutside
                }
            }

            //celdas con los valores de userData
            worksheet.getCell('C4').value = 'PERÍODO';
            worksheet.getCell('C6').value = 'RUT';
            worksheet.getCell('C8').value = 'NOMBRE/RAZÓN SOCIAL';

            for (let i = 4; i <= 8; i += 2) {
                worksheet.getCell(`C${i}`).border = { top: borderStyle, bottom: borderStyle, right: borderStyle, left: borderStyle }
                worksheet.getCell(`C${i}`).font = { bold: true, size: 10 }
                worksheet.getCell(`D${i}`).border = { top: borderStyle, bottom: borderStyle, right: borderStyle, left: borderStyle }
                worksheet.getCell(`D${i}`).font = { bold: true, size: 10 }
            }


            //Encabezado
            worksheet.mergeCells('C10:L10');
            worksheet.getCell('C10').value = 'REGISTRO DE OPERACIONES'
            worksheet.getCell('C10').font = { bold: true, size: 10 }
            worksheet.getCell('C10').alignment = { horizontal: 'center' };
            worksheet.getCell('C10').border = {
                left: borderStyleOutside,
                top: borderStyleOutside,
                bottom: borderStyleOutside,
                right: borderStyleOutside
            }

            worksheet.getCell('C11').value = 'N° CORRELATIVO'
            worksheet.getCell('D11').value = 'NTIPO OPERACIÓN (FLUJO INGRESO = 1; FLUJO EGRESO = 2)'
            worksheet.getCell('E11').value = 'N° DE DOCUMENTO'
            worksheet.getCell('F11').value = 'TIPO DOCUMENTO'
            worksheet.getCell('G11').value = 'RUT EMISOR'
            worksheet.getCell('H11').value = 'FECHA OPERACIÓN'
            worksheet.mergeCells('I11:J12');
            worksheet.getCell('I11').value = 'GLOSA DE OPERACIÓN'
            worksheet.getCell('I11').font = { bold: true, size: 10 }
            worksheet.getCell('I11').alignment = { horizontal: 'center' };
            worksheet.getCell('I11').border = {
                left: borderStyle,
                top: borderStyleOutside,
                bottom: borderStyle,
                right: borderStyle
            }
            worksheet.getCell(`I11`).alignment = {
                vertical: 'middle', // Alinea el contenido verticalmente en el medio
                horizontal: 'center', // Alinea el contenido horizontalmente en el centro
                wrapText: true // Ajusta el texto si es necesario para ocupar toda la celda
            };
            worksheet.getCell('K11').value = 'MONTO TOTAL FLUJO DE INGRESO O EGRESO'
            worksheet.getCell('L11').value = 'MONTO QUE AFECTA LA BASE IMPONIBLE'


            const columnsToFormat = ['C', 'D', 'E', 'F', 'G', 'H', 'K', 'L'];

            for (let i = 0; i < columnsToFormat.length; i++) {
                worksheet.mergeCells(`${columnsToFormat[i]}11:${columnsToFormat[i]}12`);
                worksheet.getCell(`${columnsToFormat[i]}11`).border = {
                    left: borderStyle,
                    top: borderStyleOutside,
                    bottom: borderStyle,
                    right: borderStyle
                }
                worksheet.getCell(`${columnsToFormat[i]}11`).font = { bold: true, size: 10 }
                worksheet.getCell(`${columnsToFormat[i]}11`).alignment = { horizontal: 'center' };
                worksheet.getCell(`${columnsToFormat[i]}11`).alignment = {
                    vertical: 'middle', // Alinea el contenido verticalmente en el medio
                    horizontal: 'center', // Alinea el contenido horizontalmente en el centro
                    wrapText: true // Ajusta el texto si es necesario para ocupar toda la celda
                };

            }

            worksheet.getCell('C11').border = { left: borderStyleOutside, bottom: borderStyle }
            worksheet.getCell('L11').border = { right: borderStyleOutside, bottom: borderStyle }


            //Ancho Columnas
            worksheet.getColumn(1).width = 2;   // Columna A
            worksheet.getColumn(2).width = 2;  // Columna B
            worksheet.getColumn(3).width = 16;  // Columna C
            worksheet.getColumn(4).width = 16;  // Columna D
            worksheet.getColumn(5).width = 16;  // Columna E
            worksheet.getColumn(6).width = 13;  // Columna F
            worksheet.getColumn(7).width = 13;  // Columna G
            worksheet.getColumn(8).width = 13;  // Columna H
            worksheet.getColumn(9).width = 15;  // Columna I
            worksheet.getColumn(10).width = 16; // Columna J
            worksheet.getColumn(11).width = 20; // Columna K
            worksheet.getColumn(12).width = 16; // Columna L
            worksheet.getColumn(13).width = 2; // Columna M
            worksheet.getColumn(14).width = 9; // Columna N
            worksheet.getColumn(15).width = 13; // Columna O
            worksheet.getColumn(16).width = 13; // Columna P
            worksheet.getColumn(17).width = 13; // Columna Q
            worksheet.getColumn(18).width = 13; // Columna R

            //Alto Filas
            worksheet.getRow(1).height = 12;
            worksheet.getRow(2).height = 42;
            worksheet.getRow(3).height = 17;
            worksheet.getRow(5).height = 10;
            worksheet.getRow(7).height = 10;
            worksheet.getRow(10).height = 20;
            worksheet.getRow(11).height = 38;
            worksheet.getRow(12).height = 38;


            worksheet.getCell('F10').alignment = {
                vertical: 'middle',      // Alinea el contenido verticalmente en la parte inferior
                horizontal: 'center',    // Alinea el contenido horizontalmente en el centro
                shrinkToFit: true,       // Ajusta el tamaño del texto para que quepa en la celda
                wrapText: false          // Desactiva la opción de "Ajustar texto"
            };
            const startRow = 13;



            // console.log(filteredData[i]);
            //Aqui las variables que suman los valores de las planillas finales

            let egresosTotales = 0;

            //Se genera la tabla y se entregan los valores a las filas
            worksheet.addTable({
                name: 'CombinedDataTable',
                ref: `C${startRow}`,
                headerRow: false,
                columns: columns.map(col => ({ name: col.headerName })), // Usar columnas filtradas
                rows: [
                    // Agregar una fila vacía con "SALDO INICIAL" en la columna D (índice 4)
                    columns.map((col, index) => index === 1 ? 0 : (index === 0 ? 1 : '')),
                    // Agregar las filas de datos
                    ...filteredData[i].map((row, rowIndex) => columns.map(col => {
                        if (col.field === 'fechaOperacion' || col.field === 'fechaExigibilidadPago') {
                            return formatDate(row[col.field]);
                        } else if (col.field === 'montoIngreso') {
                            return Number(row['montoTotalVentas']).toLocaleString('es-CL', { minimumFractionDigits: 0 });
                        } else if (col.field === 'montoEgreso') {
                            return Number(row['montoTotalCompras']).toLocaleString('es-CL', { minimumFractionDigits: 0 });
                        } else if (col.field === 'glosaOperacion') {
                            if (row['glosaOperacion'] === 'COMPRA AF IVA') {
                                egresosTotales += Number(row['montoNeto']); // Acumular la suma correctamente
                                console.log(`La suma del valor actual más valor neto es: ${egresosTotales}`);
                                console.log(row['montoNeto']);
                            }
                            // console.log(row['glosaOperacion']);
                            return row['glosaOperacion'];
                        }
                        return col.field === 'correlativo' ? rowIndex + 2 : row[col.field];
                    }))
                ],
                style: {
                    theme: 'TableStyleMedium2',
                    showRowStripes: true,
                }
            });

            console.log(`Egresos Totales Finales: ${egresosTotales}`);


            //Se unen las celdas de las columnas I y J (9 y 10 en Excl)
            filteredData[i].forEach((_, rowIndex) => {
                const startRowForMerge = startRow + rowIndex;

                worksheet.mergeCells(`I${startRowForMerge}:J${startRowForMerge}`);
                //Se agregan los bordes exteriores de la tabla (columna B por la izq, y M por la derecha)
                worksheet.getCell(`B${startRowForMerge}`).border = { left: borderStyleOutside, right: borderStyleOutside };
                worksheet.getCell(`M${startRowForMerge}`).border = { left: borderStyleOutside, right: borderStyleOutside };

                //Los dos siguientes bucles for son para arreglar las filas que no se estilan correctamente
                // Aplicar estilos de borde a las celdas de las columnas B y M en las siguientes n filas después del bucle
                for (let i = 1; i <= 4; i++) {
                    const currentRow = startRowForMerge + i;

                    worksheet.getCell(`B${currentRow}`).border = { left: borderStyleOutside };
                    worksheet.getCell(`M${currentRow}`).border = { right: borderStyleOutside };
                }

            });


            //Se le da estilo a las cceldas desde la 13 en adelante
            worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
                if (rowNumber >= 13) {
                    row.eachCell((cell) => {
                        cell.border = {
                            top: { style: 'thin' },
                            left: { style: 'thin' },
                            bottom: { style: 'thin' },
                            right: { style: 'thin' }
                        };
                        cell.alignment = { horizontal: 'right' };
                    });
                }

            });

            //Aqui comienza todo lo relacionado a el formato y a la segunda tabla, posterior al final de la primera tabla desde ending table en su posicion i

            const endingTablei = filteredData[i].length;
            console.log(endingTablei);
            // Arreglar las últimas 2 filas de la planilla con mergeCells
            worksheet.mergeCells(`I${startRow + endingTablei}:J${startRow + endingTablei}`)
            console.log(startRow + endingTablei);
            worksheet.getCell(`I${startRow + endingTablei}`).border = { bottom: borderStyleOutside }

            for (let i = 0; i < columnsToFormat.length; i++) {
                worksheet.getCell(`${columnsToFormat[i]}${startRow + endingTablei}`).border = { bottom: borderStyleOutside }
            }

            worksheet.getCell(`C${startRow + endingTablei}`).border = { left: borderStyleOutside, bottom: borderStyleOutside }
            worksheet.getCell(`L${startRow + endingTablei}`).border = { right: borderStyleOutside, bottom: borderStyleOutside }
            //-----------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------
            //-----------------------------------------------------------------------------------------

            //Inicio de la tabla 2 de resumenes
            const table2begining = startRow + endingTablei + 4;
            const doubleBorder = { style: 'double' }

            worksheet.mergeCells(`C${table2begining}:H${table2begining}`)
            worksheet.getCell(`C${table2begining}`).value = 'SALDOS Y TOTALES LIBRO DE CAJA';
            worksheet.getCell(`C${table2begining}`).border = {
                top: borderStyleOutside,
                left: borderStyleOutside,
                bottom: borderStyleOutside,
                right: borderStyleOutside
            };
            worksheet.getRow(table2begining).height = 35;
            worksheet.getCell(`C${table2begining}`).font = { bold: true, size: 12 };
            worksheet.getCell(`C${table2begining}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };


            worksheet.mergeCells(`C${table2begining + 1}:E${table2begining + 1}`)
            worksheet.getCell(`C${table2begining + 1}`).value = 'FLUJO DE INGRESOS Y EGRESOS';
            worksheet.getCell(`C${table2begining + 1}`).border = {
                left: borderStyleOutside,
                right: doubleBorder,
                bottom: borderStyle
            };

            worksheet.mergeCells(`F${table2begining + 1}:H${table2begining + 1}`)
            worksheet.getCell(`F${table2begining + 1}`).value = 'MONTOS QUE AFECTAN LA BASE IMPONIBLE';
            worksheet.getCell(`F${table2begining + 1}`).border = {
                right: borderStyleOutside,
                bottom: borderStyle
            };
            worksheet.getRow(table2begining + 1).height = 35;


            worksheet.getCell(`C${table2begining + 2}`).value = 'TOTAL MONTO FLUJO DE INGRESOS';
            worksheet.getCell(`C${table2begining + 2}`).border = {
                left: borderStyleOutside
            };

            worksheet.getCell(`D${table2begining + 2}`).value = 'TOTAL MONTO FLUJO DE EGRESOS';
            worksheet.getCell(`D${table2begining + 2}`).border = {
                left: borderStyle,
            };

            worksheet.getCell(`E${table2begining + 2}`).value = 'SALDO FLUJO DE CAJA';
            worksheet.getCell(`E${table2begining + 2}`).border = {
                left: borderStyle,
                right: doubleBorder
            };

            worksheet.getCell(`F${table2begining + 2}`).value = 'INGRESOS';
            worksheet.getCell(`F${table2begining + 2}`).border = {
                right: borderStyle,
            };

            worksheet.getCell(`G${table2begining + 2}`).value = 'EGRESOS';
            worksheet.getCell(`G${table2begining + 2}`).border = {
                right: borderStyle,
            };

            worksheet.getCell(`H${table2begining + 2}`).value = 'RESULTADO NETO';
            worksheet.getCell(`H${table2begining + 2}`).border = {
                right: borderStyleOutside,
            };

            worksheet.getRow(table2begining + 2).height = 60;
            worksheet.getRow(table2begining + 3).height = 25;


            for (let i = 0; i <= columnsToFormat.length - 3; i++) {
                console.log(`${columnsToFormat[i]}${table2begining + 3}`);

                worksheet.getCell(`${columnsToFormat[i]}${table2begining + 3}`).border = {
                    top: borderStyle,
                    left: borderStyle,
                    bottom: borderStyleOutside,
                    right: borderStyle
                }
            }
            worksheet.getCell(`C${table2begining + 3}`).border = { left: borderStyleOutside, bottom: borderStyleOutside, top: borderStyle }
            worksheet.getCell(`H${table2begining + 3}`).border = { right: borderStyleOutside, bottom: borderStyleOutside, top: borderStyle }


            // Aplicar estilos de borde a las celdas de las columnas B y M en después de la primera tabla
            const columnsToFormatEnding = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'K', 'L', 'M'];

            for (let i = 0; i <= 8; i++) {
                const currentRow = table2begining + i;

                worksheet.getCell(`B${currentRow}`).border = { left: borderStyleOutside };
                worksheet.getCell(`M${currentRow}`).border = { right: borderStyleOutside };

                if (i === 8) {
                    for (let i = 0; i < columnsToFormatEnding.length; i++) {
                        worksheet.getCell(`${columnsToFormatEnding[i]}${currentRow}`).border = { bottom: borderStyleOutside }
                    }

                    worksheet.getCell(`B${currentRow}`).border = { left: borderStyleOutside, bottom: borderStyleOutside }
                    worksheet.getCell(`M${currentRow}`).border = { right: borderStyleOutside, bottom: borderStyleOutside }
                }
            }

            worksheet.getCell(`I${table2begining + 8}`).border = { bottom: borderStyleOutside }
            worksheet.getCell(`J${table2begining + 8}`).border = { bottom: borderStyleOutside }


            //alineado y estilos de la tabla 2
            const cells = [
                `C${table2begining + 1}`,
                `F${table2begining + 1}`,
                `C${table2begining + 2}`,
                `D${table2begining + 2}`,
                `E${table2begining + 2}`,
                `F${table2begining + 2}`,
                `G${table2begining + 2}`,
                `H${table2begining + 2}`,
            ];

            cells.forEach(cell => {
                worksheet.getCell(cell).font = { bold: true, size: 10 };
                worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            });

            //   const firstSaldoCell = 12;
            //   worksheet.getCell(`R${firstSaldoCell}`).value = saldoInicial;
            //   worksheet.getCell(`R${firstSaldoCell}`).numFmt = '#,##0';

            //   for (let i = firstSaldoCell + 1; i <= worksheet.rowCount; i++) {
            //     worksheet.getCell(`R${i}`).value = {
            //         formula: `R${i-1}+P${i}-Q${i}+IF(N${i}="",0,N${i})`
            //     };
            //     worksheet.getCell(`R${i}`).numFmt = '#,##0';  // Aplica el formato de números con separación por miles
            // }

            // for (let i = firstSaldoCell; i <= worksheet.rowCount; i++) {
            //   const saldoCell = worksheet.getCell(`R${i}`);
            //   saldoCell.value = { formula: `R${i-1}+P${i}-Q${i}` };
            //   saldoCell.numFmt = '#.##0';  // Aplica el formato de números con separación por miles
            // }


        }

        // Procesar los datos

        // Aplicar formato y datos a las dos hojas
        applyFormating(worksheet1, eneroSetRows, initialBalance, 0, mesesExcel[0]);
        // // Obtener la columna 'R' sin incluir el encabezado vacío
        // const columnRValues = worksheet1.getColumn('R').values.filter(value => value !== null && value !== undefined);
        // // Obtener el valor de la última celda no vacía en la columna 'R'
        // console.dir(columnRValues[columnRValues.length - 1].formula.slice(-5).match(/N(\d+)\)/), { depth: null });
        // let extractedNumber = columnRValues[columnRValues.length - 1].formula.slice(-5).match(/N(\d+)\)/) ? parseInt(columnRValues[columnRValues.length - 1].formula.slice(-5).match(/N(\d+)\)/)[1], 10) : columnRValues[columnRValues.length - 1].formula.slice(-5).match(/N(\d+)\)/)[1]
        // const saldoFebrero = { formula: `Enero!R${extractedNumber}` };
        // console.dir(saldoFebrero, { depth: null });

        // const saldoFebrero = { formula: '=BUSCAR(2,1/(Enero!R:R<>""),Enero!R:R)' };    
        const saldoFebrero = { formula: '=LOOKUP(2,1/(Enero!R:R<>""),Enero!R:R)' };
        applyFormating(worksheet2, febreroSetRows, saldoFebrero, 1, mesesExcel[1]);

        const saldoMarzo = { formula: '=LOOKUP(2,1/(Febrero!R:R<>""),Febrero!R:R)' };
        applyFormating(worksheet3, marzoSetRows, saldoMarzo, 2, mesesExcel[2]);

        const saldoAbril = { formula: '=LOOKUP(2,1/(Marzo!R:R<>""),Marzo!R:R)' };
        applyFormating(worksheet4, abrilSetRows, saldoAbril, 3, mesesExcel[3]);

        const saldoMayo = { formula: '=LOOKUP(2,1/(Abril!R:R<>""),Abril!R:R)' };
        applyFormating(worksheet5, mayoSetRows, saldoMayo, 4, mesesExcel[4]);

        const saldoJunio = { formula: '=LOOKUP(2,1/(Mayo!R:R<>""),Mayo!R:R)' };
        applyFormating(worksheet6, junioSetRows, saldoJunio, 5, mesesExcel[5]);

        const saldoJulio = { formula: '=LOOKUP(2,1/(Junio!R:R<>""),Junio!R:R)' };
        applyFormating(worksheet7, julioSetRows, saldoJulio, 6, mesesExcel[6]);

        const saldoAgosto = { formula: '=LOOKUP(2,1/(Julio!R:R<>""),Julio!R:R)' };
        applyFormating(worksheet8, agostoSetRows, saldoAgosto, 7, mesesExcel[7]);

        const saldoSeptiembre = { formula: '=LOOKUP(2,1/(Agosto!R:R<>""),Agosto!R:R)' };
        applyFormating(worksheet9, septiembreSetRows, saldoSeptiembre, 8, mesesExcel[8]);

        const saldoOctubre = { formula: '=LOOKUP(2,1/(Septiembre!R:R<>""),Septiembre!R:R)' };
        applyFormating(worksheet10, octubreSetRows, saldoOctubre, 9, mesesExcel[9]);

        const saldoNoviembre = { formula: '=LOOKUP(2,1/(Octubre!R:R<>""),Octubre!R:R)' };
        applyFormating(worksheet11, noviembreSetRows, saldoNoviembre, 10, mesesExcel[10]);

        const saldoDiciembre = { formula: '=LOOKUP(2,1/(Noviembre!R:R<>""),Noviembre!R:R)' };
        applyFormating(worksheet12, diciembreSetRows, saldoDiciembre, 11, mesesExcel[11]);

        const planillas = [worksheet1, worksheet2, worksheet3, worksheet4, worksheet5, worksheet6, worksheet7, worksheet8, worksheet9, worksheet10, worksheet11, worksheet12];

        planillas.forEach((worksheet) => {
            worksheet.pageSetup = {
                orientation: 'landscape', // Configura la página en horizontal
                fitToPage: true,           // Ajusta el contenido a una sola página
                fitToHeight: 2,            // Número máximo de páginas verticales
                fitToWidth: 1,             // Número máximo de páginas horizontales
                paperSize: undefined,              // Tamaño de papel A4 (puede variar según la región)
                printArea: `=OFFSET(worksheet!$A$1, 0, 0, COUNTA(worksheet!$A:$A), COUNTA(worksheet!$1:$1))` // Define el área de impresión
            };
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/octet-stream' });
        saveAs(blob, 'Planilla_Combinada.xlsx');
    };

    return (
        <Box sx={{ padding: 4 }}>
            <Box sx={{ display: 'flex', justifyContent: 'center', marginTop: 2 }}>
                <Button variant="contained" color="primary" onClick={handleOpen}>
                    Exportar a Excel
                </Button>

                <Modal open={open} onClose={handleClose}>
                    <Box
                        sx={{
                            position: 'absolute',
                            top: '50%',
                            left: '50%',
                            transform: 'translate(-50%, -50%)',
                            width: 400,
                            bgcolor: 'background.paper',
                            border: '2px solid #000',
                            boxShadow: 24,
                            p: 4,
                        }}
                    >
                        <h2>Ingrese la información Libro Caja</h2>
                        <TextField
                            id="year"
                            label="Año"
                            fullWidth
                            margin="normal"
                            value={userInput.year}
                            onChange={handleInputChange}
                        />
                        <TextField
                            id="rutUser"
                            label="RUT"
                            fullWidth
                            margin="normal"
                            value={userInput.rutUser}
                            onChange={handleInputChange}
                        />
                        <TextField
                            id="name"
                            label="Nombre"
                            fullWidth
                            margin="normal"
                            value={userInput.name}
                            onChange={handleInputChange}
                        />
                        <TextField
                            id="initialBalance"
                            label="Saldo inicial"
                            type="number"
                            fullWidth
                            margin="normal"
                            value={userInput.initialBalance}
                            onChange={handleInputChange}
                        />
                        <Button variant="contained" onClick={handleExport}>
                            Save Data
                        </Button>
                    </Box>
                </Modal>
            </Box>
            <Button variant="contained" color="primary" onClick={handleBackButton}>
                Volver
            </Button>
            <Typography variant="h4" gutterBottom>
                Planilla Combinada
            </Typography>
            <Box sx={{ height: 600, width: '100%' }}>
                <DataGrid
                    rows={eneroSetRows}
                    columns={columns}
                    initialState={{
                        pagination: { paginationModel: { page: 0, pageSize: 10 } },
                    }}
                    pageSizeOptions={[10, 20]}
                    checkboxSelection
                />
            </Box>
            <Box sx={{ height: 600, width: '100%' }}>
                <DataGrid
                    rows={febreroSetRows}
                    columns={columns}
                    initialState={{
                        pagination: { paginationModel: { page: 0, pageSize: 10 } },
                    }}
                    pageSizeOptions={[10, 20]}
                    checkboxSelection
                />
            </Box>
        </Box>
    );
};

export default Combined;