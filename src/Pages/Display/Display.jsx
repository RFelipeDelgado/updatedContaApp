/* eslint-disable react/prop-types */
import { DataGrid } from '@mui/x-data-grid';
import { Button, Box } from '@mui/material';
import { useNavigate } from 'react-router-dom';

const Display = ({ filesData, setCombinedData }) => {
  const navigate = useNavigate();

  const generateColumns = (data) => {
    return data[0].map((header, index) => ({
      field: `column${index}`,
      headerName: header,
      width: 150,
    }));
  };

  const isValidDate = (date) => {
    return date instanceof Date && !isNaN(date);
  };

  const excelDateToJSDate = (serial) => {
    const excelEpoch = new Date(1899, 11, 30); // 30th Dec 1899
    const daysSinceExcelEpoch = Math.floor(serial);
    const millisecondsSinceExcelEpoch = daysSinceExcelEpoch * 24 * 60 * 60 * 1000;

    const fractionalDay = serial - daysSinceExcelEpoch;
    const millisecondsInDay = Math.round(fractionalDay * 24 * 60 * 60 * 1000);

    const date = new Date(excelEpoch.getTime() + millisecondsSinceExcelEpoch + millisecondsInDay);
    return date;
  };

  const formatDate = (date) => {
    if (!date) return '';

    if (!isNaN(date)) {
      // Handle Excel serial date
      const jsDate = excelDateToJSDate(parseFloat(date));
      if (isValidDate(jsDate)) {
        const day = String(jsDate.getDate()).padStart(2, '0');
        const month = String(jsDate.getMonth() + 1).padStart(2, '0');
        const year = jsDate.getFullYear();
        return `${month}/${day}/${year}`;
      }
    }

    // Handle date in format dd/mm/yyyy
    const dateParts = date.split('/');
    if (dateParts.length === 3) {
      const day = dateParts[0];
      const month = dateParts[1];
      const year = dateParts[2];
      return `${day}/${month}/${year}`;
    }

    return date;
  };

  const generateRows = (data) => {
    return data.slice(1)
      .filter(row => row.some(cell => cell !== null && cell !== ''))
      .map((row, index) => {
        const rowData = {};
        row.forEach((cell, cellIndex) => {
          if (data[0][cellIndex].toLowerCase().includes('fecha docto')) {
            const originalDate = row[data[0].indexOf('Fecha Docto')];
            const formattedDate = formatDate(originalDate);
            // console.log(`Original Date: ${originalDate}, Formatted Date: ${formattedDate}`);
            rowData[`column${cellIndex}`] = formattedDate;
          } else {
            rowData[`column${cellIndex}`] = cell;
          }
        });
        return { id: index, ...rowData };
      });
  };

  const columns1 = generateColumns(filesData[0]);
  const rows1 = generateRows(filesData[0]);

  const columns2 = generateColumns(filesData[1]);
  const rows2 = generateRows(filesData[1]);

  const columns3 = generateColumns(filesData[2]);
  const rows3 = generateRows(filesData[2]);

  const columns4 = generateColumns(filesData[3]);
  const rows4 = generateRows(filesData[3]);

  const columns5 = generateColumns(filesData[4]);
  const rows5 = generateRows(filesData[4]);

  const columns6 = generateColumns(filesData[5]);
  const rows6 = generateRows(filesData[5]);

  const columns7 = generateColumns(filesData[6]);
  const rows7 = generateRows(filesData[6]);

  const columns8 = generateColumns(filesData[7]);
  const rows8 = generateRows(filesData[7]);

  const columns9 = generateColumns(filesData[8]);
  const rows9 = generateRows(filesData[8]);

  const columns10 = generateColumns(filesData[9]);
  const rows10 = generateRows(filesData[9]);

  const columns11 = generateColumns(filesData[10]);
  const rows11 = generateRows(filesData[10]);

  const columns12 = generateColumns(filesData[11]);
  const rows12 = generateRows(filesData[11]);

  const columns13 = generateColumns(filesData[12]);
  const rows13 = generateRows(filesData[12]);

  const columns14 = generateColumns(filesData[13]);
  const rows14 = generateRows(filesData[13]);

  const columns15 = generateColumns(filesData[14]);
  const rows15 = generateRows(filesData[14]);

  const columns16 = generateColumns(filesData[15]);
  const rows16 = generateRows(filesData[15]);

  const columns17 = generateColumns(filesData[16]);
  const rows17 = generateRows(filesData[16]);

  const columns18 = generateColumns(filesData[17]);
  const rows18 = generateRows(filesData[17]);

  const columns19 = generateColumns(filesData[18]);
  const rows19 = generateRows(filesData[18]);

  const columns20 = generateColumns(filesData[19]);
  const rows20 = generateRows(filesData[19]);

  const columns21 = generateColumns(filesData[20]);
  const rows21 = generateRows(filesData[20]);

  const columns22 = generateColumns(filesData[21]);
  const rows22 = generateRows(filesData[21]);

  const columns23 = generateColumns(filesData[22]);
  const rows23 = generateRows(filesData[22]);

  const columns24 = generateColumns(filesData[23]);
  const rows24 = generateRows(filesData[23]);



  const formatNumber = (number) => {
    return new Intl.NumberFormat('de-DE').format(number);
  };



  const handleCombineSheets = () => {
    const combinedData = {
      enero: [],
      febrero: [],
      marzo: [],
      abril: [],
      mayo: [],
      junio: [],
      julio: [],
      agosto: [],
      septiembre: [],
      octubre: [],
      noviembre: [],
      diciembre: [],
    };

    // const headers = ['correlativo', 'folio', 'tipoDocumento', 'rutEmisor', 'fechaOperacion', 'montoNeto', 'iva',
    //   'montoOperacionesExentas', 'montoTotal', 'montoPercibido', 'glosaOperacion', 'operacionEntidadRelacionada',
    //   'percepcionOperacionDevengada', 'operacionPagoPlazo', 'fechaExigibilidadPago', 'montoIngreso', 'montoEgreso', 'saldo'];

    filesData.slice(0, 2).forEach((data, fileIndex) => {
      console.log(fileIndex)
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.enero.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 0 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 1 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.enero.push(combinedRow);
      });
    });

    filesData.slice(2, 4).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.febrero.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        console.log('Index of RUT Proveedor/Rut cliente:', data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente'));
        console.log('Value of rutEmisor:', row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')]);
        combinedData.febrero.push(combinedRow);
      });
    });

    filesData.slice(4, 6).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.marzo.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.marzo.push(combinedRow);
      });
    });

    filesData.slice(6, 8).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.abril.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.abril.push(combinedRow);
      });
    });

    filesData.slice(8, 10).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.mayo.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.mayo.push(combinedRow);
      });
    });

    filesData.slice(10, 12).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.junio.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.junio.push(combinedRow);
      });
    });

    filesData.slice(12, 14).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.julio.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.julio.push(combinedRow);
      });
    });

    filesData.slice(14, 16).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.agosto.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.agosto.push(combinedRow);
      });
    });

    filesData.slice(16, 18).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.septiembre.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.septiembre.push(combinedRow);
      });
    });

    filesData.slice(18, 20).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.octubre.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.octubre.push(combinedRow);
      });
    });

    filesData.slice(20, 22).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.noviembre.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.noviembre.push(combinedRow);
      });
    });

    filesData.slice(22, 24).forEach((data, fileIndex) => {
      data.slice(1).forEach((row) => {
        const montoNeto = row[data[0].indexOf('Monto Neto')] || 0;
        const montoIva = row[data[0].indexOf('Monto IVA Recuperable')] || row[data[0].indexOf('Monto IVA')] || 0;
        const montoTotal = row[data[0].indexOf('Monto Total')] || row[data[0].indexOf('Monto total')] || 0;

        let indexMontoTotal = data[0].indexOf('Monto Total');
        let indexMontoTotalLower = data[0].indexOf('Monto total');

        let glosa;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          glosa = "COMPRA AF IVA";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          glosa = "VENTA AF IVA";
        } else {
          glosa = 0;
        }

        let tipoOperacion1o2;
        if (indexMontoTotal !== -1 && row[indexMontoTotal]) {
          tipoOperacion1o2 = "2";
        } else if (indexMontoTotalLower !== -1 && row[indexMontoTotalLower]) {
          tipoOperacion1o2 = "1";
        } else {
          tipoOperacion1o2 = 0;
        }

        let tipoDocumento = row[data[0].indexOf('Tipo Doc')];
        if (tipoDocumento === 33) {
          tipoDocumento = 'FACTURA';
        } else if (tipoDocumento === 34) {
          tipoDocumento = 'FACTURA EXENTA';
        } else if (tipoDocumento === 61) {
          tipoDocumento = 'NOTA CREDITO';
        }

        let montoTotalCompras = row[data[0].indexOf('Monto Total')] || 0;
        let montoTotalVentas = row[data[0].indexOf('Monto total')] || 0;

        if (tipoDocumento === 'NOTA CREDITO') {
          montoTotalCompras = -Math.abs(montoTotalCompras);
          montoTotalVentas = -Math.abs(montoTotalVentas);
        } else {
          montoTotalCompras = Math.abs(montoTotalCompras);
          montoTotalVentas = Math.abs(montoTotalVentas);
        }

        const fechaDoctoIndex = data[0].indexOf('Fecha Docto');
        const originalDate = row[fechaDoctoIndex];
        const formattedDate = formatDate(originalDate);

        const combinedRow = {
          correlativo: combinedData.diciembre.length + 1,
          tipoOperacion: tipoOperacion1o2,
          folio: row[data[0].indexOf('Folio')],
          tipoDocumento: tipoDocumento,
          rutEmisor: row[data[0].indexOf(fileIndex === 0 ? 'RUT Proveedor' : 'Rut cliente')],
          fechaOperacion: formattedDate,
          montoNeto: montoNeto,
          iva: montoIva,
          montoOperacionesExentas: montoTotal - (montoNeto + montoIva) === 0 ? 0 : montoTotal - (montoNeto + montoIva),
          montoTotal: montoTotal,
          montoTotalCompras: montoTotalCompras,
          montoTotalVentas: montoTotalVentas,
          montoPercibido: montoTotal,
          glosaOperacion: glosa,
          operacionEntidadRelacionada: '',
          percepcionOperacionDevengada: '',
          operacionPagoPlazo: '',
          fechaExigibilidadPago: formattedDate,
          montoIngreso: fileIndex === 2 ? formatNumber(montoTotalVentas) : '',
          montoEgreso: fileIndex === 3 ? formatNumber(montoTotalCompras) : '',
          saldo: ''
        };
        combinedData.diciembre.push(combinedRow);
      });
    });

    const tester1 = (data) => {
      return data.reduce((acc, item) => {
        if (item.montoTotal > 0) {
          acc[item.folio] = item;  // Sobrescribe el valor si ya existe
        }
        return acc;
      }, {});
    };

    // console.log(combinedData.abril)
    //Esto elimina los datos con folio repetido
    combinedData.enero = Object.values(tester1(combinedData.enero))
    combinedData.febrero = Object.values(tester1(combinedData.febrero))
    combinedData.marzo = Object.values(tester1(combinedData.marzo))
    combinedData.abril = Object.values(tester1(combinedData.abril))
    combinedData.mayo = Object.values(tester1(combinedData.mayo))
    combinedData.junio = Object.values(tester1(combinedData.junio))
    combinedData.julio = Object.values(tester1(combinedData.julio))
    combinedData.agosto = Object.values(tester1(combinedData.agosto))
    combinedData.septiembre = Object.values(tester1(combinedData.septiembre))
    combinedData.octubre = Object.values(tester1(combinedData.octubre))
    combinedData.noviembre = Object.values(tester1(combinedData.noviembre))
    combinedData.diciembre = Object.values(tester1(combinedData.diciembre))
    // console.log(combinedData.abril)

    // const testerData1 = Object.values(tester1(combinedData.enero));
    // const testerData2 = Object.values(tester1(combinedData.febrero));
    // console.log(combinedData.enero)

    // Ordenar combinedData por fechaOperacion cronológicamente
    combinedData.enero.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.febrero.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.marzo.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.abril.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.mayo.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.junio.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.julio.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.agosto.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.septiembre.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.octubre.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.noviembre.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));
    combinedData.diciembre.sort((a, b) => new Date(a.fechaOperacion.split('/').reverse().join('/')) / new Date(b.fechaOperacion.split('/').reverse().join('/')));


    // Arregla el correlativo
    combinedData.enero.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.febrero.forEach((row, index) => {
      row.correlativo = index + 1;
    });
    combinedData.marzo.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.abril.forEach((row, index) => {
      row.correlativo = index + 1;
    });
    combinedData.mayo.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.junio.forEach((row, index) => {
      row.correlativo = index + 1;
    });
    combinedData.julio.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.agosto.forEach((row, index) => {
      row.correlativo = index + 1;
    });
    combinedData.septiembre.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.octubre.forEach((row, index) => {
      row.correlativo = index + 1;
    });
    combinedData.noviembre.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    combinedData.diciembre.forEach((row, index) => {
      row.correlativo = index + 1;
    });

    console.dir(combinedData, { depth: null });

    setCombinedData(combinedData);
    navigate('/combined');
  };


  return (
    <Box>
      <Button variant="contained" color="primary" onClick={handleCombineSheets}>
        Unir Planillas
      </Button>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows1}
          columns={columns1}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows2}
          columns={columns2}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows3}
          columns={columns3}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows4}
          columns={columns4}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows5}
          columns={columns5}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows6}
          columns={columns6}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows7}
          columns={columns7}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows8}
          columns={columns8}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows9}
          columns={columns9}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows10}
          columns={columns10}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows11}
          columns={columns11}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows12}
          columns={columns12}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows13}
          columns={columns13}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows14}
          columns={columns14}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows15}
          columns={columns15}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows16}
          columns={columns16}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows17}
          columns={columns17}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows18}
          columns={columns18}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows19}
          columns={columns19}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows20}
          columns={columns20}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows21}
          columns={columns21}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows22}
          columns={columns22}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%' }}>
        <DataGrid
          rows={rows23}
          columns={columns23}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />
      </Box>
      <Box sx={{ height: 600, width: '100%', marginTop: 2 }}>
        <DataGrid
          rows={rows24}
          columns={columns24}
          initialState={{
            pagination: { paginationModel: { page: 0, pageSize: 10 } },
          }}
          pageSizeOptions={[10, 10]}
          checkboxSelection
        />

      </Box>
      <Box sx={{ display: 'flex', justifyContent: 'center', marginTop: 2 }}>
      </Box>
    </Box>
  );
};

export default Display;