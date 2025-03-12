import * as XLSX from 'xlsx';
import { useNavigate } from 'react-router-dom';
import { Button, Box, TextField, Typography, InputLabel, FormControl, Accordion, AccordionDetails, AccordionSummary, Grid2, Divider, Input, InputAdornment } from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import EditCalendarIcon from '@mui/icons-material/EditCalendar';

import { useState } from 'react';

// eslint-disable-next-line react/prop-types
const HomeTester = ({ setFilesData }) => {
    const [file1, setFile1] = useState(null);
    const [file2, setFile2] = useState(null);
    const [file3, setFile3] = useState(null);
    const [file4, setFile4] = useState(null);
    const [file5, setFile5] = useState(null);
    const [file6, setFile6] = useState(null);
    const [file7, setFile7] = useState(null);
    const [file8, setFile8] = useState(null);
    const [file9, setFile9] = useState(null);
    const [file10, setFile10] = useState(null);
    const [file11, setFile11] = useState(null);
    const [file12, setFile12] = useState(null);
    const [file13, setFile13] = useState(null);
    const [file14, setFile14] = useState(null);
    const [file15, setFile15] = useState(null);
    const [file16, setFile16] = useState(null);
    const [file17, setFile17] = useState(null);
    const [file18, setFile18] = useState(null);
    const [file19, setFile19] = useState(null);
    const [file20, setFile20] = useState(null);
    const [file21, setFile21] = useState(null);
    const [file22, setFile22] = useState(null);
    const [file23, setFile23] = useState(null);
    const [file24, setFile24] = useState(null);
    const navigate = useNavigate();

    const handleFileChange = (e, setFile) => {
        setFile(e.target.files[0]);
    };

    const handleProcessFiles = async () => {
        if (!file1 || !file2 || !file3 || !file4 || !file5 || !file6 || !file7 || !file8 || !file9 || !file10 || !file11 || !file12) {
            alert('Tienes que agregar todos los archivos, stupid');
            return;
        }

        const readFile = (file) => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    resolve(jsonData);
                };
                reader.onerror = (err) => reject(err);
                reader.readAsArrayBuffer(file);
            });
        };

        try {
            const [data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12, data13, data14, data15, data16, data17, data18, data19, data20, data21, data22, data23, data24] = await Promise.all([
                readFile(file1),
                readFile(file2),
                readFile(file3),
                readFile(file4),
                readFile(file5),
                readFile(file6),
                readFile(file7),
                readFile(file8),
                readFile(file9),
                readFile(file10),
                readFile(file11),
                readFile(file12),
                readFile(file13),
                readFile(file14),
                readFile(file15),
                readFile(file16),
                readFile(file17),
                readFile(file18),
                readFile(file19),
                readFile(file20),
                readFile(file21),
                readFile(file22),
                readFile(file23),
                readFile(file24)]);
            setFilesData([data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12, data13, data14, data15, data16, data17, data18, data19, data20, data21, data22, data23, data24]);

            navigate('/display');
        } catch (error) {
            console.error('Error reading files:', error);
        }
    };

    return (
        <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: "center", gap: 2, padding: 10 }}>
            <Typography variant="h5">Subir Archivos Excel</Typography>
            <Accordion >
                <AccordionSummary
                    expandIcon={<ExpandMoreIcon />}
                    aria-controls="panel1-content"
                    id="panel1-header"
                >
                    <EditCalendarIcon sx={{ color: 'orange', mr: 1, my: 0.1 }} />
                    <Typography component="span">Enero</Typography>
                </AccordionSummary>

                <AccordionDetails>

                    <Grid2 container spacing={0}>

                        <Grid2 size={12} sx={{ m: '0' }}>
                            <Divider textAlign="center">
                                <Typography>Seleccionar archivos de Enero de:</Typography>
                            </Divider>
                        </Grid2>


                        <Grid2 size={6} textAlign="center">
                            <InputLabel >Compra</InputLabel>
                            <TextField
                                id="file-upload"
                                type="file"
                                inputProps={{ accept: '.xlsx, .xls, .csv' }}
                                onChange={(e) => handleFileChange(e, setFile1)}
                                sx={{ gap: 2, border: 1, margin: 1, mt: 1 }}
                                variant="outlined"
                                margin="normal"

                            />
                        </Grid2>
                        <Grid2 size={6} textAlign="center">
                            <InputLabel>Venta</InputLabel>
                            <TextField
                                id="file-upload"
                                type="file"
                                inputProps={{ accept: '.xlsx, .xls, .csv' }}
                                onChange={(e) => handleFileChange(e, setFile2)}
                                sx={{ gap: 0, border: 1, margin: 0.1, mt: 1 }}
                                variant="outlined"
                                margin="normal"
                            />
                        </Grid2>

                        <Grid2 size={12} sx={{ m: '1rem' }}>
                            <Divider textAlign="center">
                                <Typography>Otros Egresos:</Typography>
                            </Divider>
                        </Grid2>

                        {/* Remuneraciones */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Remuneraciones</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                        thousandSeparator
                                        valueIsNumericString
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>

                        {/* Arriendo */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Arriendo</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>
                        {/* Impuestos */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Impuestos</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>
                        {/* Comisiones */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Comisiones</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>
                        {/* Prestamos */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Prestamos</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>
                        {/* Imposiciones */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Imposiciones</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>
                        {/* Prestamos bancarios */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Prestamo bancario</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>

                        <Grid2 size={12} sx={{ m: '1rem' }}>
                            <Divider textAlign="center">
                                <Typography>Ingresos:</Typography>
                            </Divider>
                        </Grid2>

                        {/* Ingreso 1 */}
                        <Grid2 size={12}>
                            <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                <Typography sx={{width: 200}}>Otros ingresos</Typography>
                                <TextField
                                    required
                                    type='date'
                                    id={"Enero"}
                                    label={"Fecha"}
                                    // onChange={(e) => handleInputChange(month, monthInputIndex, e.target.value)}
                                    variant="outlined"
                                    margin="normal"
                                    InputLabelProps={{
                                        shrink: true, // Fuerza a que el label esté siempre contraído
                                      }}
                                />
                                <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                    <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                    <Input
                                        id="standard-adornment-amount"
                                        startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                    />
                                </FormControl>
                            </Box>
                        </Grid2>

                    </Grid2>

                </AccordionDetails>
            </Accordion>
            <FormControl fullWidth>


                {/* <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de ENERO</InputLabel> */}
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile1)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de ENERO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile2)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de FEBRERO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile3)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de FEBRERO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile4)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de MARZO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile5)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de MARZO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile6)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de ABRIL</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile7)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de ABRIL</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile8)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de MAYO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile9)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de MAYO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile10)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de JUNIO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile11)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de JUNIO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile12)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de JULIO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile13)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de JULIO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile14)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de AGOSTO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile15)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de AGOSTO</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile16)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de SEPTIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile17)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de SEPTIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile18)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de OCTUBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile19)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de OCTUBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile20)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de NOVIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile21)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de NOVIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile22)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de DICIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile23)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <FormControl fullWidth>
                <InputLabel htmlFor="file-upload" >Sube el archivo de VENTAS de DICIEMBRE</InputLabel>
                <TextField
                    id="file-upload"
                    type="file"
                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                    onChange={(e) => handleFileChange(e, setFile24)}
                    sx={{ gap: 2, border: 1, margin: 1, mt: 5 }}
                    variant="outlined"
                    margin="normal"
                    fullWidth
                />
            </FormControl>
            <Button variant="contained" color="primary" onClick={handleProcessFiles}>Procesar archivos</Button>
        </Box>
    );
};

export default HomeTester;