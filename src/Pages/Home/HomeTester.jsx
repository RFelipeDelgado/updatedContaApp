import * as XLSX from 'xlsx';
import { useNavigate } from 'react-router-dom';
import { Button, Box, TextField, Typography, InputLabel, FormControl, Accordion, AccordionDetails, AccordionSummary, Grid2, Divider, Input, InputAdornment, Autocomplete, Select, MenuItem } from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import EditCalendarIcon from '@mui/icons-material/EditCalendar';

import { useEffect, useState } from 'react';
import { NumericFormat } from 'react-number-format';

// eslint-disable-next-line react/prop-types
const HomeTester = ({ setFilesData, setRegularInputs }) => {
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

    const mesesAño = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

    const [inputs, setInputs] = useState(
        mesesAño.reduce((acc, mes) => ({
            ...acc,
            [mes]: { regularInputs: [] },

        }), {})
    )
    useEffect(() => {
        console.log(inputs);

    }, [inputs])

    const handleFileChange = (e, setFile) => {
        setFile(e.target.files[0]);
    };

    const handleProcessFiles = async () => {
        if (!file1 || !file2 || !file3 || !file4 || !file5 || !file6 || !file7 || !file8 || !file9 || !file10 || !file11 || !file12) {
            console.log('Tienes que agregar todos los archivos, stupid');
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
            setRegularInputs(inputs)
            navigate('/display');
        } catch (error) {
            console.error('Error reading files:', error);
        }
    };

    const setterFiles = [
        setFile1, setFile2, setFile3, setFile4, setFile5, setFile6,
        setFile7, setFile8, setFile9, setFile10, setFile11, setFile12,
        setFile13, setFile14, setFile15, setFile16, setFile17, setFile18,
        setFile19, setFile20, setFile21, setFile22, setFile23, setFile24
    ];

    const handleAddInput = (mes) => {
        setInputs((prevInputs) => ({
            ...prevInputs,
            [mes]: {
                regularInputs: [
                    ...prevInputs[mes].regularInputs,
                    { nombre: "", cantidad: 0, fecha: 0, selector: "gasto" }
                ],
            },
        }));
        console.log(inputs);

    };

    // Update input value directly using index
    const handleInputChange = (mes, index, value, field) => {
        let processedValue = value;
        
        // Si el campo es "cantidad", elimina símbolos y comas
        if (field === "cantidad") {
            processedValue = parseInt(value.replace(/\D/g, "")); // Elimina todo lo que no sea dígito
        }
    
        console.log(parseInt(processedValue)); // Verifica el valor limpio en consola
    
        setInputs((prevInputs) => {
            const updated = [...prevInputs[mes].regularInputs];
            updated[index] = { ...updated[index], [field]: processedValue };
            return {
                ...prevInputs,
                [mes]: { regularInputs: updated },
            };
        });
    };

    // const selectorOptions2 = ["Si", "No"]

    return (
        <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: "center", gap: 2, padding: 10 }}>
            <Typography variant="h5">Subir Archivos Excel</Typography>

            {mesesAño.map((mes, index) => (
                <Accordion key={`mes_${index}`}>
                    <AccordionSummary
                        expandIcon={<ExpandMoreIcon />}
                        aria-controls="panel1-content"
                        id="panel1-header"
                    >
                        <EditCalendarIcon sx={{ color: 'orange', mr: 1, my: 0.1 }} />
                        <Typography component="span">{mes}</Typography>
                    </AccordionSummary>

                    <AccordionDetails>

                        <Grid2 container spacing={0}>

                            <Grid2 size={12} sx={{ m: '0' }}>
                                <Divider textAlign="center">
                                    <Typography>Seleccionar archivos de {mes} de:</Typography>
                                </Divider>
                            </Grid2>


                            <Grid2 size={6} textAlign="center">
                                <InputLabel >Compra</InputLabel>
                                <TextField
                                    id="file-upload"
                                    type="file"
                                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                                    onChange={(e) => handleFileChange(e, setterFiles[2 * index])}
                                    sx={{ gap: 2, border: 1, margin: 1, mt: 1 }}
                                    variant="outlined"
                                    margin="normal"

                                />
                                <p>{`index: ${index}, mes: ${mes}, valor compra: ${2 * index + 1}, valor venta: ${2 * index + 2}`}</p>
                            </Grid2>
                            <Grid2 size={6} textAlign="center">
                                <InputLabel>Venta</InputLabel>
                                <TextField
                                    id="file-upload"
                                    type="file"
                                    inputProps={{ accept: '.xlsx, .xls, .csv' }}
                                    onChange={(e) => handleFileChange(e, setterFiles[2 * index + 1])}
                                    sx={{ gap: 0, border: 1, margin: 0.1, mt: 1 }}
                                    variant="outlined"
                                    margin="normal"
                                />
                            </Grid2>

                            <Grid2 size={12} sx={{ m: '1rem' }}>
                                <Divider textAlign="center">
                                    <Typography>Otros:</Typography>
                                </Divider>
                            </Grid2>

                            {inputs[mes].regularInputs.map((elem, idx) => (
                                <Grid2 size={12} key={idx}>
                                    <Box sx={{
                                        display: 'flex',
                                        alignItems: 'center',  // Cambiado de 'safe-center' a 'center' para mejor soporte
                                        justifyContent: "space-between",
                                        gap: 2,  // Añade un espacio consistente entre elementos
                                    }}>
                                        <TextField
                                            required
                                            type='normal'
                                            id={"Enero"}
                                            label={"Nombre"}
                                            variant="outlined"
                                            sx={{
                                                flex: 1,
                                                '& .MuiOutlinedInput-root': { height: 56 } // Altura fija
                                            }}
                                            onChange={(e) => handleInputChange(mes, idx, e.target.value, "nombre")}
                                        />
                                        <Select
                                            sx={{
                                                width: 200,
                                                height: 56,  // Misma altura que TextField
                                                marginTop: 1, // Compensa el margen 'normal' de TextField
                                                marginBottom: 1,
                                            }}
                                            onChange={(e) => handleInputChange(mes, idx, e.target.value, "selector")}
                                            value={inputs[mes]?.regularInputs[idx]?.selector || ""}
                                        >
                                            <MenuItem value={"gasto"}>Gasto</MenuItem>
                                            <MenuItem value={"ingreso"}>Ingreso</MenuItem>
                                        </Select>
                                        <TextField
                                            required
                                            type='date'
                                            id={"Enero"}
                                            label={"Fecha"}
                                            variant="outlined"
                                            sx={{
                                                flex: 1,
                                                '& .MuiOutlinedInput-root': { height: 56 } // Altura fija
                                            }}
                                            value={inputs[mes]?.regularInputs[idx]?.fecha || ""}
                                            onChange={(e) => handleInputChange(mes, idx, e.target.value, "fecha")}
                                            InputLabelProps={{ shrink: true }}
                                        />
                                        <FormControl sx={{
                                            m: 1,
                                            width: 220,
                                            '& .MuiInput-root': { height: 56 } // Altura consistente
                                        }} variant="standard">
                                            <NumericFormat
                                                value={inputs[mes]?.regularInputs[idx]?.cantidad || ""}
                                                onChange={(e) => handleInputChange(mes, idx, e.target.value, "cantidad")}
                                                customInput={TextField}
                                                thousandSeparator
                                                valueIsNumericString
                                                prefix="$"
                                                variant="standard"
                                                label="react-number-format"
                                            />
                                        </FormControl>
                                    </Box>
                                </Grid2>
                            ))}

                            <Button onClick={() => handleAddInput(mes)}>Agregar otro</Button>

                            {/* <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Remuneraciones</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 0, e.target.value, "Remuneraciones")}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2> */}

                            {/* <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Arriendo</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 1, e.target.value)}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2>

                            <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Impuestos</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 2, e.target.value)}

                                            startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                        />
                                    </FormControl>
                                </Box>
                            </Grid2>

                            <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Comisiones</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 3, e.target.value)}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2>

                            <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Imposiciones</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 4, e.target.value)}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2>

                            <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Prestamo bancario</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
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
                                            onChange={(e) => handleInputChange(mes, 5, e.target.value)}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2>

                            <Grid2 size={12} sx={{ m: '1rem' }}>
                                <Divider textAlign="center">
                                    <Typography>Ingresos:</Typography>
                                </Divider>
                            </Grid2>

                            <Grid2 size={12}>
                                <Box sx={{ display: 'flex', alignItems: 'flex-end', justifyContent: "space-between" }}>
                                    <Typography sx={{ width: 200 }}>Otros ingresos</Typography>
                                    <TextField
                                        required
                                        type='date'
                                        id={"Enero"}
                                        label={"Fecha"}
                                        variant="outlined"
                                        margin="normal"
                                        InputLabelProps={{
                                            shrink: true, // Fuerza a que el label esté siempre contraído
                                        }}
                                    />
                                    <TextField
                                        required
                                        type='normal'
                                        id={"Enero"}
                                        label={"Glosa"}
                                        onChange={(e) => handleInputChange(mes, 6, e.target.value)}
                                        variant="outlined"
                                        margin="normal"
                                    />
                                    <FormControl sx={{ m: 1, width: 220 }} variant="standard">
                                        <InputLabel htmlFor="standard-adornment-amount">Cantidad</InputLabel>
                                        <Input
                                            id="standard-adornment-amount"
                                            startAdornment={<InputAdornment position="start">$</InputAdornment>}
                                            onChange={(e) => handleInputChange(mes, 7, e.target.value)}

                                        />
                                    </FormControl>
                                </Box>
                            </Grid2> */}

                        </Grid2>

                    </AccordionDetails>
                </Accordion>
            ))}


            <Button variant="contained" color="primary" onClick={handleProcessFiles}>Procesar archivos</Button>
        </Box>
    );
};

export default HomeTester;