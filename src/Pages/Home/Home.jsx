import * as XLSX from 'xlsx';
import { useNavigate } from 'react-router-dom';
import { Button, Box, TextField, Typography, InputLabel, FormControl } from '@mui/material';
import { useState } from 'react';

// eslint-disable-next-line react/prop-types
const Home = ({ setFilesData }) => {
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
    if (!file1 || !file2 || !file3 || !file4 || !file5 || !file6 || !file7 || !file8 || !file9 || !file10 || !file11 || !file12 ) {
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
      setFilesData([data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12, data13, data14, data15, data16, data17, data18, data19, data20, data21, data22, data23, data24 ]);

      navigate('/display');
    } catch (error) {
      console.error('Error reading files:', error);
    }
  };

  return (
    <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2, padding: 10 }}>
      <Typography variant="h4">Subir Archivos Excel</Typography>
      <Typography variant="h3">Enero</Typography>
      <FormControl fullWidth>

        <InputLabel htmlFor="file-upload" >Sube el archivo de COMPRAS de ENERO</InputLabel>
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

export default Home;