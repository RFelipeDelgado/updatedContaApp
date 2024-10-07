import React, { useState } from 'react';
import { TextField, Button, Grid, Box, Typography } from '@mui/material';

const FormularioMeses = () => {
  const [openForm, setOpenForm] = useState(false);

  const placeholders = ['IVA', 'Imposiciones', 'Arriendo', 'Remuneraciones', 'Boletas', 'Vouchers'];
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

  const [valores, setValores] = useState(
    Array.from({ length: 12 }, () => ({ 1: '', 2: '', 3: '', 4: '', 5: '', 6: '' }))
  );

  const handleChange = (e, mes, index) => {
    const updatedValores = [...valores];
    updatedValores[mes - 1][index] = e.target.value;
    setValores(updatedValores);
  };

  const handleSubmit = () => {
    console.log('Valores ingresados:', valores);
    // Aquí puedes manejar el envío de los datos, guardarlos en un estado global o enviarlos a una API
  };

  return (
    <Box>
      <Button variant="contained" onClick={() => setOpenForm(true)}>
        Ingresar valores
      </Button>

      {openForm && (
        <Box sx={{ marginTop: 3 }}>
          <Typography variant="h6" gutterBottom>
            Ingresar valores numéricos para cada mes
          </Typography>

          <form onSubmit={handleSubmit}>
            {valores.map((mes, mesIndex) => (
              <Box key={mesIndex} sx={{ marginBottom: 2 }}>
                <Typography variant="subtitle1">{meses[mesIndex]}</Typography>
                <Grid container spacing={2}>
                  {Object.keys(mes).map((index) => (
                    <Grid item xs={2} key={index}>
                      <TextField
                        type="number"
                        placeholder={placeholders[index - 1]} // Muestra el placeholder correspondiente
                        value={mes[index]}
                        onChange={(e) => handleChange(e, mesIndex + 1, index)}
                        fullWidth
                      />
                    </Grid>
                  ))}
                </Grid>
              </Box>
            ))}

            <Button onClick={handleSubmit} variant="contained" color="primary">
              Enviar
            </Button>
          </form>
        </Box>
      )}
    </Box>
  );
};

export default FormularioMeses;
