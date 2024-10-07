import { useState } from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Home from './Pages/Home/Home';
import Display from './Pages/Display/Display';
import Combined from './Pages/Combined/Combined';
import FormularioMeses from './Pages/tester/tester';

const App = () => {
  const [filesData, setFilesData] = useState([]);
  const [combinedData, setCombinedData] = useState([]);

  return (
    <Router>
      <Routes>
        <Route path="/" element={<Home setFilesData={setFilesData} />} />
        <Route path="/display" element={<Display filesData={filesData} setCombinedData={setCombinedData} />} />
        <Route path="/combined" element={<Combined combinedData={combinedData} />} />
        <Route path="/tester" element={<FormularioMeses FormularioMeses ={FormularioMeses } />} />
      </Routes>
    </Router>
  );
};

export default App;