import { useState } from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
// import Home from './Pages/Home/Home';
import Display from './Pages/Display/Display';
import Combined from './Pages/Combined/Combined';
import FormularioMeses from './Pages/tester/tester';
import HomeTester from './Pages/Home/HomeTester';

const App = () => {
  const [filesData, setFilesData] = useState([]);
  const [regularInputs, setRegularInputs] = useState({});
  const [combinedData, setCombinedData] = useState([]);

  return (
    <Router>
      <Routes>
        <Route path="/" element={<HomeTester setFilesData={setFilesData} setRegularInputs={setRegularInputs} />} />
        <Route path="/display" element={<Display filesData={filesData} regularInputs={regularInputs} setCombinedData={setCombinedData}  />} />
        <Route path="/combined" element={<Combined combinedData={combinedData} />} />
        <Route path="/tester" element={<FormularioMeses FormularioMeses ={FormularioMeses } />} />
        {/* <Route path="/hometester" element={<Home setFilesData={setFilesData}  />} /> */}
      </Routes>
    </Router>
  );
};

export default App;