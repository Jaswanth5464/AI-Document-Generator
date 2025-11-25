import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom';
import { useState, useEffect } from 'react';
import { auth } from './firebase';
import Login from './pages/Login';
import Signup from './pages/Signup';
import Dashboard from './pages/Dashboard';
import ConfigureWord from './pages/ConfigureWord';
import ConfigurePPT from './pages/ConfigurePPT';
import GenerateContent from './pages/GenerateContent';
import RefineContent from './pages/RefineContent';

function App() {
  const [user, setUser] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = auth.onAuthStateChanged((user) => {
      setUser(user);
      setLoading(false);
    });
    return unsubscribe;
  }, []);

  if (loading) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '24px'
      }}>
        ⏳ Loading...
      </div>
    );
  }

  return (
    <BrowserRouter>
      <Routes>
        <Route path="/login" element={user ? <Navigate to="/dashboard" /> : <Login />} />
        <Route path="/signup" element={user ? <Navigate to="/dashboard" /> : <Signup />} />
        <Route path="/dashboard" element={user ? <Dashboard /> : <Navigate to="/login" />} />
        <Route path="/configure-word/:projectId" element={user ? <ConfigureWord /> : <Navigate to="/login" />} />
        <Route path="/configure-ppt/:projectId" element={user ? <ConfigurePPT /> : <Navigate to="/login" />} />
        <Route path="/generate/:projectId" element={user ? <GenerateContent /> : <Navigate to="/login" />} />
        <Route path="/refine/:projectId" element={user ? <RefineContent /> : <Navigate to="/login" />} />
        <Route path="/" element={<Navigate to="/dashboard" />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;