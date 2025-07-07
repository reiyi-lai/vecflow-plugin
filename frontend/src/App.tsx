import React from 'react';
import Sidebar from './components/Sidebar';
import './App.css';

function App() {
  console.log('App component rendering');
  
  return (
    <div className="app">
      <Sidebar />
    </div>
  );
}

export default App;
