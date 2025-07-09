import React, { useState, useEffect } from 'react';
import { wordApi } from '../office/wordApi';
import SummarizeButton from './SummarizeButton';
import ComparePanel from './ComparePanel';
import RedraftPanel from './RedraftPanel';
import AssistantChat from './AssistantChat';

type ActivePanel = 'summarize' | 'compare' | 'redraft' | 'chat';

const Sidebar: React.FC = () => {
  const [activePanel, setActivePanel] = useState<ActivePanel>('summarize');
  const [isInitialized, setIsInitialized] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const initializeOffice = async () => {
      try {
        await wordApi.initialize();
        setIsInitialized(true);
        console.log('Office.js initialized successfully');
      } catch (error) {
        console.error('Failed to initialize Office.js:', error);
        // For development, we'll continue even if Office.js fails
        setIsInitialized(true);
      } finally {
        setLoading(false);
      }
    };

    initializeOffice();
  }, []);
  
  if (loading) {
    return (
      <div className="sidebar-container">
        <div className="sidebar-header">
          <h2>Oliver AI Assistant</h2>
          {/* <p>Initializing...</p> */}
        </div>
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <p>Loading Office.js integration...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="sidebar-container">
      <div className="sidebar-header" style={{ backgroundColor: 'white', color: 'black' }}>
        <h2 style={{ color: 'black' }}>Oliver AI Assistant</h2>
        <p style={{ color: 'black' }}>Your AI-powered legal workflow assistant</p>
        {isInitialized}
      </div>

      <div className="sidebar-nav">
        <button
          className={`nav-button ${activePanel === 'summarize' ? 'active' : ''}`}
          onClick={() => setActivePanel('summarize')}
        >
          Summarize
        </button>
        <button
          className={`nav-button ${activePanel === 'compare' ? 'active' : ''}`}
          onClick={() => setActivePanel('compare')}
        >
          Compare
        </button>
        <button
          className={`nav-button ${activePanel === 'redraft' ? 'active' : ''}`}
          onClick={() => setActivePanel('redraft')}
        >
          Redraft
        </button>
        <button
          className={`nav-button ${activePanel === 'chat' ? 'active' : ''}`}
          onClick={() => setActivePanel('chat')}
        >
          Assistant
        </button>
      </div>

      <div className="sidebar-content">
        {activePanel === 'summarize' && <SummarizeButton />}
        {activePanel === 'compare' && <ComparePanel />}
        {activePanel === 'redraft' && <RedraftPanel />}
        {activePanel === 'chat' && <AssistantChat />}
      </div>
    </div>
  );
};

export default Sidebar;