import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';
import { apiService } from '../services/apiService';

const SummarizeButton: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<string>('');

  const handleSummarize = async () => {
    try {
      setLoading(true);
      setResult('');
      
      const selectedText = await wordApi.getSelectedText();
      
      if (!selectedText || selectedText.trim() === '') {
        setResult('Please select some text to summarize.');
        return;
      }

      const response = await apiService.summarize({ text: selectedText });
      setResult(response.summary);
      
    } catch (error) {
      setResult('Error: Failed to summarize text. Please try again.');
      console.error('Summarization error:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleInsertSummary = async () => {
    if (!result) return;
    
    try {
      await wordApi.insertTextAtSelection(`\n\nSummary: ${result}`);
      setResult('');
    } catch (error) {
      console.error('Failed to insert summary:', error);
    }
  };

  return (
    <div className="summarize-panel">
      <div className="panel-header">
        <h3>Summarize Text</h3>
        <p>Select text in your document and click summarize to get an AI-generated summary.</p>
      </div>
      
      <button 
        onClick={handleSummarize}
        disabled={loading}
        className="action-button primary"
      >
        {loading ? 'Summarizing...' : 'Summarize Selected Text'}
      </button>
      
      {result && (
        <div className="result-container">
          <h4>Summary:</h4>
          <div className="result-text">{result}</div>
          <div className="result-actions">
            <button 
              onClick={handleInsertSummary}
              className="action-button secondary"
            >
              Insert into Document
            </button>
            <button 
              onClick={() => setResult('')}
              className="action-button tertiary"
            >
              Clear
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default SummarizeButton;