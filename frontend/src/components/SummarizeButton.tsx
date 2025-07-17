import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';

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

      // Mock API call for now - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 2000)); // Simulate API delay
      const mockSummary = `Key points:\n• abc\n• xyz\n• 123`;
      
      setResult(mockSummary);
      
    } catch (error) {
      setResult('Error: Failed to summarize text. Please try again.');
      console.error('Summarization error:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="summarize-panel" style={{ padding: '20px' }}>
      <div className="panel-header">
        <h3>Summarize Text</h3>
        <p>Select text in your document and get a summary by Oliver.</p>
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
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600' }}>Summary:</h4>
          <div className="result-text">{result}</div>
          <div className="result-actions">
            <button 
              onClick={() => setResult('')}
              className="action-button tertiary"
            >
              Reset
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default SummarizeButton;