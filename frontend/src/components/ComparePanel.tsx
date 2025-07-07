import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';
import { apiService } from '../services/apiService';

const ComparePanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [clause1, setClause1] = useState<string>('');
  const [clause2, setClause2] = useState<string>('');
  const [result, setResult] = useState<any>(null);
  const [step, setStep] = useState<'select-first' | 'select-second' | 'ready'>('select-first');

  const handleSelectClause = async (clauseNumber: 1 | 2) => {
    try {
      const selectedText = await wordApi.getSelectedText();
      
      if (!selectedText || selectedText.trim() === '') {
        alert('Please select some text first.');
        return;
      }

      if (clauseNumber === 1) {
        setClause1(selectedText);
        setStep('select-second');
      } else {
        setClause2(selectedText);
        setStep('ready');
      }
    } catch (error) {
      console.error('Failed to get selected text:', error);
      alert('Failed to get selected text. Please try again.');
    }
  };

  const handleCompare = async () => {
    if (!clause1 || !clause2) return;
    
    try {
      setLoading(true);
      setResult(null);
      
      const response = await apiService.compare({ clause1, clause2 });
      setResult(response);
      
    } catch (error) {
      console.error('Comparison error:', error);
      alert('Error: Failed to compare clauses. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleInsertComparison = async () => {
    if (!result) return;
    
    try {
      const comparisonText = `\n\nClause Comparison:\n${result.comparison}\n\nKey Differences:\n${result.differences.join('\n')}\n\nRecommendations:\n${result.recommendations.join('\n')}`;
      await wordApi.insertTextAtSelection(comparisonText);
      handleReset();
    } catch (error) {
      console.error('Failed to insert comparison:', error);
    }
  };

  const handleReset = () => {
    setClause1('');
    setClause2('');
    setResult(null);
    setStep('select-first');
  };

  return (
    <div className="compare-panel">
      <div className="panel-header">
        <h3>Compare Clauses</h3>
        <p>Select two clauses to compare and analyze their differences.</p>
      </div>
      
      <div className="compare-steps">
        <div className="step">
          <div className="step-header">
            <span className="step-number">1</span>
            <span className="step-title">First Clause</span>
          </div>
          <button 
            onClick={() => handleSelectClause(1)}
            disabled={loading}
            className={`action-button ${step === 'select-first' ? 'primary' : clause1 ? 'success' : 'secondary'}`}
          >
            {clause1 ? 'First Clause Selected ✓' : 'Select First Clause'}
          </button>
          {clause1 && (
            <div className="clause-preview">
              <strong>Preview:</strong> {clause1.substring(0, 100)}...
            </div>
          )}
        </div>

        <div className="step">
          <div className="step-header">
            <span className="step-number">2</span>
            <span className="step-title">Second Clause</span>
          </div>
          <button 
            onClick={() => handleSelectClause(2)}
            disabled={loading || !clause1}
            className={`action-button ${step === 'select-second' ? 'primary' : clause2 ? 'success' : 'secondary'}`}
          >
            {clause2 ? 'Second Clause Selected ✓' : 'Select Second Clause'}
          </button>
          {clause2 && (
            <div className="clause-preview">
              <strong>Preview:</strong> {clause2.substring(0, 100)}...
            </div>
          )}
        </div>

        <div className="step">
          <div className="step-header">
            <span className="step-number">3</span>
            <span className="step-title">Compare</span>
          </div>
          <button 
            onClick={handleCompare}
            disabled={loading || !clause1 || !clause2}
            className="action-button primary"
          >
            {loading ? 'Comparing...' : 'Compare Clauses'}
          </button>
        </div>
      </div>

      {result && (
        <div className="result-container">
          <h4>Comparison Result:</h4>
          <div className="result-text">
            <div className="comparison-section">
              <h5>Analysis:</h5>
              <p>{result.comparison}</p>
            </div>
            <div className="comparison-section">
              <h5>Key Differences:</h5>
              <ul>
                {result.differences.map((diff: string, index: number) => (
                  <li key={index}>{diff}</li>
                ))}
              </ul>
            </div>
            <div className="comparison-section">
              <h5>Recommendations:</h5>
              <ul>
                {result.recommendations.map((rec: string, index: number) => (
                  <li key={index}>{rec}</li>
                ))}
              </ul>
            </div>
          </div>
          <div className="result-actions">
            <button 
              onClick={handleInsertComparison}
              className="action-button secondary"
            >
              Insert into Document
            </button>
            <button 
              onClick={handleReset}
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

export default ComparePanel;