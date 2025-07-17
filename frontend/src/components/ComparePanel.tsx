import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';

const ComparePanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [clause1, setClause1] = useState<string>('');
  const [clause2, setClause2] = useState<string>('');
  const [result, setResult] = useState<boolean>(false);
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
      setResult(false);
      
      // Mock API call - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 2500));
      
      setResult(true);
      
    } catch (error) {
      console.error('Comparison error:', error);
      alert('Error: Failed to compare clauses. Please try again.');
    } finally {
      setLoading(false);
    }
  };


  const handleReset = () => {
    setClause1('');
    setClause2('');
    setResult(false);
    setStep('select-first');
  };

  return (
    <div className="compare-panel" style={{ padding: '20px' }}>
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
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600' }}>Comparison Result:</h4>
          <div className="result-text">
            <p>Lorem ipsum dolor sit amet</p>
          </div>
          <div className="result-actions">
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