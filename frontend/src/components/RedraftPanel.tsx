import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';
import { apiService } from '../services/apiService';

const RedraftPanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [originalText, setOriginalText] = useState<string>('');
  const [instructions, setInstructions] = useState<string>('');
  const [result, setResult] = useState<any>(null);
  const [showSuggestion, setShowSuggestion] = useState(false);

  const handleSelectText = async () => {
    try {
      const selectedText = await wordApi.getSelectedText();
      
      if (!selectedText || selectedText.trim() === '') {
        alert('Please select some text to redraft.');
        return;
      }

      setOriginalText(selectedText);
      setResult(null);
      setShowSuggestion(false);
    } catch (error) {
      console.error('Failed to get selected text:', error);
      alert('Failed to get selected text. Please try again.');
    }
  };

  const handleRedraft = async () => {
    if (!originalText) return;
    
    try {
      setLoading(true);
      setResult(null);
      
      const response = await apiService.redraft({ 
        text: originalText, 
        instructions: instructions || undefined 
      });
      setResult(response);
      setShowSuggestion(true);
      
    } catch (error) {
      console.error('Redraft error:', error);
      alert('Error: Failed to redraft text. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleAcceptSuggestion = async () => {
    if (!result) return;
    
    try {
      await wordApi.replaceSelectedText(result.redraftedText);
      setOriginalText('');
      setResult(null);
      setShowSuggestion(false);
      setInstructions('');
    } catch (error) {
      console.error('Failed to replace text:', error);
      alert('Failed to replace text. Please try again.');
    }
  };

  const handleRejectSuggestion = () => {
    setShowSuggestion(false);
    setResult(null);
  };

  const handleReset = () => {
    setOriginalText('');
    setResult(null);
    setShowSuggestion(false);
    setInstructions('');
  };

  return (
    <div className="redraft-panel">
      <div className="panel-header">
        <h3>Redraft Text</h3>
        <p>Select text to redraft with AI assistance. Add optional instructions for specific style or requirements.</p>
      </div>
      
      <div className="redraft-steps">
        <div className="step">
          <div className="step-header">
            <span className="step-number">1</span>
            <span className="step-title">Select Text</span>
          </div>
          <button 
            onClick={handleSelectText}
            disabled={loading}
            className={`action-button ${originalText ? 'success' : 'primary'}`}
          >
            {originalText ? 'Text Selected ✓' : 'Select Text to Redraft'}
          </button>
          {originalText && (
            <div className="text-preview">
              <strong>Selected Text:</strong>
              <div className="preview-text">{originalText}</div>
            </div>
          )}
        </div>

        <div className="step">
          <div className="step-header">
            <span className="step-number">2</span>
            <span className="step-title">Instructions (Optional)</span>
          </div>
          <textarea
            value={instructions}
            onChange={(e) => setInstructions(e.target.value)}
            placeholder="Enter specific instructions (e.g., 'make it more formal', 'simplify language', 'add more detail')"
            className="instructions-input"
            rows={3}
            disabled={loading}
          />
        </div>

        <div className="step">
          <div className="step-header">
            <span className="step-number">3</span>
            <span className="step-title">Redraft</span>
          </div>
          <button 
            onClick={handleRedraft}
            disabled={loading || !originalText}
            className="action-button primary"
          >
            {loading ? 'Redrafting...' : 'Generate Redraft'}
          </button>
        </div>
      </div>

      {showSuggestion && result && (
        <div className="suggestion-container">
          <h4>AI Suggestion:</h4>
          
          <div className="suggestion-comparison">
            <div className="original-text">
              <h5>Original:</h5>
              <div className="text-block original">{originalText}</div>
            </div>
            
            <div className="redrafted-text">
              <h5>Redrafted:</h5>
              <div className="text-block redrafted">{result.redraftedText}</div>
            </div>
          </div>

          {result.changes && result.changes.length > 0 && (
            <div className="changes-section">
              <h5>Key Changes:</h5>
              <ul>
                {result.changes.map((change: string, index: number) => (
                  <li key={index}>{change}</li>
                ))}
              </ul>
            </div>
          )}

          <div className="suggestion-actions">
            <button 
              onClick={handleAcceptSuggestion}
              className="action-button primary"
            >
              Accept & Replace
            </button>
            <button 
              onClick={handleRejectSuggestion}
              className="action-button secondary"
            >
              Reject
            </button>
            <button 
              onClick={handleReset}
              className="action-button tertiary"
            >
              Start Over
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

export default RedraftPanel;