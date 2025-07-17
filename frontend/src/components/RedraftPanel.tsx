import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';

const RedraftPanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [originalText, setOriginalText] = useState<string>('');
  const [instructions, setInstructions] = useState<string>('');
  const [result, setResult] = useState<{ redraftedText: string } | null>(null);
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
      
      // Mock API call - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 3000));
      
      const mockResult = {
        redraftedText: `<Placeholder text for replacement>`
      };
      
      setResult(mockResult);
      setShowSuggestion(true);
      
    } catch (error) {
      console.error('Redraft error:', error);
      alert('Error: Failed to redraft text. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleAcceptSuggestion = async () => {
    if (!result || !originalText) return;
    
    try {
      await wordApi.replaceSelectedText(result.redraftedText);
      
      setOriginalText('');
      setResult(null);
      setShowSuggestion(false);
      setInstructions('');
      alert('Text replaced successfully!');
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
    <div className="redraft-panel" style={{ padding: '20px' }}>
      <div className="panel-header">
        <h3>Redraft Text</h3>
        <p>Select text to redraft. Add optional instructions for specific requirements.</p>
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
            placeholder="Add specific instructions, e.g. 'make it more formal', 'simplify language', 'add more detail'"
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
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600', color: 'black' }}>AI Suggestion:</h4>
          
          <div className="suggestion-comparison">
            <div className="original-text" style={{ marginBottom: '12px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600', color: 'black' }}>Original:</h5>
              <div className="text-block original">{originalText}</div>
            </div>
            
            <div className="redrafted-text">
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600', color: 'black' }}>Redrafted:</h5>
              <div className="text-block redrafted">{result.redraftedText}</div>
            </div>
          </div>

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