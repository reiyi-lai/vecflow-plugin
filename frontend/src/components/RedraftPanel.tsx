import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';

const RedraftPanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [originalText, setOriginalText] = useState<string>('');
  const [instructions, setInstructions] = useState<string>('');
  const [result, setResult] = useState<any>(null);
  const [showSuggestion, setShowSuggestion] = useState(false);

  const handleSelectText = async () => {
    try {
      // Store the selected range for later replacement
      await new Promise((resolve) => {
        Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.load('text');
          await context.sync();
          
          if (!selection.text || selection.text.trim() === '') {
            alert('Please select some text to redraft.');
            return;
          }
          
          // Store both the text and the range
          setOriginalText(selection.text);
          setResult(null);
          setShowSuggestion(false);
          resolve(true);
        });
      });
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
        redraftedText: `[Placeholder text for replacement]`,
        changes: [
          'This is a mock redraft for testing purposes',
          'The selected text will be replaced with the placeholder',
          'In production, this would be AI-generated content',
          instructions ? `Applied specific instruction: "${instructions}"` : 'General style improvements applied'
        ].filter(Boolean)
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
      // Use Word.run to replace the text more reliably
      await new Promise((resolve, reject) => {
        Word.run(async (context) => {
          try {
            // Find and replace the original text with the redrafted text
            const searchResults = context.document.body.search(originalText, { matchCase: false, matchWholeWord: false });
            searchResults.load('items');
            await context.sync();
            
            if (searchResults.items.length > 0) {
              // Replace the first occurrence (should be our selected text)
              searchResults.items[0].insertText(result.redraftedText, Word.InsertLocation.replace);
              await context.sync();
              resolve(true);
            } else {
              // Fallback: just insert at current selection
              const selection = context.document.getSelection();
              selection.insertText(result.redraftedText, Word.InsertLocation.replace);
              await context.sync();
              resolve(true);
            }
          } catch (error) {
            reject(error);
          }
        });
      });
      
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
      <div className="panel-header" style={{ marginBottom: '20px' }}>
        <h3>Redraft Text</h3>
        <p>Select text to redraft. Add optional instructions for specific requirements.</p>
      </div>
      
      <div className="redraft-steps">
        <div className="step" style={{ 
          marginBottom: '20px',
          padding: '16px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <div className="step-header" style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
            <span className="step-number" style={{ 
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '24px',
              height: '24px',
              backgroundColor: '#2b5ce6',
              color: 'white',
              borderRadius: '50%',
              fontSize: '12px',
              fontWeight: '600',
              marginRight: '8px'
            }}>1</span>
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>Select Text</span>
          </div>
          <button 
            onClick={handleSelectText}
            disabled={loading}
            className={`action-button ${originalText ? 'success' : 'primary'}`}
          >
            {originalText ? 'Text Selected ✓' : 'Select Text to Redraft'}
          </button>
          {originalText && (
            <div className="text-preview" style={{ 
              marginTop: '12px',
              padding: '12px',
              backgroundColor: 'white',
              borderRadius: '4px',
              border: '1px solid #dee2e6',
              fontSize: '13px'
            }}>
              <strong>Selected Text:</strong>
              <div className="preview-text" style={{ 
                marginTop: '8px',
                padding: '8px',
                backgroundColor: '#f8f9fa',
                borderRadius: '4px',
                border: '1px solid #e9ecef',
                fontSize: '13px',
                whiteSpace: 'pre-wrap',
                maxHeight: '100px',
                overflow: 'auto'
              }}>{originalText}</div>
            </div>
          )}
        </div>

        <div className="step" style={{ 
          marginBottom: '20px',
          padding: '16px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <div className="step-header" style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
            <span className="step-number" style={{ 
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '24px',
              height: '24px',
              backgroundColor: '#2b5ce6',
              color: 'white',
              borderRadius: '50%',
              fontSize: '12px',
              fontWeight: '600',
              marginRight: '8px'
            }}>2</span>
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>Instructions (Optional)</span>
          </div>
          <textarea
            value={instructions}
            onChange={(e) => setInstructions(e.target.value)}
            placeholder="Enter specific instructions (e.g., 'make it more formal', 'simplify language', 'add more detail')"
            className="instructions-input"
            rows={3}
            disabled={loading}
            style={{ 
              width: '100%',
              padding: '12px',
              border: '1px solid #ced4da',
              borderRadius: '4px',
              fontSize: '14px',
              fontFamily: 'inherit',
              resize: 'vertical',
              minHeight: '60px',
              backgroundColor: 'white',
              color: 'black',
              boxSizing: 'border-box'
            }}
          />
        </div>

        <div className="step" style={{ 
          marginBottom: '20px',
          padding: '16px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <div className="step-header" style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
            <span className="step-number" style={{ 
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '24px',
              height: '24px',
              backgroundColor: '#2b5ce6',
              color: 'white',
              borderRadius: '50%',
              fontSize: '12px',
              fontWeight: '600',
              marginRight: '8px'
            }}>3</span>
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>Redraft</span>
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
        <div className="suggestion-container" style={{ 
          marginTop: '20px',
          padding: '16px',
          backgroundColor: '#fff3cd',
          borderRadius: '8px',
          border: '1px solid #ffeaa7'
        }}>
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600', color: 'black' }}>AI Suggestion:</h4>
          
          <div className="suggestion-comparison" style={{ marginBottom: '16px' }}>
            <div className="original-text" style={{ marginBottom: '12px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600', color: 'black' }}>Original:</h5>
              <div className="text-block original" style={{ 
                marginBottom: '12px',
                padding: '12px',
                borderRadius: '4px',
                fontSize: '13px',
                lineHeight: '1.4',
                whiteSpace: 'pre-wrap',
                backgroundColor: '#f8d7da',
                border: '1px solid #f5c6cb',
                color: 'black'
              }}>{originalText}</div>
            </div>
            
            <div className="redrafted-text">
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600', color: 'black' }}>Redrafted:</h5>
              <div className="text-block redrafted" style={{ 
                marginBottom: '12px',
                padding: '12px',
                borderRadius: '4px',
                fontSize: '13px',
                lineHeight: '1.4',
                whiteSpace: 'pre-wrap',
                backgroundColor: '#d4edda',
                border: '1px solid #c3e6cb',
                color: 'black'
              }}>{result.redraftedText}</div>
            </div>
          </div>

          {/* {result.changes && result.changes.length > 0 && (
            <div className="changes-section" style={{ marginBottom: '16px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Key Changes:</h5>
              <ul style={{ margin: '0', paddingLeft: '20px' }}>
                {result.changes.map((change: string, index: number) => (
                  <li key={index} style={{ marginBottom: '4px', fontSize: '13px', lineHeight: '1.4' }}>{change}</li>
                ))}
              </ul>
            </div>
          )} */}

          <div className="suggestion-actions" style={{ display: 'flex', gap: '8px' }}>
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