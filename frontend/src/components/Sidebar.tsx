import React, { useState, useEffect } from 'react';
import { wordApi } from '../office/wordApi';

type ActivePanel = 'summarize' | 'compare' | 'redraft' | 'chat';

const SummarizePanel: React.FC = () => {
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
      const mockSummary = `Summary of selected text: "${selectedText.substring(0, 50)}..."\n\nThis is a mock AI-generated summary. Key points:\n• Main topic identified\n• Important details extracted\n• Concise overview provided`;
      
      setResult(mockSummary);
      
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
      alert('Summary inserted successfully!');
    } catch (error) {
      console.error('Failed to insert summary:', error);
      alert('Failed to insert summary');
    }
  };

  return (
    <div className="summarize-panel" style={{ padding: '20px' }}>
      <div className="panel-header" style={{ marginBottom: '20px' }}>
        <h3>Summarize Text</h3>
        <p>Select text in your document and click summarize to get an AI-generated summary.</p>
      </div>
      
      <button 
        onClick={handleSummarize}
        disabled={loading}
        className="action-button primary"
        style={{ marginBottom: '16px' }}
      >
        {loading ? 'Summarizing...' : 'Summarize Selected Text'}
      </button>
      
      {result && (
        <div className="result-container" style={{ 
          marginTop: '20px',
          padding: '16px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600' }}>Summary:</h4>
          <div className="result-text" style={{ 
            marginBottom: '16px',
            fontSize: '14px',
            lineHeight: '1.4',
            whiteSpace: 'pre-wrap'
          }}>{result}</div>
          <div className="result-actions">
            <button 
              onClick={handleInsertSummary}
              className="action-button secondary"
              style={{ marginRight: '8px' }}
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
      
      // Mock API call - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 2500));
      
      const mockResult = {
        comparison: `Analysis of the two selected clauses reveals significant differences in scope and obligations. The first clause focuses on ${clause1.substring(0, 30)}... while the second addresses ${clause2.substring(0, 30)}...`,
        differences: [
          'Clause 1 uses more restrictive language',
          'Clause 2 includes broader scope provisions',
          'Different liability standards apply',
          'Termination conditions vary significantly'
        ],
        recommendations: [
          'Consider harmonizing the liability standards',
          'Clarify the scope overlap between clauses',
          'Review termination provisions for consistency',
          'Add cross-references between related sections'
        ]
      };
      
      setResult(mockResult);
      
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
      const comparisonText = `\n\nClause Comparison:\n${result.comparison}\n\nKey Differences:\n${result.differences.map((diff: string) => `• ${diff}`).join('\n')}\n\nRecommendations:\n${result.recommendations.map((rec: string) => `• ${rec}`).join('\n')}`;
      await wordApi.insertTextAtSelection(comparisonText);
      handleReset();
      alert('Comparison inserted successfully!');
    } catch (error) {
      console.error('Failed to insert comparison:', error);
      alert('Failed to insert comparison');
    }
  };

  const handleReset = () => {
    setClause1('');
    setClause2('');
    setResult(null);
    setStep('select-first');
  };

  return (
    <div className="compare-panel" style={{ padding: '20px' }}>
      <div className="panel-header" style={{ marginBottom: '20px' }}>
        <h3>Compare Clauses</h3>
        <p>Select two clauses to compare and analyze their differences.</p>
      </div>
      
      <div className="compare-steps">
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
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>First Clause</span>
          </div>
          <button 
            onClick={() => handleSelectClause(1)}
            disabled={loading}
            className={`action-button ${step === 'select-first' ? 'primary' : clause1 ? 'success' : 'secondary'}`}
          >
            {clause1 ? 'First Clause Selected ✓' : 'Select First Clause'}
          </button>
          {clause1 && (
            <div className="clause-preview" style={{ 
              marginTop: '12px',
              padding: '12px',
              backgroundColor: 'white',
              borderRadius: '4px',
              border: '1px solid #dee2e6',
              fontSize: '13px'
            }}>
              <strong>Preview:</strong> {clause1.substring(0, 100)}...
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
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>Second Clause</span>
          </div>
          <button 
            onClick={() => handleSelectClause(2)}
            disabled={loading || !clause1}
            className={`action-button ${step === 'select-second' ? 'primary' : clause2 ? 'success' : 'secondary'}`}
          >
            {clause2 ? 'Second Clause Selected ✓' : 'Select Second Clause'}
          </button>
          {clause2 && (
            <div className="clause-preview" style={{ 
              marginTop: '12px',
              padding: '12px',
              backgroundColor: 'white',
              borderRadius: '4px',
              border: '1px solid #dee2e6',
              fontSize: '13px'
            }}>
              <strong>Preview:</strong> {clause2.substring(0, 100)}...
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
            }}>3</span>
            <span className="step-title" style={{ fontSize: '14px', fontWeight: '600' }}>Compare</span>
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
        <div className="result-container" style={{ 
          marginTop: '20px',
          padding: '16px',
          backgroundColor: '#f8f9fa',
          borderRadius: '8px',
          border: '1px solid #e9ecef'
        }}>
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600' }}>Comparison Result:</h4>
          <div className="result-text">
            <div className="comparison-section" style={{ marginBottom: '16px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Analysis:</h5>
              <p style={{ fontSize: '13px', lineHeight: '1.4' }}>{result.comparison}</p>
            </div>
            <div className="comparison-section" style={{ marginBottom: '16px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Key Differences:</h5>
              <ul style={{ margin: '0', paddingLeft: '20px' }}>
                {result.differences.map((diff: string, index: number) => (
                  <li key={index} style={{ marginBottom: '4px', fontSize: '13px', lineHeight: '1.4' }}>{diff}</li>
                ))}
              </ul>
            </div>
            <div className="comparison-section">
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Recommendations:</h5>
              <ul style={{ margin: '0', paddingLeft: '20px' }}>
                {result.recommendations.map((rec: string, index: number) => (
                  <li key={index} style={{ marginBottom: '4px', fontSize: '13px', lineHeight: '1.4' }}>{rec}</li>
                ))}
              </ul>
            </div>
          </div>
          <div className="result-actions" style={{ display: 'flex', gap: '8px', marginTop: '16px' }}>
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

const RedraftPanel: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [originalText, setOriginalText] = useState<string>('');
  const [instructions, setInstructions] = useState<string>('');
  const [result, setResult] = useState<any>(null);
  const [showSuggestion, setShowSuggestion] = useState(false);
  const [selectedRange, setSelectedRange] = useState<any>(null);

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
          setSelectedRange(selection);
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
        redraftedText: `[Placeholder text for replacement]]`,
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
      setSelectedRange(null);
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
    setSelectedRange(null);
  };

  return (
    <div className="redraft-panel" style={{ padding: '20px' }}>
      <div className="panel-header" style={{ marginBottom: '20px' }}>
        <h3>Redraft Text</h3>
        <p>Select text to redraft with AI assistance. Add optional instructions for specific style or requirements.</p>
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
              minHeight: '60px'
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
          <h4 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: '600' }}>AI Suggestion:</h4>
          
          <div className="suggestion-comparison" style={{ marginBottom: '16px' }}>
            <div className="original-text" style={{ marginBottom: '12px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Original:</h5>
              <div className="text-block original" style={{ 
                marginBottom: '12px',
                padding: '12px',
                borderRadius: '4px',
                fontSize: '13px',
                lineHeight: '1.4',
                whiteSpace: 'pre-wrap',
                backgroundColor: '#f8d7da',
                border: '1px solid #f5c6cb'
              }}>{originalText}</div>
            </div>
            
            <div className="redrafted-text">
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Redrafted:</h5>
              <div className="text-block redrafted" style={{ 
                marginBottom: '12px',
                padding: '12px',
                borderRadius: '4px',
                fontSize: '13px',
                lineHeight: '1.4',
                whiteSpace: 'pre-wrap',
                backgroundColor: '#d4edda',
                border: '1px solid #c3e6cb'
              }}>{result.redraftedText}</div>
            </div>
          </div>

          {result.changes && result.changes.length > 0 && (
            <div className="changes-section" style={{ marginBottom: '16px' }}>
              <h5 style={{ margin: '0 0 8px 0', fontSize: '13px', fontWeight: '600' }}>Key Changes:</h5>
              <ul style={{ margin: '0', paddingLeft: '20px' }}>
                {result.changes.map((change: string, index: number) => (
                  <li key={index} style={{ marginBottom: '4px', fontSize: '13px', lineHeight: '1.4' }}>{change}</li>
                ))}
              </ul>
            </div>
          )}

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

interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
}

const AssistantChat: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [currentPrompt, setCurrentPrompt] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [includeSelection, setIncludeSelection] = useState(true);
  const [includeDocument, setIncludeDocument] = useState(false);

  const handleSendMessage = async () => {
    if (!currentPrompt.trim()) return;
    
    const userMessage: ChatMessage = {
      role: 'user',
      content: currentPrompt,
      timestamp: new Date()
    };

    setMessages(prev => [...prev, userMessage]);
    setCurrentPrompt('');
    setLoading(true);

    try {
      let context = '';
      let selectedText = '';

      if (includeSelection) {
        try {
          selectedText = await wordApi.getSelectedText();
        } catch (error) {
          console.warn('Could not get selected text:', error);
        }
      }

      if (includeDocument) {
        try {
          context = await wordApi.getEntireDocumentText();
        } catch (error) {
          console.warn('Could not get document text:', error);
        }
      }

      // Mock API call - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      let mockResponse = `Here's my analysis of your query: "${currentPrompt}"\n\n`;
      
      if (selectedText) {
        mockResponse += `Based on the selected text: "${selectedText.substring(0, 50)}..."\n\n`;
      }
      
      if (includeDocument) {
        mockResponse += `Considering the full document context...\n\n`;
      }
      
      mockResponse += `[This is a mock AI response. In production, this would be Oliver's intelligent analysis of your legal document and query.]`;

      const assistantMessage: ChatMessage = {
        role: 'assistant',
        content: mockResponse,
        timestamp: new Date()
      };

      setMessages(prev => [...prev, assistantMessage]);
      
    } catch (error) {
      console.error('Chat error:', error);
      const errorMessage: ChatMessage = {
        role: 'assistant',
        content: 'Sorry, I encountered an error processing your request. Please try again.',
        timestamp: new Date()
      };
      setMessages(prev => [...prev, errorMessage]);
    } finally {
      setLoading(false);
    }
  };

  const handleKeyPress = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  const handleInsertResponse = async (content: string) => {
    try {
      await wordApi.insertTextAtSelection(`\n\n${content}`);
      alert('Response inserted successfully!');
    } catch (error) {
      console.error('Failed to insert response:', error);
      alert('Failed to insert response');
    }
  };

  const handleClearChat = () => {
    setMessages([]);
  };

  const formatTime = (date: Date) => {
    return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  };

  return (
    <div className="assistant-chat" style={{ padding: '20px' }}>
      <div className="panel-header" style={{ marginBottom: '20px' }}>
        <h3>AI Assistant</h3>
        <p>Chat with Oliver about your document. Ask questions, request analysis, or get help with legal tasks.</p>
      </div>

      <div className="chat-options" style={{ 
        marginBottom: '16px',
        padding: '12px',
        backgroundColor: '#f8f9fa',
        borderRadius: '6px',
        border: '1px solid #e9ecef'
      }}>
        <label className="option-label" style={{ 
          display: 'flex',
          alignItems: 'center',
          marginBottom: '8px',
          fontSize: '14px',
          cursor: 'pointer'
        }}>
          <input
            type="checkbox"
            checked={includeSelection}
            onChange={(e) => setIncludeSelection(e.target.checked)}
            style={{ marginRight: '8px' }}
          />
          Include selected text
        </label>
        <label className="option-label" style={{ 
          display: 'flex',
          alignItems: 'center',
          fontSize: '14px',
          cursor: 'pointer'
        }}>
          <input
            type="checkbox"
            checked={includeDocument}
            onChange={(e) => setIncludeDocument(e.target.checked)}
            style={{ marginRight: '8px' }}
          />
          Include entire document
        </label>
      </div>

      <div className="chat-messages" style={{ 
        maxHeight: '300px',
        overflowY: 'auto',
        border: '1px solid #e9ecef',
        borderRadius: '8px',
        padding: '16px',
        marginBottom: '16px',
        backgroundColor: 'white'
      }}>
        {messages.length === 0 ? (
          <div className="empty-state" style={{ textAlign: 'center', padding: '40px 20px', color: '#6c757d' }}>
            <p>Start a conversation with Oliver!</p>
            <div className="suggested-prompts" style={{ marginTop: '20px' }}>
              <button 
                onClick={() => setCurrentPrompt('Analyze the selected text for potential legal issues')}
                className="prompt-suggestion"
                style={{ 
                  display: 'block',
                  width: '100%',
                  marginBottom: '8px',
                  padding: '12px',
                  backgroundColor: '#f8f9fa',
                  border: '1px solid #e9ecef',
                  borderRadius: '6px',
                  fontSize: '13px',
                  color: '#495057',
                  cursor: 'pointer'
                }}
              >
                Analyze for legal issues
              </button>
              <button 
                onClick={() => setCurrentPrompt('Summarize the key points of this document')}
                className="prompt-suggestion"
                style={{ 
                  display: 'block',
                  width: '100%',
                  marginBottom: '8px',
                  padding: '12px',
                  backgroundColor: '#f8f9fa',
                  border: '1px solid #e9ecef',
                  borderRadius: '6px',
                  fontSize: '13px',
                  color: '#495057',
                  cursor: 'pointer'
                }}
              >
                Summarize key points
              </button>
              <button 
                onClick={() => setCurrentPrompt('What are the main obligations in this contract?')}
                className="prompt-suggestion"
                style={{ 
                  display: 'block',
                  width: '100%',
                  marginBottom: '8px',
                  padding: '12px',
                  backgroundColor: '#f8f9fa',
                  border: '1px solid #e9ecef',
                  borderRadius: '6px',
                  fontSize: '13px',
                  color: '#495057',
                  cursor: 'pointer'
                }}
              >
                Find main obligations
              </button>
            </div>
          </div>
        ) : (
          messages.map((message, index) => (
            <div key={index} className={`message ${message.role}`} style={{ 
              marginBottom: '16px',
              padding: '12px',
              borderRadius: '8px',
              backgroundColor: message.role === 'user' ? '#e3f2fd' : '#f3e5f5',
              marginLeft: message.role === 'user' ? '20px' : '0',
              marginRight: message.role === 'assistant' ? '20px' : '0'
            }}>
              <div className="message-header" style={{ 
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: '8px'
              }}>
                <span className="message-role" style={{ fontWeight: '600', fontSize: '12px', color: '#495057' }}>
                  {message.role === 'user' ? 'You' : 'Oliver'}
                </span>
                <span className="message-time" style={{ fontSize: '11px', color: '#6c757d' }}>
                  {formatTime(message.timestamp)}
                </span>
              </div>
              <div className="message-content" style={{ 
                fontSize: '14px',
                lineHeight: '1.4',
                color: '#212529',
                whiteSpace: 'pre-wrap'
              }}>
                {message.content}
              </div>
              {message.role === 'assistant' && (
                <div className="message-actions" style={{ marginTop: '8px' }}>
                  <button 
                    onClick={() => handleInsertResponse(message.content)}
                    className="action-button small"
                    style={{ padding: '6px 12px', fontSize: '12px' }}
                  >
                    Insert into Document
                  </button>
                </div>
              )}
            </div>
          ))
        )}
        {loading && (
          <div className="message assistant" style={{ 
            marginBottom: '16px',
            padding: '12px',
            borderRadius: '8px',
            backgroundColor: '#f3e5f5',
            marginRight: '20px'
          }}>
            <div className="message-header" style={{ 
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              marginBottom: '8px'
            }}>
              <span className="message-role" style={{ fontWeight: '600', fontSize: '12px', color: '#495057' }}>Oliver</span>
            </div>
            <div className="message-content">
              <div className="typing-indicator" style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <span style={{ 
                  height: '8px',
                  width: '8px',
                  backgroundColor: '#6c757d',
                  borderRadius: '50%',
                  display: 'inline-block',
                  animation: 'typing 1.4s infinite ease-in-out',
                  animationDelay: '-0.32s'
                }}></span>
                <span style={{ 
                  height: '8px',
                  width: '8px',
                  backgroundColor: '#6c757d',
                  borderRadius: '50%',
                  display: 'inline-block',
                  animation: 'typing 1.4s infinite ease-in-out',
                  animationDelay: '-0.16s'
                }}></span>
                <span style={{ 
                  height: '8px',
                  width: '8px',
                  backgroundColor: '#6c757d',
                  borderRadius: '50%',
                  display: 'inline-block',
                  animation: 'typing 1.4s infinite ease-in-out'
                }}></span>
              </div>
            </div>
          </div>
        )}
      </div>

      <div className="chat-input" style={{ borderTop: '1px solid #e9ecef', paddingTop: '16px' }}>
        <textarea
          value={currentPrompt}
          onChange={(e) => setCurrentPrompt(e.target.value)}
          onKeyPress={handleKeyPress}
          placeholder="Ask Oliver about your document..."
          className="prompt-input"
          rows={3}
          disabled={loading}
          style={{ 
            width: '100%',
            padding: '12px',
            border: '1px solid #ced4da',
            borderRadius: '6px',
            fontSize: '14px',
            fontFamily: 'inherit',
            resize: 'vertical',
            minHeight: '60px',
            marginBottom: '12px'
          }}
        />
        <div className="input-actions" style={{ 
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center'
        }}>
          <button 
            onClick={handleClearChat}
            className="action-button tertiary small"
            disabled={loading || messages.length === 0}
            style={{ padding: '6px 12px', fontSize: '12px' }}
          >
            Clear Chat
          </button>
          <button 
            onClick={handleSendMessage}
            disabled={loading || !currentPrompt.trim()}
            className="action-button primary"
          >
            {loading ? 'Sending...' : 'Send'}
          </button>
        </div>
      </div>
    </div>
  );
};

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
      <div className="sidebar-header">
        <h2>Oliver AI Assistant</h2>
        <p>Your AI-powered legal workflow assistant</p>
        {isInitialized && (
          <small style={{ opacity: 0.8 }}>✅ Connected to Word</small>
        )}
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
        {activePanel === 'summarize' && <SummarizePanel />}
        {activePanel === 'compare' && <ComparePanel />}
        {activePanel === 'redraft' && <RedraftPanel />}
        {activePanel === 'chat' && <AssistantChat />}
      </div>
    </div>
  );
};

export default Sidebar;