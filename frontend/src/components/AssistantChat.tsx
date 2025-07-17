import React, { useState } from 'react';
import { wordApi } from '../office/wordApi';

interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
}

const AssistantChat: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [currentPrompt, setCurrentPrompt] = useState<string>('');
  const [loading, setLoading] = useState(false);
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
      try {
        await wordApi.getSelectedText();
      } catch (error) {
        console.warn('Could not get selected text:', error);
      }

      if (includeDocument) {
        try {
          await wordApi.getEntireDocumentText();
        } catch (error) {
          console.warn('Could not get document text:', error);
        }
      }

      // Mock API call - replace with actual API call later
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const mockResponse = `Lorem ipsum dolor sit amet.`;

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

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
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
        <p>Chat with Oliver about your document.</p>
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
        maxHeight: '200px',
        overflowY: 'auto',
        border: '1px solid #e9ecef',
        borderRadius: '8px',
        padding: '16px',
        marginBottom: '16px',
        backgroundColor: 'white'
      }}>
        {messages.length === 0 ? (
          <div className="empty-state" style={{ textAlign: 'center', padding: '40px 20px', color: '#6c757d' }}>
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
              marginLeft: message.role === 'user' ? '5%' : '0',
              marginRight: message.role === 'assistant' ? '5%' : '0'
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
            </div>
          ))
        )}
        {loading && (
          <div className="message assistant" style={{ 
            marginBottom: '16px',
            padding: '12px',
            borderRadius: '8px',
            backgroundColor: '#f3e5f5',
            marginRight: '5%'
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
          onKeyDown={handleKeyDown}
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
            marginBottom: '12px',
            backgroundColor: 'white',
            color: 'black',
            boxSizing: 'border-box'
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

export default AssistantChat;