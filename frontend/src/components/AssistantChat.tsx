import React, { useState, useRef, useEffect } from 'react';
import { wordApi } from '../office/wordApi';
import { apiService, ChatMessage } from '../services/apiService';

const AssistantChat: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [currentPrompt, setCurrentPrompt] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [includeSelection, setIncludeSelection] = useState(true);
  const [includeDocument, setIncludeDocument] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

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

      const response = await apiService.analyze({
        text: selectedText,
        prompt: currentPrompt,
        context: context || undefined
      });

      const assistantMessage: ChatMessage = {
        role: 'assistant',
        content: response.response,
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
    } catch (error) {
      console.error('Failed to insert response:', error);
    }
  };

  const handleClearChat = () => {
    setMessages([]);
  };

  const formatTime = (date: Date) => {
    return date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  };

  return (
    <div className="assistant-chat">
      <div className="panel-header">
        <h3>AI Assistant</h3>
        <p>Chat with Oliver about your document. Ask questions, request analysis, or get help with legal tasks.</p>
      </div>

      <div className="chat-options">
        <label className="option-label">
          <input
            type="checkbox"
            checked={includeSelection}
            onChange={(e) => setIncludeSelection(e.target.checked)}
          />
          Include selected text
        </label>
        <label className="option-label">
          <input
            type="checkbox"
            checked={includeDocument}
            onChange={(e) => setIncludeDocument(e.target.checked)}
          />
          Include entire document
        </label>
      </div>

      <div className="chat-messages">
        {messages.length === 0 ? (
          <div className="empty-state">
            <p>Start a conversation with Oliver!</p>
            <div className="suggested-prompts">
              <button 
                onClick={() => setCurrentPrompt('Analyze the selected text for potential legal issues')}
                className="prompt-suggestion"
              >
                Analyze for legal issues
              </button>
              <button 
                onClick={() => setCurrentPrompt('Summarize the key points of this document')}
                className="prompt-suggestion"
              >
                Summarize key points
              </button>
              <button 
                onClick={() => setCurrentPrompt('What are the main obligations in this contract?')}
                className="prompt-suggestion"
              >
                Find main obligations
              </button>
            </div>
          </div>
        ) : (
          messages.map((message, index) => (
            <div key={index} className={`message ${message.role}`}>
              <div className="message-header">
                <span className="message-role">
                  {message.role === 'user' ? 'You' : 'Oliver'}
                </span>
                <span className="message-time">
                  {formatTime(message.timestamp)}
                </span>
              </div>
              <div className="message-content">
                {message.content}
              </div>
              {message.role === 'assistant' && (
                <div className="message-actions">
                  <button 
                    onClick={() => handleInsertResponse(message.content)}
                    className="action-button small"
                  >
                    Insert into Document
                  </button>
                </div>
              )}
            </div>
          ))
        )}
        {loading && (
          <div className="message assistant">
            <div className="message-header">
              <span className="message-role">Oliver</span>
            </div>
            <div className="message-content">
              <div className="typing-indicator">
                <span></span>
                <span></span>
                <span></span>
              </div>
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      <div className="chat-input">
        <textarea
          value={currentPrompt}
          onChange={(e) => setCurrentPrompt(e.target.value)}
          onKeyPress={handleKeyPress}
          placeholder="Ask Oliver about your document..."
          className="prompt-input"
          rows={3}
          disabled={loading}
        />
        <div className="input-actions">
          <button 
            onClick={handleClearChat}
            className="action-button tertiary small"
            disabled={loading || messages.length === 0}
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