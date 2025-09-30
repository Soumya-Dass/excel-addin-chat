// Import Markdown parser
import { marked } from 'marked';

// UIService - Handles all user interface interactions and DOM manipulation
class UIService {
    constructor() {
        this.setupMarkdown();
        this.onSendMessageCallback = null;
        this.onClearConversationCallback = null;
        this.onDataSourceToggleCallback = null;
    }

    // Set up markdown options
    setupMarkdown() {
        marked.setOptions({
            breaks: true,
            gfm: true,
            sanitize: false
        });
    }

    // Set up event listeners
    setupEventListeners() {
        const userInput = document.getElementById('user-input');
        const sendBtn = document.getElementById('send-btn');
        
        if (sendBtn && this.onSendMessageCallback) {
            sendBtn.addEventListener('click', this.onSendMessageCallback);
        }
        
        if (userInput && this.onSendMessageCallback) {
            userInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    this.onSendMessageCallback();
                }
            });
            
            userInput.focus();
        }
        
        this.addClearConversationButton();
        this.setupDataSourceToggle();
    }

    // Set callback functions
    setOnSendMessageCallback(callback) {
        this.onSendMessageCallback = callback;
    }

    setOnClearConversationCallback(callback) {
        this.onClearConversationCallback = callback;
    }

    setOnDataSourceToggleCallback(callback) {
        this.onDataSourceToggleCallback = callback;
    }

    // Add clear conversation button
    addClearConversationButton() {
        const inputContainer = document.querySelector('.chat-input-container');
        if (!inputContainer) return;
        
        const clearBtn = document.createElement('button');
        clearBtn.innerHTML = 'New Conversation';
        clearBtn.className = 'clear-btn';
        clearBtn.style.cssText = `
            width: 100%;
            padding: 8px;
            margin-top: 8px;
            background: #404040;
            color: #ffffff;
            border: 1px solid #606060;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
            transition: background-color 0.2s;
        `;
        
        clearBtn.addEventListener('mouseover', () => {
            clearBtn.style.backgroundColor = '#ff6b35';
        });
        
        clearBtn.addEventListener('mouseout', () => {
            clearBtn.style.backgroundColor = '#404040';
        });
        
        if (this.onClearConversationCallback) {
            clearBtn.addEventListener('click', this.onClearConversationCallback);
        }
        
        inputContainer.appendChild(clearBtn);
    }

    // Setup data source toggle functionality
    setupDataSourceToggle() {
        const useSelectionCheckbox = document.getElementById('use-selection');
        if (useSelectionCheckbox && this.onDataSourceToggleCallback) {
            useSelectionCheckbox.addEventListener('change', this.onDataSourceToggleCallback);
        }
    }

    // Show clean welcome message
    showWelcomeMessage() {
        const chatMessages = document.getElementById('chat-messages');
        if (!chatMessages) return;
        
        const welcomeHtml = `
            <div class="welcome-message">
                <h3>Excel Data Assistant</h3>
                <p>Ask questions about your Excel data and get intelligent analysis.</p>
                
                <div style="text-align: left; max-width: 300px; margin: 20px auto;">
                    <h4 style="color: #ff8c42; font-size: 14px; margin: 15px 0 10px 0; font-weight: 600;">Example Questions:</h4>
                    <ul style="font-size: 13px; line-height: 1.6; margin: 0; padding-left: 18px; color: #ccc;">
                        <li>"What's the total revenue for Q4?"</li>
                        <li>"Show me the trend in expenses"</li>
                        <li>"Compare this year vs last year"</li>
                        <li>"What are the key metrics in this sheet?"</li>
                    </ul>
                </div>
            </div>
        `;
        chatMessages.innerHTML = welcomeHtml;
    }

    // Add chat message with enhanced markdown support and data indicators
    addChatMessage(message, isUser = false, isError = false, conversationLength = 0, dataInfo = null) {
        const chatMessages = document.getElementById('chat-messages');
        if (!chatMessages) return;
        
        const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${isUser ? 'user' : 'assistant'}`;
        
        const bubbleDiv = document.createElement('div');
        bubbleDiv.className = `message-bubble ${isError ? 'error-message' : ''}`;
        
        if (!isUser && !isError) {
            try {
                bubbleDiv.innerHTML = marked.parse(message);
            } catch (error) {
                console.warn('Markdown parsing failed, falling back to plain text:', error);
                bubbleDiv.textContent = message;
            }
        } else {
            bubbleDiv.textContent = message;
        }
        
        const timeDiv = document.createElement('div');
        timeDiv.className = 'message-time';
        timeDiv.textContent = timestamp;
        
        // Clean context indicators
        if (!isUser && !isError && conversationLength > 1) {
            const contextIndicator = document.createElement('div');
            contextIndicator.style.cssText = `
                font-size: 10px;
                color: #999;
                margin-top: 4px;
                opacity: 0.7;
            `;
            contextIndicator.textContent = `Context: ${conversationLength} exchanges`;
            timeDiv.appendChild(contextIndicator);
        }
        
        messageDiv.appendChild(bubbleDiv);
        messageDiv.appendChild(timeDiv);
        chatMessages.appendChild(messageDiv);
        
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    // Update status message
    updateStatus(message) {
        const statusText = document.getElementById('status-text');
        if (statusText) {
            statusText.textContent = message;
        }
    }

    // Show/hide loading overlay
    showLoading(show) {
        const loadingOverlay = document.getElementById('loading-overlay');
        if (loadingOverlay) {
            loadingOverlay.style.display = show ? 'flex' : 'none';
        }
    }

    // Enable/disable UI controls
    setControlsEnabled(enabled) {
        const userInput = document.getElementById('user-input');
        const sendBtn = document.getElementById('send-btn');
        
        if (userInput) userInput.disabled = !enabled;
        if (sendBtn) sendBtn.disabled = !enabled;
    }

    // Show clean error message
    showError(message) {
        let cleanMessage = message;
        
        // Add helpful context for common errors
        if (message.includes('API key')) {
            cleanMessage += '\n\n**Troubleshooting:** Please ensure your Gemini API key is configured in your .env file.';
        } else if (message.includes('quota')) {
            cleanMessage += '\n\n**Troubleshooting:** API quota exceeded. Please try again later or check your billing settings.';
        } else if (message.includes('Excel data')) {
            cleanMessage += '\n\n**Troubleshooting:** Try selecting a data range first, or ensure your worksheet contains data.';
        }
        
        this.addChatMessage(cleanMessage, false, true);
        this.updateStatus('Error occurred - see message for details');
    }

    // Get user input value
    getUserInput() {
        const userInput = document.getElementById('user-input');
        return userInput ? userInput.value.trim() : '';
    }

    // Clear user input
    clearUserInput() {
        const userInput = document.getElementById('user-input');
        if (userInput) {
            userInput.value = '';
        }
    }

    // Focus on user input
    focusUserInput() {
        const userInput = document.getElementById('user-input');
        if (userInput) {
            userInput.focus();
        }
    }

    // Clear chat messages
    clearChatMessages() {
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }
    }

    // Get data source checkbox state
    getUseSelectionState() {
        const useSelectionCheckbox = document.getElementById('use-selection');
        return useSelectionCheckbox ? useSelectionCheckbox.checked : true;
    }

    // Handle data source toggle
    handleDataSourceToggle() {
        if (this.getUseSelectionState()) {
            this.updateStatus('Using selected range (or auto-detect if no selection)');
        } else {
            this.updateStatus('Using full sheet with intelligent sampling');
        }
    }
}

// Export for use in other modules
if (typeof window !== 'undefined') {
    window.UIService = UIService;
}