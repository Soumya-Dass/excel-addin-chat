/* global console, Excel, Office */

// Import the services for webpack bundling
import './dataService.js';
import './aiService.js';
import './uiService.js';

// Enhanced Excel Data Assistant - Main Orchestrator
// Uses DataService, AIService, and UIService for clean separation of concerns

// Service instances
let dataService;
let aiService;
let uiService;

// Initialize Office Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeServices();
        console.log('Enhanced Excel Chatbot initialized successfully');
    }
});

// Initialize all services and set up the application
function initializeServices() {
    try {
        // Initialize services
        dataService = new window.DataService();
        aiService = new window.AIService();
        uiService = new window.UIService();
        
        // Initialize AI service
        aiService.initializeGemini();
        
        // Set up UI callbacks
        uiService.setOnSendMessageCallback(handleSendMessage);
        uiService.setOnClearConversationCallback(handleClearConversation);
        uiService.setOnDataSourceToggleCallback(() => uiService.handleDataSourceToggle());
        
        // Initialize UI
        uiService.setupEventListeners();
        uiService.showWelcomeMessage();
        uiService.updateStatus('Ready to analyze your data');
        
    } catch (error) {
        console.error('Error initializing services:', error);
        if (uiService) {
            uiService.showError('Failed to initialize application: ' + error.message);
        }
    }
}

// Main message handling orchestration
async function handleSendMessage() {
    const message = uiService.getUserInput();
    if (!message) return;
    
    // Clear input and disable controls
    uiService.clearUserInput();
    uiService.setControlsEnabled(false);
    
    // Add user message to chat
    uiService.addChatMessage(message, true);
    
    // Add to conversation history
    aiService.addToConversationHistory('user', message);
    
    // Show loading state
    uiService.showLoading(true);
    uiService.updateStatus('Reading Excel data...');
    
    try {
        // Read Excel data
        const shouldUseSelection = uiService.getUseSelectionState();
        const worksheetData = await dataService.readCurrentWorksheetDataEnhanced(shouldUseSelection);
        
        // Process with AI
        uiService.updateStatus('Analyzing data...');
        const response = await aiService.askGeminiWithContext(message, worksheetData);
        
        // Add AI response to conversation history
        aiService.addToConversationHistory('assistant', response);
        
        // Prepare enhanced data info for UI indicators
        const dataInfo = {
            isSampled: worksheetData.isSampled || false,
            hasFormulas: worksheetData.structuredData?.formulaAnalysis?.hasFormulas || false,
            multiSheet: worksheetData.workbookData?.totalSheets > 1 || false,
            hasObjects: (worksheetData.dataMetadata?.tables?.length > 0) || 
                       (worksheetData.dataMetadata?.pivotTables?.length > 0) || 
                       (worksheetData.dataMetadata?.charts?.length > 0) || false
        };
        
        // Display response with enhanced indicators
        const conversationLength = aiService.getConversationLength();
        uiService.addChatMessage(response, false, false, conversationLength, dataInfo);
        
        // Clean status message
        let statusMessage = `Analysis complete (${conversationLength} exchanges)`;
        if (dataInfo.isSampled) statusMessage += ' • Sampled';
        statusMessage += ' • Ready';
        
        uiService.updateStatus(statusMessage);
        
    } catch (error) {
        console.error('Error processing message:', error);
        uiService.addChatMessage('Sorry, I encountered an error processing your request. Please try again.', false, true);
        uiService.updateStatus('Error occurred - please try again');
    } finally {
        // Re-enable controls
        uiService.showLoading(false);
        uiService.setControlsEnabled(true);
        uiService.focusUserInput();
    }
}

// Handle clear conversation
function handleClearConversation() {
    // Clear AI conversation
    aiService.clearConversation();
    
    // Clear UI
    uiService.clearChatMessages();
    uiService.showWelcomeMessage();
    uiService.updateStatus('Conversation cleared - ready for new questions');
    
    // Clear data service cache
    dataService.clearCurrentData();
    
    console.log('Conversation and enhanced data cache cleared');
}