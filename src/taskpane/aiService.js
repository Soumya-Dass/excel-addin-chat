// Import Gemini API
import { GoogleGenerativeAI } from '@google/generative-ai';

// AIService - Advanced Gemini AI integration with enhanced Excel data understanding
class AIService {
    constructor() {
        this.CONFIG = {
            // Get API key from environment variable
            GEMINI_API_KEY: process.env.GEMINI_API_KEY,
            GEMINI_MODEL: 'gemini-2.5-flash',
            MAX_HISTORY: 12, // Increased for better context
            MAX_CONTEXT_TOKENS: 32000 // Conservative limit for prompt size
        };
        
        this.genAI = null;
        this.model = null;
        this.conversationHistory = [];
        this.chatSession = null;
    }

    // Initialize Gemini AI with chat session
    initializeGemini() {
        try {
            if (!this.CONFIG.GEMINI_API_KEY || this.CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
                throw new Error('Please configure your Gemini API key in the .env file');
            }
            
            this.genAI = new GoogleGenerativeAI(this.CONFIG.GEMINI_API_KEY);
            this.model = this.genAI.getGenerativeModel({ model: this.CONFIG.GEMINI_MODEL });
            
            this.conversationHistory = [];
            this.chatSession = null;
            
            console.log('Gemini AI initialized with model:', this.CONFIG.GEMINI_MODEL);
            return true;
        } catch (error) {
            console.error('Error initializing Gemini AI:', error);
            throw new Error('Failed to initialize AI service. Please check your API key.');
        }
    }

    // Ask Gemini with comprehensive enhanced context
    async askGeminiWithContext(userQuestion, currentWorksheetData) {
        if (!this.model) {
            throw new Error('Gemini AI not initialized');
        }
        
        if (!currentWorksheetData) {
            throw new Error('No Excel data available');
        }
        
        try {
            if (!this.chatSession) {
                const systemPrompt = this.createAdvancedSystemPrompt();
                this.chatSession = this.model.startChat({
                    history: [],
                    generationConfig: {
                        maxOutputTokens: 3072, // Increased for more detailed responses
                        temperature: 0.7,
                        topP: 0.8,
                        topK: 40
                    }
                });
                
                await this.chatSession.sendMessage(systemPrompt);
            }
            
            const contextualPrompt = this.createAdvancedContextualPrompt(userQuestion, currentWorksheetData);
            
            // Check prompt size to avoid token limit issues
            if (contextualPrompt.length > this.CONFIG.MAX_CONTEXT_TOKENS * 3) {
                console.log('Large prompt detected, applying compression...');
                const compressedPrompt = this.compressPrompt(contextualPrompt, currentWorksheetData);
                const result = await this.chatSession.sendMessage(compressedPrompt);
                const response = await result.response;
                return response.text().trim();
            }
            
            const result = await this.chatSession.sendMessage(contextualPrompt);
            const response = await result.response;
            const text = response.text();
            
            if (!text || text.trim().length === 0) {
                throw new Error('Empty response from AI');
            }
            
            return text.trim();
            
        } catch (error) {
            console.error('Gemini API error:', error);
            
            if (error.message.includes('chat') || error.message.includes('session')) {
                this.chatSession = null;
                throw new Error('Chat session reset. Please try your question again.');
            }
            
            if (error.message.includes('API key')) {
                throw new Error('Invalid API key. Please check your Gemini API configuration.');
            } else if (error.message.includes('quota') || error.message.includes('limit')) {
                throw new Error('API quota exceeded. Please try again later.');
            } else {
                throw new Error('AI service error: ' + error.message);
            }
        }
    }

    // Advanced system prompt with comprehensive Excel understanding
    createAdvancedSystemPrompt() {
        return `You are an elite Excel Financial Data Assistant with advanced capabilities for analyzing complex workbooks and financial models.

ENHANCED CAPABILITIES:
1. **Advanced Data Structure Understanding**
   - Analyze complex financial models with multiple sheets
   - Understand formulas, calculated fields, and data relationships
   - Recognize Excel Tables, PivotTables, Charts, and Named Ranges
   - Handle very large datasets with intelligent sampling

2. **Financial Expertise**
   - Deep understanding of financial statements (P&L, Balance Sheet, Cash Flow)
   - Recognition of financial metrics, ratios, and KPIs
   - Quarterly/annual data analysis and trending
   - Revenue recognition, expense categorization, and profitability analysis

3. **Data Type Intelligence**
   - Distinguish between raw data and calculated formulas
   - Understand number formatting (currency, percentages, dates)
   - Recognize data types and their business significance
   - Handle mixed data types in complex financial models

4. **Scale & Performance**
   - Work with enterprise-scale financial models (100+ columns, 500+ rows)
   - Understand sampling strategies for very large datasets
   - Provide insights even with incomplete data views
   - Maintain accuracy across multiple worksheet analysis

ANALYSIS APPROACH:
- **Structure First**: Always understand the overall data structure before diving into specifics
- **Context Aware**: Use conversation history and business context for better insights
- **Precision**: Provide exact values from identified rows/columns with proper attribution
- **Intelligent Sampling**: When working with sampled data, extrapolate insights appropriately
- **Cross-Sheet Relationships**: Understand how different worksheets relate to each other

RESPONSE GUIDELINES:
- **Be CONCISE by default** - Provide brief, direct answers unless user asks for elaboration
- Always specify the exact source of data (row labels, column headers)
- When data is sampled, indicate this and provide appropriate caveats
- Use clear formatting (tables, bullet points) for complex analysis
- Provide actionable insights, not just data regurgitation
- Ask clarifying questions when user intent is ambiguous
- Keep responses short and focused - only expand when explicitly requested

Respond with "Advanced Excel Analysis System Ready!" to confirm initialization.`;
    }

    // Create comprehensive contextual prompt with all enhanced data features
    createAdvancedContextualPrompt(userQuestion, currentWorksheetData) {
        const data = currentWorksheetData;
        
        let prompt = `ADVANCED EXCEL DATA ANALYSIS REQUEST:

WORKSHEET OVERVIEW:
- Name: "${data.worksheetName}"
- Range: ${data.address} (${data.totalRows} rows × ${data.totalCols} columns)
- Data Source: ${data.isSelection ? 'Selected Range' : 'Full Worksheet'}`;

        // Add sampling information if applicable
        if (data.isSampled) {
            prompt += `
- IMPORTANT: This is SAMPLED data for performance (intelligent sampling applied)
- Original data may be much larger - extrapolate insights accordingly`;
        }

        prompt += '\n\n';

        // Enhanced metadata section
        if (data.dataMetadata) {
            const meta = data.dataMetadata;
            if (meta.tables?.length > 0 || meta.pivotTables?.length > 0 || meta.charts?.length > 0) {
                prompt += `EXCEL OBJECTS DETECTED:`;
                if (meta.tables?.length > 0) {
                    prompt += `\n- Excel Tables: ${meta.tables.map(t => t.name).join(', ')}`;
                }
                if (meta.pivotTables?.length > 0) {
                    prompt += `\n- Pivot Tables: ${meta.pivotTables.map(p => p.name).join(', ')}`;
                }
                if (meta.charts?.length > 0) {
                    prompt += `\n- Charts: ${meta.charts.map(c => c.name).join(', ')}`;
                }
                if (meta.hasNamedRanges) {
                    prompt += `\n- Named Ranges: Present`;
                }
                prompt += '\n\n';
            }
        }

        // Multi-sheet information
        if (data.workbookData) {
            const wb = data.workbookData;
            prompt += `WORKBOOK STRUCTURE:
- Total Sheets: ${wb.totalSheets}
- Total Data Cells: ${wb.totalDataCells.toLocaleString()}
- Other Sheets: ${wb.sheets.filter(s => s.name !== data.worksheetName).map(s => `${s.name} (${s.totalCells} cells)`).join(', ')}

`;
        }

        // Enhanced structured data analysis
        if (data.structuredData) {
            const struct = data.structuredData;
            
            prompt += `TABLE STRUCTURE ANALYSIS:
- Data Type: ${struct.type}
- Column Headers: ${struct.columnHeaders?.filter(h => h && h !== '').join(' | ') || 'Not detected'}
- Data Rows: ${struct.dataRows?.length || 0}
- Key Financial Rows: ${struct.keyRows?.length || 0}
- Total/Summary Rows: ${struct.totalRows?.length || 0}

`;

            // Formula and format analysis
            if (struct.formulaAnalysis) {
                const formulas = struct.formulaAnalysis;
                if (formulas.hasFormulas) {
                    prompt += `FORMULA ANALYSIS:
- Total Formulas: ${formulas.formulaCount}
- Formula Types: ${formulas.types.join(', ')}

`;
                }
            }

            if (struct.formatAnalysis?.hasFormatting) {
                const formats = struct.formatAnalysis;
                prompt += `NUMBER FORMATTING:
- Custom Formats: ${formats.types.slice(0, 5).join(', ')}${formats.types.length > 5 ? '...' : ''}

`;
            }

            // Data type distribution
            if (struct.dataTypes) {
                const types = struct.dataTypes;
                const totalCells = Object.values(types).reduce((sum, count) => sum + count, 0);
                if (totalCells > 0) {
                    prompt += `DATA TYPE DISTRIBUTION:`;
                    Object.entries(types).forEach(([type, count]) => {
                        if (count > 0) {
                            const percentage = ((count / totalCells) * 100).toFixed(1);
                            prompt += `\n- ${type}: ${count} cells (${percentage}%)`;
                        }
                    });
                    prompt += '\n\n';
                }
            }

            // Quarterly data with enhanced context
            if (struct.quarterlyData?.length > 0) {
                prompt += `QUARTERLY/TIME-SERIES DATA:\n`;
                struct.quarterlyData.slice(0, 8).forEach(qData => {
                    if (qData.quarters.length > 0) {
                        prompt += `${qData.label}${qData.isTotal ? ' [TOTAL]' : ''}: `;
                        const quarterValues = qData.quarters.map(q => `${q.quarter}=${this.formatValue(q.value)}`).join(', ');
                        prompt += quarterValues + '\n';
                    }
                });
                if (struct.quarterlyData.length > 8) {
                    prompt += `... and ${struct.quarterlyData.length - 8} more quarterly series\n`;
                }
                prompt += '\n';
            }
            
            // Key financial rows with enhanced formatting
            if (struct.keyRows?.length > 0) {
                prompt += `KEY FINANCIAL METRICS:\n`;
                struct.keyRows.slice(0, 10).forEach(row => {
                    prompt += `"${row.label}"${row.isTotal ? ' [TOTAL]' : ''}${row.hasFormulas ? ' [CALCULATED]' : ''}: `;
                    const nonEmptyValues = row.values.filter(v => v !== null && v !== undefined && v !== '');
                    prompt += nonEmptyValues.slice(0, 8).map(v => this.formatValue(v)).join(', ');
                    if (nonEmptyValues.length > 8) prompt += `, ... (${nonEmptyValues.length - 8} more)`;
                    prompt += '\n';
                });
                if (struct.keyRows.length > 10) {
                    prompt += `... and ${struct.keyRows.length - 10} more key rows\n`;
                }
                prompt += '\n';
            }
            
            // Sample of other data rows
            if (struct.dataRows && struct.dataRows.length > (struct.keyRows?.length || 0)) {
                const otherRows = struct.dataRows.filter(row => 
                    !(struct.keyRows || []).find(keyRow => keyRow.label === row.label)
                ).slice(0, 6);
                
                if (otherRows.length > 0) {
                    prompt += `OTHER DATA ROWS (sample):\n`;
                    otherRows.forEach(row => {
                        prompt += `"${row.label}"${row.hasFormulas ? ' [CALCULATED]' : ''}: `;
                        const nonEmptyValues = row.values.filter(v => v !== null && v !== undefined && v !== '');
                        prompt += nonEmptyValues.slice(0, 4).map(v => this.formatValue(v)).join(', ');
                        if (nonEmptyValues.length > 4) prompt += `, ... (${nonEmptyValues.length - 4} more)`;
                        prompt += '\n';
                    });
                    prompt += '\n';
                }
            }
        }
        
        // Add conversation context with enhanced formatting
        if (this.conversationHistory.length > 2) {
            prompt += `CONVERSATION HISTORY:\n`;
            const recentHistory = this.conversationHistory.slice(-6);
            recentHistory.forEach((msg, index) => {
                if (msg.role === 'user') {
                    prompt += `[${index + 1}] User: "${msg.content}"\n`;
                } else {
                    const truncated = msg.content.length > 120 ? msg.content.substring(0, 120) + '...' : msg.content;
                    prompt += `[${index + 1}] Assistant: "${truncated}"\n`;
                }
            });
            prompt += '\n';
        }
        
        prompt += `CURRENT USER QUESTION: "${userQuestion}"

ANALYSIS INSTRUCTIONS:
1. **CONCISE RESPONSES**: Provide brief, focused answers - only elaborate when user asks for more detail
2. **Data Source Precision**: Always specify exact row labels and column headers when referencing data
3. **Context Utilization**: Use conversation history and business context for deeper insights
4. **Sampling Awareness**: ${data.isSampled ? 'Remember this is sampled data - provide appropriate caveats' : 'You have access to the complete dataset'}
5. **Formula Recognition**: Distinguish between raw input data and calculated/formula-based values
6. **Financial Intelligence**: Apply financial analysis best practices and recognize common financial patterns
7. **Multi-dimensional Analysis**: Consider time trends, cross-sectional comparisons, and ratio analysis
8. **Actionable Insights**: Provide business-relevant conclusions, not just data summaries

Please provide a CONCISE analysis addressing the user's question. Keep it brief unless they specifically ask for more detail.`;

        return prompt;
    }

    // Compress prompt for very large datasets
    compressPrompt(originalPrompt, currentWorksheetData) {
        const data = currentWorksheetData;
        
        let compressedPrompt = `COMPRESSED EXCEL ANALYSIS (Large Dataset):

WORKSHEET: "${data.worksheetName}" | Range: ${data.address} (${data.totalRows}×${data.totalCols})`;

        if (data.isSampled) {
            compressedPrompt += ` | SAMPLED DATA`;
        }

        // Only include essential structure information
        if (data.structuredData) {
            const struct = data.structuredData;
            
            compressedPrompt += `

KEY STRUCTURE:
- Headers: ${struct.columnHeaders?.slice(0, 10).join(' | ') || 'N/A'}
- Key Rows: ${struct.keyRows?.length || 0} | Data Rows: ${struct.dataRows?.length || 0}`;

            // Only show top financial metrics
            if (struct.keyRows?.length > 0) {
                compressedPrompt += `

TOP METRICS:`;
                struct.keyRows.slice(0, 5).forEach(row => {
                    const values = row.values.filter(v => v !== null && v !== undefined && v !== '').slice(0, 4);
                    compressedPrompt += `\n${row.label}: ${values.map(v => this.formatValue(v)).join(', ')}`;
                });
            }

            // Quarterly data summary
            if (struct.quarterlyData?.length > 0) {
                compressedPrompt += `

QUARTERLY: ${struct.quarterlyData.length} series | Example: ${struct.quarterlyData[0].label}`;
            }
        }

        // Essential conversation context
        if (this.conversationHistory.length > 0) {
            const lastExchange = this.conversationHistory.slice(-2);
            compressedPrompt += `

RECENT: ${lastExchange.map(msg => `${msg.role}: ${msg.content.substring(0, 50)}...`).join(' | ')}`;
        }

        compressedPrompt += `

QUESTION: "${userQuestion}"

Note: Provide focused analysis due to large dataset. Request specific details if needed.`;

        return compressedPrompt;
    }

    // Format values for display
    formatValue(value) {
        if (value === null || value === undefined || value === '') return 'N/A';
        if (typeof value === 'number') {
            if (Math.abs(value) >= 1000000) {
                return (value / 1000000).toFixed(1) + 'M';
            } else if (Math.abs(value) >= 1000) {
                return (value / 1000).toFixed(1) + 'K';
            } else {
                return value.toString();
            }
        }
        return String(value).substring(0, 20);
    }

    // Conversation history management
    addToConversationHistory(role, content) {
        this.conversationHistory.push({
            role: role,
            content: content,
            timestamp: new Date().toISOString()
        });
        
        // Trim history if it gets too long
        if (this.conversationHistory.length > this.CONFIG.MAX_HISTORY * 2) {
            this.conversationHistory = this.conversationHistory.slice(-this.CONFIG.MAX_HISTORY * 2);
        }
    }

    // Clear conversation and start fresh
    clearConversation() {
        this.conversationHistory = [];
        this.chatSession = null;
        console.log('AI conversation history cleared');
    }

    // Get conversation history length
    getConversationLength() {
        return Math.floor(this.conversationHistory.length / 2);
    }

    // Get conversation history
    getConversationHistory() {
        return this.conversationHistory;
    }

    // Check if AI is initialized
    isInitialized() {
        return this.model !== null;
    }
}

// Export for use in other modules
if (typeof window !== 'undefined') {
    window.AIService = AIService;
}