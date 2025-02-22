require('dotenv').config();
const express = require('express');
const path = require('path');
const axios = require('axios');

// Initialize Express app
const app = express();

// Basic CORS middleware
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    if (req.method === 'OPTIONS') {
        return res.sendStatus(200);
    }
    next();
});

app.use(express.json());

// Serve static files from both root and src directories
app.use(express.static(path.join(__dirname)));
app.use(express.static(path.join(__dirname, 'src')));

// Serve index.html for root path
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'src', 'index.html'));
});

// Validate Claude API key middleware
const validateApiKey = (req, res, next) => {
    const apiKey = process.env.CLAUDE_API_KEY;
    if (!apiKey) {
        console.error('CLAUDE_API_KEY not found in environment variables');
        return res.status(500).json({ error: 'Claude API key not configured' });
    }
    next();
};

// API endpoint for processing emails
app.post('/api/process-email', validateApiKey, async (req, res) => {
    try {
        console.log('Received request to /api/process-email');
        
        const { emailContent, action = 'analyze', tone } = req.body;
        
        if (!emailContent) {
            return res.status(400).json({
                error: 'Email content is required'
            });
        }

        console.log('Calling Claude API with action:', action);
        
        // Call Claude API
        const response = await axios.post('https://api.anthropic.com/v1/messages', {
            model: 'claude-3-5-sonnet-20241022',
            max_tokens: 1000,
            messages: [{
                role: 'user',
                content: `You are writing a reply to this email. Generate a single ${tone || 'professional'} response that ${action.toLowerCase()}. Use simple line breaks to separate:
                - The greeting on its own line
                - The message body with appropriate paragraph breaks
                - The closing on its own line
                Do not include any explanations or notes.

                Email to reply to:
                ${emailContent}`
            }]
        }, {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${process.env.CLAUDE_API_KEY}`,
                'anthropic-version': '2023-06-01',
                'x-api-key': process.env.CLAUDE_API_KEY
            },
            timeout: 30000 // 30 second timeout
        });

        if (!response.data || !response.data.content || !response.data.content[0]) {
            throw new Error('Invalid response from Claude API');
        }

        console.log('Claude API response received');
        res.json({
            response: response.data.content[0].text
        });

    } catch (error) {
        console.error('Error processing email:', error);
        
        if (error.response) {
            const status = error.response.status;
            const errorData = error.response.data;
            
            if (status === 401) {
                return res.status(401).json({ 
                    error: 'Invalid API key',
                    details: errorData.error?.message
                });
            } else if (status === 429) {
                return res.status(429).json({ 
                    error: 'Rate limit exceeded',
                    details: errorData.error?.message
                });
            } else if (status === 400) {
                return res.status(400).json({ 
                    error: 'Bad request',
                    details: errorData.error?.message || 'Invalid request to Claude API'
                });
            }
        }
        
        if (error.code === 'ECONNABORTED') {
            return res.status(504).json({ 
                error: 'Request to Claude API timed out',
                details: error.message
            });
        }

        if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
            return res.status(503).json({ 
                error: 'Could not connect to Claude API',
                details: error.message
            });
        }

        res.status(500).json({
            error: 'Failed to process email',
            details: error.message,
            apiError: error.response?.data?.error
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'ok',
        apiConfigured: !!process.env.CLAUDE_API_KEY
    });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Unhandled error:', err);
    res.status(500).json({ 
        error: 'An unexpected error occurred',
        message: process.env.NODE_ENV === 'development' ? err.message : 'Internal server error'
    });
});

// Start server
const port = process.env.PORT || 8080; // Azure Web Apps expects port 8080
app.listen(port, () => {
    console.log(`Server running on port ${port}`);
    console.log(`Claude API ${process.env.CLAUDE_API_KEY ? 'configured' : 'not configured'}`);
});