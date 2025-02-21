// Initialize Office.js
Office.onReady((info) => {
    console.log('[Office.onReady] Office.js is ready', info);
    if (info.host === Office.HostType.Outlook) {
        // Wait for both DOM and Fluent UI
        if (document.readyState === 'loading') {
            console.log('[Office.onReady] DOM still loading, waiting...');
            document.addEventListener('DOMContentLoaded', () => {
                console.log('[Office.onReady] DOMContentLoaded fired');
                waitForFluent();
            });
        } else {
            console.log('[Office.onReady] DOM already loaded');
            waitForFluent();
        }
    }
});

// Wait for Fluent UI components to be defined
function waitForFluent() {
    console.log('[waitForFluent] Checking for Fluent UI components...');
    
    // Check for multiple Fluent UI components to ensure everything is loaded
    if (customElements.get('fluent-button') && 
        customElements.get('fluent-tab') && 
        customElements.get('fluent-select')) {
        console.log('[waitForFluent] All Fluent UI components ready');
        initializeUI();
    } else {
        console.log('[waitForFluent] Some Fluent UI components not ready, retrying...');
        setTimeout(waitForFluent, 100);
    }
}

// Initialize UI
function initializeUI() {
    console.log('[initializeUI] Starting initialization');
    
    try {
        // Get UI elements
        const loadingDiv = document.getElementById('loading');
        const resultDiv = document.getElementById('result-container');
        const toneSelect = document.getElementById('tone-select');
        
        if (!loadingDiv || !resultDiv || !toneSelect) {
            throw new Error('Required UI elements not found');
        }

        // Reset UI state
        loadingDiv.style.display = 'none';
        resultDiv.style.display = 'none';

        // Add event listeners to buttons without specific handlers
        const buttonsWithoutHandlers = document.querySelectorAll('fluent-button:not([id^="btn-yes"]):not([id^="btn-no"]):not([id^="btn-agree"]):not([id^="btn-disagree"]):not([id^="btn-need-time"]):not([id^="btn-clarify"]):not([id^="btn-suggest-meeting"]):not([id^="btn-project-update"]):not([id^="btn-status-report"]):not([id^="btn-meeting-invite"]):not([id^="btn-schedule-meeting"]):not([id^="btn-follow-up"]):not([id^="btn-introduction"]):not([id^="btn-request-info"])');
        
        console.log(`[initializeUI] Found ${buttonsWithoutHandlers.length} buttons without specific handlers`);
        
        buttonsWithoutHandlers.forEach(button => {
            const id = button.id;
            if (!id) {
                console.warn('[initializeUI] Button found without ID');
                return;
            }

            const action = id.replace('btn-', '');
            console.log(`[initializeUI] Adding listener for ${action}`);
            
            button.addEventListener('click', (event) => {
                event.preventDefault();
                console.log(`[Click Event] Button clicked: ${action}`);
                handleAction(action);
            });
        });

        // Reply Actions with specific handlers
        document.getElementById('btn-yes').onclick = (e) => {
            e.preventDefault();
            handleAction('says yes to');
        };
        document.getElementById('btn-no').onclick = (e) => {
            e.preventDefault();
            handleAction('says no to');
        };
        document.getElementById('btn-agree').onclick = (e) => {
            e.preventDefault();
            handleAction('agrees with');
        };
        document.getElementById('btn-disagree').onclick = (e) => {
            e.preventDefault();
            handleAction('disagrees with');
        };
        document.getElementById('btn-need-time').onclick = (e) => {
            e.preventDefault();
            handleAction('requests more time to address');
        };
        document.getElementById('btn-clarify').onclick = (e) => {
            e.preventDefault();
            handleAction('asks for clarification about');
        };
        document.getElementById('btn-suggest-meeting').onclick = (e) => {
            e.preventDefault();
            handleAction('suggests a meeting to discuss');
        };

        // New Email Actions
        document.getElementById('btn-project-update').onclick = (e) => {
            e.preventDefault();
            handleAction('Write a project update email');
        };
        document.getElementById('btn-status-report').onclick = (e) => {
            e.preventDefault();
            handleAction('Write a status report email');
        };
        document.getElementById('btn-meeting-invite').onclick = (e) => {
            e.preventDefault();
            handleAction('Write a meeting invitation email');
        };
        document.getElementById('btn-schedule-meeting').onclick = (e) => {
            e.preventDefault();
            handleAction('Write an email to schedule a meeting');
        };
        document.getElementById('btn-follow-up').onclick = (e) => {
            e.preventDefault();
            handleAction('Write a follow-up email');
        };
        document.getElementById('btn-introduction').onclick = (e) => {
            e.preventDefault();
            handleAction('Write an introduction email');
        };
        document.getElementById('btn-request-info').onclick = (e) => {
            e.preventDefault();
            handleAction('Write an email requesting information');
        };

        console.log('[initializeUI] Successfully initialized UI');
    } catch (error) {
        console.error('[initializeUI] Error during initialization:', error);
        displayError('Failed to initialize the add-in. Please refresh and try again.');
    }
}

// Action prompts map
const ACTION_PROMPTS = {
    'confirm-receipt': 'Generate a brief email confirming receipt of this message',
    'thank-you': 'Generate a brief thank you reply to this email',
    'summarize': 'Provide a concise summary of this email',
    'polite-decline': 'Generate a polite email declining or rejecting the request in this message',
    'need-time': 'Generate an email requesting more time to respond to this message',
    'clarify': 'Generate an email requesting clarification about points in this message',
    'suggest-meeting': 'Generate an email suggesting a meeting to discuss this further',
    'agree': 'Generate an email expressing agreement with this message',
    'proofread': 'Proofread and suggest improvements for this email',
    'extract-tasks': 'Extract and list all action items and tasks from this email',
    'detect-deadline': 'Identify and list all deadlines mentioned in this email',
    'flag-important': 'Identify the most important points from this email'
};

// Get email content
async function getEmailContent() {
    console.log('[getEmailContent] Getting email content');
    return new Promise((resolve, reject) => {
        try {
            if (!Office.context.mailbox.item) {
                console.error('[getEmailContent] No email selected');
                reject(new Error("No email selected"));
                return;
            }

            Office.context.mailbox.item.body.getAsync(
                "text",
                (result) => {
                    console.log('[getEmailContent] GetAsync result:', result);
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const subject = Office.context.mailbox.item.subject;
                        const emailContent = `Subject: ${subject}\n\nBody:\n${result.value}`;
                        resolve(emailContent);
                    } else {
                        console.error('[getEmailContent] Failed to get content:', result.error);
                        reject(new Error(`Failed to get email content: ${result.error.message}`));
                    }
                }
            );
        } catch (error) {
            console.error('[getEmailContent] Error in getEmailContent:', error);
            reject(new Error(`Error accessing email: ${error.message}`));
        }
    });
}

function getPrompt(action) {
    return ACTION_PROMPTS[action] || 'Please analyze this email';
}

// Handle actions
async function handleAction(action) {
    console.log(`[handleAction] Handling action: ${action}`);
    
    const loadingDiv = document.getElementById('loading');
    const resultDiv = document.getElementById('result-container');
    
    // Reset UI state
    if (loadingDiv) {
        loadingDiv.style.display = 'none';
    }
    if (resultDiv) {
        resultDiv.style.display = 'none';
        resultDiv.innerHTML = ''; // Clear previous results
    }
    
    showLoading(true);

    try {
        // Get email content
        const content = await getEmailContent();
        console.log('[handleAction] Got email content');

        if (!content || content.trim() === '') {
            throw new Error("No email content found to process");
        }

        // Process with Claude
        const response = await processWithClaude(content, action);
        console.log('[handleAction] Got response from Claude');

        // Show result
        if (resultDiv && response) {
            console.log('[handleAction] Displaying result');
            displayResult(response);
            
            const analysisActions = ['summarize', 'extractTasks', 'detectDeadline', 'flagImportant'];
            if (!analysisActions.includes(action)) {
                console.log('[handleAction] Opening reply window');
                // Convert newlines to HTML breaks for Outlook
                const htmlResponse = response.replace(/\n/g, '<br>');
                Office.context.mailbox.item.displayReplyForm(htmlResponse);
            }
        }
    } catch (error) {
        console.error('[handleAction] Error:', error);
        displayError(error.message);
    } finally {
        showLoading(false);
    }
}

// Format email for reply
function formatEmailForReply(text) {
    // Just handle basic spacing
    return text
        .replace(/\n{3,}/g, '\n\n') // Replace multiple newlines with double newlines
        .trim(); // Remove leading/trailing whitespace
}

// Process with Claude
async function processWithClaude(content, prompt) {
    console.log('[processWithClaude] Starting API call');
    try {
        const timestamp = new Date().getTime();
        const toneSelect = document.getElementById('tone-select');
        const tone = toneSelect ? toneSelect.value : 'professional';
        
        console.log('[processWithClaude] Using tone:', tone);
        
        const response = await fetch(`https://localhost:3000/api/process-email?t=${timestamp}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Cache-Control': 'no-cache'
            },
            body: JSON.stringify({
                emailContent: content,
                action: prompt,
                tone: tone
            })
        });

        console.log('[processWithClaude] Response status:', response.status);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('[processWithClaude] Server error:', errorText);
            throw new Error(`Server error: ${response.status} ${response.statusText}\n${errorText}`);
        }

        const data = await response.json();
        console.log('[processWithClaude] Got response data:', data);
        
        if (!data.response) {
            throw new Error('No response message received');
        }

        return formatEmailForReply(data.response);
    } catch (error) {
        console.error('[processWithClaude] Error:', error);
        throw error;
    }
}

// UI Helpers
function showLoading(show) {
    console.log('[showLoading] Setting loading:', show);
    const loadingElement = document.getElementById('loading');
    const resultContainer = document.getElementById('result-container');
    
    if (loadingElement && resultContainer) {
        if (show) {
            resultContainer.style.display = 'none';
            loadingElement.style.display = 'flex';
        } else {
            loadingElement.style.display = 'none';
        }
    } else {
        console.error('[showLoading] Loading or result container elements not found');
    }
}

function displayResult(message) {
    const resultDiv = document.getElementById('result-container');
    if (!resultDiv) return;

    // Simple newline to <br> conversion
    const formattedMessage = message.replace(/\n/g, '<br>');

    resultDiv.innerHTML = formattedMessage;
    resultDiv.style.display = 'block';
    resultDiv.style.whiteSpace = 'pre-wrap';
    resultDiv.style.lineHeight = '1.5';
    resultDiv.style.padding = '16px';
}

function displayError(message) {
    console.error('[displayError] Displaying error:', message);
    const container = document.getElementById('result-container');
    const result = document.getElementById('result');
    if (container && result) {
        result.textContent = message;
        result.style.color = '#d13438'; // Red color for errors
        container.style.display = 'block';
    } else {
        console.error('[displayError] Result elements not found');
    }
}