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
        // Initialize button handlers
        initializeButtonHandlers();

        // Show initial tab
        showTab('reply');

        function initializeButtonHandlers() {
            // Reply tab buttons
            document.getElementById('btn-yes')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('says yes to', false);
            });
            document.getElementById('btn-no')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('says no to', false);
            });
            document.getElementById('btn-agree')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('agrees with', false);
            });
            document.getElementById('btn-disagree')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('disagrees with', false);
            });
            document.getElementById('btn-need-time')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('requests more time to address', false);
            });
            document.getElementById('btn-clarify')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('asks for clarification about', false);
            });
            document.getElementById('btn-suggest-meeting')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('suggests a meeting to discuss', false);
            });

            // Reply All tab buttons
            document.getElementById('btn-yes-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('says yes to', true);
            });
            document.getElementById('btn-no-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('says no to', true);
            });
            document.getElementById('btn-agree-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('agrees with', true);
            });
            document.getElementById('btn-disagree-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('disagrees with', true);
            });
            document.getElementById('btn-need-time-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('requests more time to address', true);
            });
            document.getElementById('btn-clarify-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('asks for clarification about', true);
            });
            document.getElementById('btn-suggest-meeting-all')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('suggests a meeting to discuss', true);
            });

            // New Email tab buttons
            document.getElementById('btn-project-update')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write a project update email');
            });
            document.getElementById('btn-status-report')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write a status report email');
            });
            document.getElementById('btn-meeting-invite')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write a meeting invitation email');
            });
            document.getElementById('btn-schedule-meeting')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write an email to schedule a meeting');
            });
            document.getElementById('btn-follow-up')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write a follow-up email');
            });
            document.getElementById('btn-introduction')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write an introduction email');
            });
            document.getElementById('btn-request-info')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('Write an email requesting information');
            });

            // Analyze tab buttons
            document.getElementById('btn-proofread')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('proofread');
            });
            document.getElementById('btn-summarize')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('summarize');
            });
            document.getElementById('btn-extract-tasks')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('extract-tasks');
            });
            document.getElementById('btn-detect-deadlines')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('detect-deadline');
            });
            document.getElementById('btn-flag-important')?.addEventListener('click', (e) => {
                e.preventDefault();
                handleAction('flag-important');
            });

            // Custom input buttons
            const customButtons = ['reply', 'reply-all', 'new', 'analyze'];
            customButtons.forEach(type => {
                const button = document.getElementById(`btn-custom-${type}`);
                if (button) {
                    button.addEventListener('click', (event) => {
                        event.preventDefault();
                        const input = document.getElementById(`custom-input-${type}`);
                        if (input && input.value.trim()) {
                            console.log(`[Click Event] Custom ${type} button clicked with instruction:`, input.value);
                            handleCustomAction(input.value.trim(), type);
                        }
                    });
                }
            });
        }

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

// Handle actions
// Handle actions
async function handleAction(action, isReplyAll = false) {
    console.log(`[handleAction] Handling action: ${action}, isReplyAll: ${isReplyAll}`);
    
    // Show the result container and loading indicator first
    const resultContainer = document.getElementById('result-container');
    if (resultContainer) {
        resultContainer.style.display = 'block';
    }
    
    showLoading(true);

    try {
        const content = await getEmailContent();
        console.log('[handleAction] Got email content');

        if (!content?.trim()) {
            throw new Error("No email content found to process");
        }

        const response = await processWithClaude(content, action);
        console.log('[handleAction] Got response from Claude');

        // Hide loading and show result
        showLoading(false);
        displayResult(response);

        // Wait a moment to let user see the result before opening email window
        const analysisActions = ['summarize', 'extractTasks', 'detectDeadline', 'flagImportant'];
        if (!analysisActions.includes(action)) {
            setTimeout(() => {
                console.log(`[handleAction] Opening ${isReplyAll ? 'reply-all' : 'reply'} window`);
                const htmlResponse = response.replace(/\n/g, '<br>');
                if (isReplyAll) {
                    Office.context.mailbox.item.displayReplyAllForm(htmlResponse);
                } else {
                    Office.context.mailbox.item.displayReplyForm(htmlResponse);
                }
            }, 1500); // Wait 1.5 seconds before opening email window
        }
    } catch (error) {
        console.error('[handleAction] Error:', error);
        showLoading(false);
        displayError(error.message);
    }
}

// Handle custom actions
// Handle custom actions
async function handleCustomAction(instruction, type) {
    console.log(`[handleCustomAction] Processing custom instruction for ${type}:`, instruction);
    
    // Show the result container and loading indicator first
    const resultContainer = document.getElementById('result-container');
    if (resultContainer) {
        resultContainer.style.display = 'block';
    }
    
    showLoading(true);

    try {
        const content = await getEmailContent();
        console.log('[handleCustomAction] Got email content');

        if (!content?.trim()) {
            throw new Error("No email content found to process");
        }

        const response = await processWithClaude(content, instruction);
        console.log('[handleCustomAction] Got response from Claude');

        // Hide loading and show result
        showLoading(false);
        displayResult(response);
            
        // Wait a moment to let user see the result before opening email window
        if (type !== 'analyze') {
            setTimeout(() => {
                const htmlResponse = response.replace(/\n/g, '<br>');
                switch(type) {
                    case 'reply':
                        Office.context.mailbox.item.displayReplyForm(htmlResponse);
                        break;
                    case 'reply-all':
                        Office.context.mailbox.item.displayReplyAllForm(htmlResponse);
                        break;
                    case 'new':
                        Office.context.mailbox.item.displayNewMessageForm({
                            htmlBody: htmlResponse,
                            subject: 'Re: ' + Office.context.mailbox.item.subject
                        });
                        break;
                }
            }, 1500); // Wait 1.5 seconds before opening email window
        }
    } catch (error) {
        console.error('[handleCustomAction] Error:', error);
        showLoading(false);
        displayError(error.message);
    }
}
// Process with Claude
async function processWithClaude(content, prompt) {
    console.log('[processWithClaude] Starting API call');
    try {
        const timestamp = new Date().getTime();
        const toneSelect = document.getElementById('tone-select');
        const tone = toneSelect ? toneSelect.value : 'professional';
        
        console.log('[processWithClaude] Using tone:', tone);
        
        const response = await fetch(`https://autopen-e3e2eyezbsdsg5ar.centralus-01.azurewebsites.net/api/process-email?t=${timestamp}`, {
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

// Format email for reply
function formatEmailForReply(text) {
    return text
        .replace(/\n{3,}/g, '\n\n')
        .trim();
}

// UI Helper Functions
function showLoading(show) {
    console.log('[showLoading] Setting loading:', show);
    const loadingElement = document.getElementById('loading');
    const resultContainer = document.getElementById('result-container');
    const resultElement = document.getElementById('result');
    
    if (!loadingElement || !resultContainer) {
        console.error('[showLoading] Required elements not found');
        return;
    }

    // Always ensure the result container is visible
    resultContainer.style.display = 'block';
    
    if (show) {
        // When showing loading, clear the result
        if (resultElement) {
            resultElement.innerHTML = '';
        }
        loadingElement.style.display = 'flex';
    } else {
        loadingElement.style.display = 'none';
    }
}

function displayResult(message) {
    console.log('[displayResult] Displaying result');
    const resultContainer = document.getElementById('result-container');
    const resultElement = document.getElementById('result');
    const loadingElement = document.getElementById('loading');
    
    if (!resultContainer || !resultElement) {
        console.error('[displayResult] Required elements not found');
        return;
    }

    // Hide loading if it's still showing
    if (loadingElement) {
        loadingElement.style.display = 'none';
    }

    // Make sure container is visible and show the message
    resultContainer.style.display = 'block';
    
    // Format and display the message
    const formattedMessage = message.replace(/\n/g, '<br>');
    resultElement.innerHTML = formattedMessage;
    resultElement.style.whiteSpace = 'pre-wrap';
    resultElement.style.lineHeight = '1.5';
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