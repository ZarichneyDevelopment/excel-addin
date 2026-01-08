export function handleError(error: unknown, context: string) {
    const errorContainer = document.getElementById('error-container');
    const errorContent = document.getElementById('error-content');
    const errorDetails = document.getElementById('error-details');

    if (!errorContainer || !errorContent || !errorDetails) {
        console.error('Error UI elements not found. Logging to console only.');
        console.error(`[${context}]`, error);
        return;
    }

    const timestamp = new Date().toISOString();
    const errorMessage = error instanceof Error ? error.message : String(error);
    const stackTrace = error instanceof Error ? error.stack : 'No stack trace available';

    const formattedError = `--- ERROR REPORT ---
Context: ${context}
Time: ${timestamp}
Message: ${errorMessage}

Stack Trace:
${stackTrace}
--------------------`;

    // Display the error
    errorDetails.textContent = formattedError;
    errorContainer.style.display = 'block';
    
    // Log to console as well
    console.error(`[${context}]`, error);
}

export function copyErrorToClipboard() {
    const errorDetails = document.getElementById('error-details');
    if (!errorDetails) return;

    const text = errorDetails.textContent || '';

    // Try modern API
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text)
            .then(showCopySuccess)
            .catch(err => {
                console.warn('Clipboard API failed, trying fallback', err);
                fallbackCopyText(text);
            });
    } else {
        fallbackCopyText(text);
    }
}

function fallbackCopyText(text: string) {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    
    // Ensure it's part of the DOM but invisible to user, yet "visible" to browser focus
    textArea.style.position = "fixed";
    textArea.style.top = "0";
    textArea.style.left = "0";
    textArea.style.opacity = "0";
    textArea.style.zIndex = "-1";
    document.body.appendChild(textArea);
    
    textArea.focus();
    textArea.select();
    
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            console.log('Fallback copy successful');
            showCopySuccess();
        } else {
            console.error('Fallback copy failed (execCommand returned false).');
        }
    } catch (err) {
        console.error('Fallback copy failed with error:', err);
    }
    
    document.body.removeChild(textArea);
}

function showCopySuccess() {
    const copyBtn = document.getElementById('error-copy-btn');
    if (copyBtn) {
        const originalText = copyBtn.textContent;
        copyBtn.textContent = 'Copied!';
        setTimeout(() => copyBtn.textContent = originalText, 2000);
    }
}

export function closeErrorConsole() {
    const errorContainer = document.getElementById('error-container');
    if (errorContainer) {
        errorContainer.style.display = 'none';
    }
}
