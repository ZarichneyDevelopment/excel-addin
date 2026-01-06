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
    if (errorDetails) {
        navigator.clipboard.writeText(errorDetails.textContent || '')
            .then(() => {
                const copyBtn = document.getElementById('error-copy-btn');
                if (copyBtn) {
                    const originalText = copyBtn.textContent;
                    copyBtn.textContent = 'Copied!';
                    setTimeout(() => copyBtn.textContent = originalText, 2000);
                }
            })
            .catch(err => {
                console.error('Failed to copy error: ', err);
            });
    }
}

export function closeErrorConsole() {
    const errorContainer = document.getElementById('error-container');
    if (errorContainer) {
        errorContainer.style.display = 'none';
    }
}
