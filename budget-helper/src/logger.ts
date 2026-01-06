// Logger Helper
export function logToConsole(message: string, type: 'info' | 'success' | 'warn' | 'error' = 'info') {
    const consoleOutput = document.getElementById('console-output');
    if (consoleOutput) {
        const entry = document.createElement('div');
        entry.className = `log-entry ${type}`;
        entry.textContent = `> ${message}`;
        consoleOutput.appendChild(entry);
        consoleOutput.scrollTop = consoleOutput.scrollHeight;
    }
}

export function clearConsole() {
    const consoleOutput = document.getElementById('console-output');
    if (consoleOutput) {
        consoleOutput.innerHTML = '';
        logToConsole('Console cleared.');
    }
}
