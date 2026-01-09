type LogType = 'info' | 'success' | 'warn' | 'error';

const rawConsole = {
    log: console.log.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
};

function safeGetConsoleOutput(): HTMLElement | null {
    if (typeof document === 'undefined') return null;
    return document.getElementById('console-output');
}

function alsoLogToBrowserConsole(message: string, type: LogType) {
    const prefix = `[Budget Helper] ${message}`;
    switch (type) {
        case 'error':
            rawConsole.error(prefix);
            break;
        case 'warn':
            rawConsole.warn(prefix);
            break;
        default:
            rawConsole.log(prefix);
            break;
    }
}

function appendToTaskpaneConsole(message: string, type: LogType) {
    const consoleOutput = safeGetConsoleOutput();
    if (!consoleOutput) return;
    if (typeof (consoleOutput as any).appendChild !== 'function') return;

    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = `> ${message}`;
    consoleOutput.appendChild(entry);
    consoleOutput.scrollTop = consoleOutput.scrollHeight;
}

export function logToTaskpane(message: string, type: LogType = 'info') {
    appendToTaskpaneConsole(message, type);
}

export function logToConsole(message: string, type: LogType = 'info') {
    alsoLogToBrowserConsole(message, type);
    appendToTaskpaneConsole(message, type);
}

export function clearConsole() {
    const consoleOutput = safeGetConsoleOutput();
    if (!consoleOutput) return;

    consoleOutput.innerHTML = '';
    logToConsole('Console cleared.');
}
