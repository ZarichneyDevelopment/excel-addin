/** @jest-environment jsdom */
import { handleFileDrop, ProcessFileDrop, preventDefaults, handleDragOver } from './file-drop';
import { WriteToTable } from './excel-helpers';
import { ProcessTransactions } from './transaction';

// Mock logger to avoid DOM interaction issues if not fully supported by jsdom default
jest.mock('./logger', () => ({
    logToConsole: jest.fn(),
    clearConsole: jest.fn()
}));

jest.mock('./excel-helpers', () => ({
  WriteToTable: jest.fn(),
}));

jest.mock('./transaction', () => ({
  ProcessTransactions: jest.fn().mockResolvedValue([]),
}));

describe('File drop helpers', () => {
  it('prevents default events', () => {
    const event = { preventDefault: jest.fn(), stopPropagation: jest.fn() };
    preventDefaults(event);
    expect(event.preventDefault).toHaveBeenCalled();
    expect(event.stopPropagation).toHaveBeenCalled();
  });

  it('handles drag over', () => {
    const event = {
      preventDefault: jest.fn(),
      stopPropagation: jest.fn(),
      dataTransfer: { dropEffect: '' },
    };
    handleDragOver(event);
    expect(event.dataTransfer.dropEffect).toBe('copy');
  });

  it('initializes a FileReader and processes dropped files', () => {
    const fakeFile = new Blob(['csv-contents'], { type: 'text/csv' });
    const event = {
      dataTransfer: { files: [fakeFile] },
    };

    const reader = {
      readAsText: jest.fn(),
      onload: null,
    };
    (global as any).FileReader = jest.fn(() => reader);

    handleFileDrop(event);

    expect((global as any).FileReader).toHaveBeenCalled();
    expect(reader.readAsText).toHaveBeenCalledWith(fakeFile);
    
    // Simulate onload
    reader.onload({ target: { result: 'csv-contents' } });
  });

  it('processes file contents and writes transactions', async () => {
    const event = { target: { result: 'csv-contents' } };
    const mockTransactions = [{ id: '123' }];
    (ProcessTransactions as jest.Mock).mockResolvedValue(mockTransactions);

    await ProcessFileDrop(event);

    expect(ProcessTransactions).toHaveBeenCalledWith('csv-contents');
    expect(WriteToTable).toHaveBeenCalledWith('Transactions', mockTransactions);
  });
});