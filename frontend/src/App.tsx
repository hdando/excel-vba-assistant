import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Upload, FileSpreadsheet, MessageCircle, Download, Loader2, AlertCircle, Check, X, ChevronRight, Code, Table, Search, Zap, FileText, Settings, RefreshCw } from 'lucide-react';

// Types
interface ExcelStructure {
  sheets: Array<{
    name: string;
    max_row: number;
    max_column: number;
    has_data: boolean;
    formulas: Array<{ cell: string; formula: string }>;
    data?: string[][];
    headers?: string[];
  }>;
  total_sheets: number;
  has_vba: boolean;
}

interface ChatMessage {
  id: string;
  type: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
}

interface Session {
  session_id: string;
  filename: string;
  structure: ExcelStructure;
  vba_modules: string[];
  initial_analysis: string;
}

// API Configuration
const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:8000';

// Components
const UploadZone: React.FC<{ onUpload: (file: File) => void }> = ({ onUpload }) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      onUpload(files[0]);
    }
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onUpload(e.target.files[0]);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 p-4">
      <div className="max-w-2xl w-full">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-gray-900 mb-4">Excel VBA Assistant</h1>
          <p className="text-xl text-gray-600">
            Analysez et modifiez vos fichiers Excel avec l'aide d'une IA conversationnelle
          </p>
        </div>
        
        <div
          className={`border-2 border-dashed rounded-lg p-12 text-center transition-colors cursor-pointer
            ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-gray-400'}`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
        >
          <FileSpreadsheet className="mx-auto h-16 w-16 text-gray-400 mb-4" />
          <p className="text-lg font-medium text-gray-900 mb-2">
            Glissez votre fichier Excel ici
          </p>
          <p className="text-sm text-gray-500 mb-4">ou cliquez pour sélectionner</p>
          <p className="text-xs text-gray-400">Formats supportés: .xlsx, .xlsm (max 50MB)</p>
          
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xlsm"
            onChange={handleFileSelect}
            className="hidden"
          />
        </div>
      </div>
    </div>
  );
};

const ExcelViewer: React.FC<{ structure: ExcelStructure; activeSheet: string; onSheetChange: (sheet: string) => void }> = ({ structure, activeSheet, onSheetChange }) => {
  const currentSheet = structure.sheets.find(s => s.name === activeSheet);
  const [editingCell, setEditingCell] = React.useState<{row: number, col: number} | null>(null);
  const [cellValue, setCellValue] = React.useState('');
  
  // Fonction pour générer les noms de colonnes Excel (A, B, ... Z, AA, AB, etc.)
  const getColumnName = (index: number): string => {
    let name = '';
    let num = index;
    while (num >= 0) {
      name = String.fromCharCode(65 + (num % 26)) + name;
      num = Math.floor(num / 26) - 1;
    }
    return name;
  };
  
  const handleCellClick = (rowIndex: number, colIndex: number, value: string) => {
    setEditingCell({ row: rowIndex, col: colIndex });
    setCellValue(value);
  };
  
  const handleCellChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setCellValue(e.target.value);
  };
  
  const handleCellBlur = () => {
    // Ici on pourrait sauvegarder la valeur
    setEditingCell(null);
  };
  
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      setEditingCell(null);
    }
    if (e.key === 'Escape') {
      setEditingCell(null);
    }
  };
  
  return (
    <div className="h-full flex flex-col bg-white">
      {/* Sheet Tabs */}
      <div className="flex border-b overflow-x-auto bg-gray-50">
        {structure.sheets.map((sheet) => (
          <button
            key={sheet.name}
            onClick={() => onSheetChange(sheet.name)}
            className={`px-4 py-2 text-sm font-medium whitespace-nowrap transition-colors
              ${activeSheet === sheet.name 
                ? 'text-blue-600 border-b-2 border-blue-600 bg-white' 
                : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'}`}
          >
            <Table className="inline w-4 h-4 mr-1" />
            {sheet.name}
          </button>
        ))}
      </div>
      
      {/* Sheet Content - Container avec scroll horizontal et vertical */}
      <div className="flex-1 overflow-x-auto overflow-y-auto" style={{ 
		overflowX: 'scroll', 
		overflowY: 'scroll',
		width: '100%'
	  }}>
        {currentSheet && currentSheet.data && currentSheet.data.length > 0 ? (
          <table className="border-collapse" style={{ 
		    width: `${(currentSheet.data[0]?.length || 0) * 50 + 10}px`,
		    minWidth: `${(currentSheet.data[0]?.length || 0) * 50 + 10}px`
		  }}>
            <thead className="sticky top-0 z-10">
              <tr>
                <th className="border border-gray-300 px-3 py-2 text-xs font-medium text-gray-700 bg-gray-200" style={{ width: '60px' }}></th>
                {currentSheet.data[0].map((_, index) => (
                  <th key={index} className="border border-gray-300 px-2 py-2 text-sm font-medium text-gray-700 bg-gray-100" style={{ minWidth: '120px' }}>
                    {getColumnName(index)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {currentSheet.data.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  <td className="border border-gray-300 px-3 py-1 text-xs font-medium text-gray-700 bg-gray-100" style={{ width: '60px' }}>
                    {rowIndex + 1}
                  </td>
                  {row.map((cell, colIndex) => (
                    <td 
                      key={colIndex} 
                      className="border border-gray-300 text-sm relative hover:bg-gray-50"
                      onClick={() => handleCellClick(rowIndex, colIndex, cell)}
                      style={{ padding: 0, minWidth: '120px', height: '32px', cursor: 'cell' }}
                    >
                      {editingCell?.row === rowIndex && editingCell?.col === colIndex ? (
                        <input
                          type="text"
                          value={cellValue}
                          onChange={handleCellChange}
                          onBlur={handleCellBlur}
                          onKeyDown={handleKeyDown}
                          className="w-full h-full px-2 py-1 border-2 border-blue-500 outline-none"
                          autoFocus
                        />
                      ) : (
                        <div className="px-2 py-1 h-full flex items-center">
                          {cell}
                        </div>
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div className="p-8 text-center text-gray-500">
            Aucune donnée à afficher
          </div>
        )}
      </div>
      
      {/* Indicateur de cellule active */}
      {editingCell && (
        <div className="absolute bottom-2 left-2 bg-blue-100 px-2 py-1 rounded text-sm">
          {getColumnName(editingCell.col)}{editingCell.row + 1}
        </div>
      )}
    </div>
  );
};

const VBAEditor: React.FC<{ modules: string[]; activeModule: string; onModuleChange: (module: string) => void }> = ({ modules, activeModule, onModuleChange }) => {
  return (
    <div className="h-full flex flex-col bg-gray-900 text-gray-100">
      {/* Module Tabs */}
      <div className="flex border-b border-gray-700 overflow-x-auto bg-gray-800">
        {modules.map((module) => (
          <button
            key={module}
            onClick={() => onModuleChange(module)}
            className={`px-4 py-2 text-sm font-medium whitespace-nowrap transition-colors
              ${activeModule === module 
                ? 'text-blue-400 border-b-2 border-blue-400 bg-gray-900' 
                : 'text-gray-400 hover:text-gray-200 hover:bg-gray-700'}`}
          >
            <Code className="inline w-4 h-4 mr-1" />
            {module}
          </button>
        ))}
      </div>
      
      {/* Code Editor */}
      <div className="flex-1 overflow-auto p-4">
        <div className="font-mono text-sm">
          <div className="text-gray-500 mb-4">
            ' Module: {activeModule}
          </div>
          <pre className="text-green-400">
            {`Sub Example()
    ' Votre code VBA sera affiché ici
    MsgBox "Hello from ${activeModule}"
End Sub`}
          </pre>
        </div>
      </div>
    </div>
  );
};

const ChatInterface: React.FC<{ 
  messages: ChatMessage[]; 
  onSendMessage: (message: string) => void;
  isLoading: boolean;
}> = ({ messages, onSendMessage, isLoading }) => {
  const [input, setInput] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };
  
  useEffect(() => {
    scrollToBottom();
  }, [messages]);
  
  const handleSubmit = (e: any) => {
    e.preventDefault();
    if (input.trim() && !isLoading) {
      onSendMessage(input);
      setInput('');
    }
  };
  
  const quickActions = [
    { icon: Search, label: "Analyser tout", action: "Analyse complète du fichier" },
    { icon: Zap, label: "Optimiser", action: "Optimise les performances de mon fichier" },
    { icon: FileText, label: "Documentation", action: "Génère la documentation" },
  ];
  
  return (
    <div className="h-full flex flex-col bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b px-4 py-3">
        <h2 className="text-lg font-semibold flex items-center">
          <MessageCircle className="w-5 h-5 mr-2" />
          Assistant Excel VBA
        </h2>
      </div>
      
      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {messages.map((message) => (
          <div
            key={message.id}
            className={`flex ${message.type === 'user' ? 'justify-end' : 'justify-start'}`}
          >
            <div
              className={`max-w-[80%] rounded-lg px-4 py-3 ${
                message.type === 'user'
                  ? 'bg-blue-600 text-white'
                  : message.type === 'assistant'
                  ? 'bg-white border'
                  : 'bg-yellow-50 text-yellow-800 border border-yellow-200'
              }`}
            >
              <p className="text-sm whitespace-pre-wrap">{message.content}</p>
            </div>
          </div>
        ))}
        {isLoading && (
          <div className="flex justify-start">
            <div className="bg-white border rounded-lg px-4 py-3">
              <Loader2 className="w-4 h-4 animate-spin" />
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>
      
      {/* Quick Actions */}
      <div className="px-4 py-2 bg-white border-t">
        <div className="flex gap-2 mb-2">
          {quickActions.map((action) => (
            <button
              key={action.label}
              onClick={() => onSendMessage(action.action)}
              className="flex items-center gap-1 px-3 py-1 text-xs bg-gray-100 hover:bg-gray-200 rounded-full transition-colors"
            >
              <action.icon className="w-3 h-3" />
              {action.label}
            </button>
          ))}
        </div>
      </div>
      
      {/* Input */}
      <div className="p-4 bg-white border-t">
        <div className="flex gap-2">
          <input
            type="text"
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyPress={(e) => {
              if (e.key === 'Enter' && !isLoading && input.trim()) {
                handleSubmit(e);
              }
            }}
            placeholder="Posez votre question..."
            className="flex-1 px-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
            disabled={isLoading}
          />
          <button
            onClick={handleSubmit}
            disabled={isLoading || !input.trim()}
            className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            <ChevronRight className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
};

// Main App Component
export default function App() {
  const [session, setSession] = useState<Session | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [activeSheet, setActiveSheet] = useState('');
  const [activeModule, setActiveModule] = useState('');
  const [viewMode, setViewMode] = useState<'excel' | 'vba'>('excel');
  
  const handleFileUpload = async (file: File) => {
    setIsUploading(true);
    setError(null);
    
    const formData = new FormData();
    formData.append('file', file);
    
    try {
      const response = await fetch(`${API_URL}/api/upload`, {
        method: 'POST',
        body: formData,
      });
      
      if (!response.ok) {
        throw new Error('Upload failed');
      }
      
      const data = await response.json();
      setSession(data);
      setActiveSheet(data.structure.sheets[0]?.name || '');
      setActiveModule(data.vba_modules[0] || '');
      
      // Add initial analysis as system message
      setMessages([{
        id: Date.now().toString(),
        type: 'system',
        content: data.initial_analysis,
        timestamp: new Date(),
      }]);
    } catch (err) {
      setError('Erreur lors de l\'upload du fichier');
      console.error(err);
    } finally {
      setIsUploading(false);
    }
  };
  
  const handleSendMessage = async (message: string) => {
    if (!session) return;
    
    // Add user message
    const userMessage: ChatMessage = {
      id: Date.now().toString(),
      type: 'user',
      content: message,
      timestamp: new Date(),
    };
    setMessages(prev => [...prev, userMessage]);
    
    setIsLoading(true);
    
    try {
      const response = await fetch(`${API_URL}/api/chat`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          session_id: session.session_id,
          message: message,
        }),
      });
      
      if (!response.ok) {
        throw new Error('Chat request failed');
      }
      
      // Handle streaming response
      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let assistantMessage = '';
      
      if (reader) {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          
          const chunk = decoder.decode(value);
          const lines = chunk.split('\n');
          
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              try {
                const data = JSON.parse(line.slice(6));
                if (data.chunk) {
                  assistantMessage += data.chunk;
                }
              } catch (e) {
                // Ignore parsing errors
              }
            }
          }
        }
      }
      
      // Add assistant message
      setMessages(prev => [...prev, {
        id: (Date.now() + 1).toString(),
        type: 'assistant',
        content: assistantMessage,
        timestamp: new Date(),
      }]);
    } catch (err) {
      console.error(err);
      setMessages(prev => [...prev, {
        id: (Date.now() + 1).toString(),
        type: 'system',
        content: 'Erreur lors de la communication avec l\'assistant',
        timestamp: new Date(),
      }]);
    } finally {
      setIsLoading(false);
    }
  };
  
  const handleExport = async () => {
    if (!session) return;
    
    try {
      const response = await fetch(`${API_URL}/api/export?session_id=${session.session_id}`, {
        method: 'POST',
      });
      
      if (!response.ok) {
        throw new Error('Export failed');
      }
      
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `modified_${session.filename}`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      console.error(err);
      setError('Erreur lors de l\'export du fichier');
    }
  };
  
  if (!session) {
    return (
      <>
        <UploadZone onUpload={handleFileUpload} />
        {isUploading && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center">
            <div className="bg-white rounded-lg p-6 flex items-center space-x-4">
              <Loader2 className="w-6 h-6 animate-spin" />
              <span>Analyse du fichier en cours...</span>
            </div>
          </div>
        )}
        {error && (
          <div className="fixed bottom-4 left-4 right-4 bg-red-50 border border-red-200 rounded-lg p-4 flex items-center">
            <AlertCircle className="w-5 h-5 text-red-600 mr-2" />
            <span className="text-red-800">{error}</span>
          </div>
        )}
      </>
    );
  }
  
  return (
    <div className="h-screen flex flex-col bg-gray-100">
      {/* Header */}
      <header className="bg-white border-b px-4 py-3 flex items-center justify-between">
        <div className="flex items-center space-x-4">
          <FileSpreadsheet className="w-6 h-6 text-blue-600" />
          <div>
            <h1 className="font-semibold">{session.filename}</h1>
            <p className="text-sm text-gray-600">
              {session.structure.total_sheets} feuilles
              {session.structure.has_vba && ' • Contient du VBA'}
            </p>
          </div>
        </div>
        
        <div className="flex items-center space-x-2">
          <button
            onClick={handleExport}
            className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
          >
            <Download className="w-4 h-4" />
            Exporter
          </button>
          <button
            onClick={() => {
              if (window.confirm('Êtes-vous sûr de vouloir fermer cette session ?')) {
                setSession(null);
                setMessages([]);
              }
            }}
            className="p-2 text-gray-600 hover:text-gray-900"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
      </header>
      
      {/* Main Content */}
      <div className="flex-1 flex overflow-hidden">
        {/* Left Panel - Excel/VBA Viewer (75%) */}
        <div className="flex-1 flex flex-col">
          {/* View Mode Selector */}
          {session.structure.has_vba && (
            <div className="bg-white border-b px-4 py-2 flex space-x-2">
              <button
                onClick={() => setViewMode('excel')}
                className={`px-3 py-1 rounded transition-colors ${
                  viewMode === 'excel' 
                    ? 'bg-blue-100 text-blue-700' 
                    : 'text-gray-600 hover:bg-gray-100'
                }`}
              >
                <Table className="inline w-4 h-4 mr-1" />
                Données Excel
              </button>
              <button
                onClick={() => setViewMode('vba')}
                className={`px-3 py-1 rounded transition-colors ${
                  viewMode === 'vba' 
                    ? 'bg-blue-100 text-blue-700' 
                    : 'text-gray-600 hover:bg-gray-100'
                }`}
              >
                <Code className="inline w-4 h-4 mr-1" />
                Code VBA
              </button>
            </div>
          )}
          
          {/* Content Area */}
          <div className="flex-1 overflow-hidden">
            {viewMode === 'excel' ? (
              <ExcelViewer 
                structure={session.structure} 
                activeSheet={activeSheet}
                onSheetChange={setActiveSheet}
              />
            ) : (
              <VBAEditor
                modules={session.vba_modules}
                activeModule={activeModule}
                onModuleChange={setActiveModule}
              />
            )}
          </div>
        </div>
        
        {/* Right Panel - Chat (25%) */}
        <div className="w-96 border-l">
          <ChatInterface
            messages={messages}
            onSendMessage={handleSendMessage}
            isLoading={isLoading}
          />
        </div>
      </div>
    </div>
  );
}