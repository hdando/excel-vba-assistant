import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Upload, FileSpreadsheet, MessageCircle, Download, Loader2, AlertCircle, Check, X, ChevronRight, Code, Table, Search, Zap, FileText, Settings, RefreshCw, Clock } from 'lucide-react';
import ExcelGrid from './ExcelGrid';
import './ExcelGrid.css';

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
  vba_code?: { [key: string]: string };
  initial_analysis: string;
}

// API Configuration - MODIFIER CETTE URL POUR LA PRODUCTION
const API_URL = process.env.REACT_APP_API_URL || 'http://localhost:8000';

// Configuration des timeouts
const SESSION_TIMEOUT = 36000; // 1 heure en secondes
const WARNING_TIME = 300; // Avertir 5 minutes avant expiration

// Components
const SessionTimeoutWarning: React.FC<{ 
  timeLeft: number; 
  onExtend: () => void; 
  onClose: () => void 
}> = ({ timeLeft, onExtend, onClose }) => {
  const minutes = Math.floor(timeLeft / 60);
  const seconds = timeLeft % 60;

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg p-6 max-w-md mx-4">
        <div className="flex items-center mb-4">
          <Clock className="w-6 h-6 text-orange-500 mr-2" />
          <h3 className="text-lg font-semibold">Session bientôt expirée</h3>
        </div>
        <p className="text-gray-600 mb-4">
          Votre session va expirer dans {minutes}:{seconds.toString().padStart(2, '0')}. 
          Tous vos changements non exportés seront perdus.
        </p>
        <div className="flex space-x-3">
          <button
            onClick={onExtend}
            className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
          >
            Continuer à travailler
          </button>
          <button
            onClick={onClose}
            className="px-4 py-2 border rounded-lg hover:bg-gray-50"
          >
            Ignorer
          </button>
        </div>
      </div>
    </div>
  );
};

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
          <div className="mt-4 text-sm text-gray-500">
            💡 Session temporaire - Pensez à exporter vos modifications avant de quitter
          </div>
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

const VBAEditor: React.FC<{ 
  modules: string[]; 
  activeModule: string; 
  onModuleChange: (module: string) => void;
  vbaCode?: { [key: string]: string };
}> = ({ modules = [], activeModule, onModuleChange, vbaCode = {} }) => {
  if (!modules || modules.length === 0) {
    return (
      <div className="h-full flex flex-col bg-gray-900 text-gray-100 items-center justify-center">
        <Code className="w-16 h-16 text-gray-600 mb-4" />
        <p className="text-gray-400">Aucun module VBA trouvé dans ce fichier</p>
        <p className="text-sm text-gray-500 mt-2">Ce fichier ne contient pas de code VBA</p>
      </div>
    );
  }
  
  const currentCode = vbaCode[activeModule] || `' Code du module ${activeModule} non disponible`;
  
  const highlightVBA = (code: string) => {
    const escapeHtml = (str: string) => {
      return str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    };
    
    let escaped = escapeHtml(code);
    
    const keywords = /\b(Sub|End Sub|Function|End Function|Dim|As|Set|If|Then|Else|ElseIf|End If|For|To|Step|Next|Do|While|Until|Loop|With|End With|Private|Public|Option Explicit|Option Base|ByVal|ByRef|Const|True|False|Nothing|Null|Empty|And|Or|Not|Is|Like|Mod|New|ReDim|Preserve|Select Case|Case|Case Else|End Select|Exit|Exit Sub|Exit Function|Exit For|Exit Do|GoTo|On Error|On Error Resume Next|On Error GoTo|Resume|Resume Next|Call|Let|Get|Property|End Property|Type|End Type|Enum|End Enum|Declare|Static|Friend|Global|Implements|Inherits|Interface|Module|Namespace|Imports|Class|End Class|Variant|Integer|Long|Single|Double|Currency|String|Boolean|Date|Object|Byte|Array|Collection|Dictionary|Worksheet|Workbook|Range|Cells|ActiveSheet|ActiveWorkbook|Application|MsgBox|InputBox|Debug|Print|Error|Err|Raise|Clear|Description|Number|Source|Each|In|Me|ThisWorkbook|Sheets|Worksheets|Charts|UserForm|Controls|Value|Formula|Text|Count|Rows|Columns|Offset|Resize|End|xlUp|xlDown|xlToLeft|xlToRight|CurrentRegion|UsedRange|SpecialCells|Visible|Hidden|Name|Names|Address|Row|Column)\b/gi;
    const comments = /(&#039;.*$)/gm;
    const strings = /(&quot;[^&]*&quot;)/g;
    const numbers = /\b(\d+\.?\d*)\b/g;
    
    let highlighted = escaped
      .replace(strings, '<span style="color: #ce9178;">$1</span>')
      .replace(comments, '<span style="color: #608b4e; font-style: italic;">$1</span>')
      .replace(keywords, '<span style="color: #569cd6; font-weight: bold;">$1</span>')
      .replace(numbers, '<span style="color: #b5cea8;">$1</span>');
    
    return highlighted;
  };
  
  return (
    <div className="h-full flex flex-col bg-gray-900 text-gray-100">
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
      
      <div className="flex-1 overflow-auto">
        <div className="font-mono text-sm flex">
          <div className="bg-gray-800 text-gray-500 select-none border-r border-gray-700 px-2 py-4" style={{ minWidth: '3rem' }}>
            {currentCode.split('\n').map((_, index) => (
              <div key={index} className="text-right pr-2" style={{ height: '1.5em' }}>
                {index + 1}
              </div>
            ))}
          </div>
          <div className="flex-1 p-4 overflow-x-auto">
            <pre 
              className="text-gray-300"
              style={{ lineHeight: '1.5em' }}
              dangerouslySetInnerHTML={{ __html: highlightVBA(currentCode) }}
            />
          </div>
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
      <div className="bg-white border-b px-4 py-3">
        <h2 className="text-lg font-semibold flex items-center">
          <MessageCircle className="w-5 h-5 mr-2" />
          Assistant Excel VBA
        </h2>
      </div>
      
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
  const [pendingUpdates, setPendingUpdates] = useState<Set<string>>(new Set());
  
  // Gestion du timeout de session
  const [sessionStartTime, setSessionStartTime] = useState<Date | null>(null);
  const [showTimeoutWarning, setShowTimeoutWarning] = useState(false);
  const [timeLeft, setTimeLeft] = useState(0);
  
  // Timer pour le timeout de session
  useEffect(() => {
    if (!sessionStartTime) return;

    const interval = setInterval(() => {
      const now = new Date();
      const elapsed = Math.floor((now.getTime() - sessionStartTime.getTime()) / 1000);
      const remaining = SESSION_TIMEOUT - elapsed;
      
      if (remaining <= 0) {
        // Session expirée
        setSession(null);
        setMessages([]);
        setSessionStartTime(null);
        setShowTimeoutWarning(false);
        setError('Session expirée. Veuillez recharger un fichier.');
      } else if (remaining <= WARNING_TIME && !showTimeoutWarning) {
        // Montrer l'avertissement
        setShowTimeoutWarning(true);
        setTimeLeft(remaining);
      } else if (showTimeoutWarning) {
        setTimeLeft(remaining);
      }
    }, 1000);

    return () => clearInterval(interval);
  }, [sessionStartTime, showTimeoutWarning]);
  
  const handleExtendSession = () => {
    // Simuler une activité pour "étendre" la session
    setSessionStartTime(new Date());
    setShowTimeoutWarning(false);
    
    // Optionnel : faire un ping au serveur pour réinitialiser le timer côté backend
    if (session) {
      fetch(`${API_URL}/api/session/${session.session_id}`)
        .catch(() => {}); // Ignorer les erreurs, c'est juste pour réinitialiser le timer
    }
  };
  
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
      const sessionData = {
        ...data,
        vba_modules: data.vba_modules || [],
        vba_code: data.vba_code || {}
      };
      setSession(sessionData);
      setActiveSheet(sessionData.structure.sheets[0]?.name || '');
      setActiveModule(sessionData.vba_modules && sessionData.vba_modules.length > 0 ? sessionData.vba_modules[0] : '');
      
      // Démarrer le timer de session
      setSessionStartTime(new Date());
      setShowTimeoutWarning(false);
      
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
  
  const updateTimeouts = useRef<{ [key: string]: NodeJS.Timeout }>({});
  
  const handleCellChange = useCallback(async (row: number, col: number, value: string) => {
    if (!session) return;
    
    const cellKey = `${activeSheet}-${row}-${col}`;
    
    setSession(prev => {
      if (!prev) return prev;
      const newSession = { ...prev };
      const sheetIndex = newSession.structure.sheets.findIndex(s => s.name === activeSheet);
      if (sheetIndex >= 0 && newSession.structure.sheets[sheetIndex].data) {
        newSession.structure.sheets[sheetIndex].data![row][col] = value;
      }
      return newSession;
    });
    
    if (updateTimeouts.current[cellKey]) {
      clearTimeout(updateTimeouts.current[cellKey]);
    }
    
    setPendingUpdates(prev => new Set(prev).add(cellKey));
    
    updateTimeouts.current[cellKey] = setTimeout(async () => {
      try {
        const response = await fetch(`${API_URL}/api/update-cell`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            session_id: session.session_id,
            sheet_name: activeSheet,
            row,
            col,
            value,
          }),
        });
        
        if (!response.ok) {
          throw new Error('Update failed');
        }
        
        setPendingUpdates(prev => {
          const newSet = new Set(prev);
          newSet.delete(cellKey);
          return newSet;
        });
      } catch (err) {
        console.error('Error updating cell:', err);
        setError('Erreur lors de la mise à jour de la cellule');
      }
    }, 500);
  }, [session, activeSheet]);
  
  const handleSendMessage = async (message: string) => {
    if (!session) return;
    
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
      
      setMessages(prev => [...prev, {
        id: (Date.now() + 1).toString(),
        type: 'assistant',
        content: assistantMessage,
        timestamp: new Date(),
      }]);
      
      if (assistantMessage.includes('modifié') || assistantMessage.includes('écrit')) {
        const response = await fetch(`${API_URL}/api/session/${session.session_id}`);
        if (response.ok) {
          const updatedSession = await response.json();
          setSession(updatedSession);
        }
      }
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
  
  const handleSelectionChange = useCallback((selection: any) => {
    console.log('Selection changed:', selection);
  }, []);
  
  // Calculer le temps restant pour l'affichage
  const getTimeRemaining = () => {
    if (!sessionStartTime) return '';
    const now = new Date();
    const elapsed = Math.floor((now.getTime() - sessionStartTime.getTime()) / 1000);
    const remaining = SESSION_TIMEOUT - elapsed;
    const minutes = Math.floor(remaining / 60);
    const hours = Math.floor(minutes / 60);
    
    if (hours > 0) {
      return `${hours}h ${minutes % 60}m restantes`;
    } else if (minutes > 0) {
      return `${minutes}m restantes`;
    } else {
      return `${remaining}s restantes`;
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
  
  const currentSheet = session.structure.sheets.find(s => s.name === activeSheet);
  const sheetData = currentSheet?.data || [];
  
  return (
    <div className="h-screen flex flex-col bg-gray-100">
      {/* Avertissement de timeout */}
      {showTimeoutWarning && (
        <SessionTimeoutWarning
          timeLeft={timeLeft}
          onExtend={handleExtendSession}
          onClose={() => setShowTimeoutWarning(false)}
        />
      )}
      
      {/* Header */}
      <header className="bg-white border-b px-4 py-3 flex items-center justify-between">
        <div className="flex items-center space-x-4">
          <FileSpreadsheet className="w-6 h-6 text-blue-600" />
          <div>
            <h1 className="font-semibold">{session.filename}</h1>
            <p className="text-sm text-gray-600">
              {session.structure.total_sheets} feuilles
              {session.structure.has_vba && ' • Contient du VBA'}
              {pendingUpdates.size > 0 && ` • ${pendingUpdates.size} modifications en cours...`}
              {sessionStartTime && (
                <span className="ml-2 text-orange-600">
                  🕐 {getTimeRemaining()}
                </span>
              )}
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
                setSessionStartTime(null);
                setShowTimeoutWarning(false);
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
          <div className="bg-white border-b px-4 py-2 flex items-center justify-between">
            <div className="flex space-x-2">
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
              {session.structure.has_vba && (
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
              )}
            </div>
            
            {viewMode === 'excel' && (
              <div className="flex space-x-1">
                {session.structure.sheets.map((sheet) => (
                  <button
                    key={sheet.name}
                    onClick={() => setActiveSheet(sheet.name)}
                    className={`px-3 py-1 text-sm rounded transition-colors ${
                      activeSheet === sheet.name 
                        ? 'bg-gray-200 text-gray-900' 
                        : 'text-gray-600 hover:bg-gray-100'
                    }`}
                  >
                    {sheet.name}
                  </button>
                ))}
              </div>
            )}
          </div>
          
          <div className="flex-1 overflow-hidden">
            {viewMode === 'excel' ? (
              <ExcelGrid 
                data={sheetData}
                onCellChange={handleCellChange}
                onSelectionChange={handleSelectionChange}
              />
            ) : (
              <VBAEditor
                modules={session.vba_modules || []}
                activeModule={activeModule}
                onModuleChange={setActiveModule}
                vbaCode={session.vba_code || {}}
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