import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import './ExcelGrid.css';

interface Cell {
  value: string;
  formula?: string;
  style?: {
    backgroundColor?: string;
    color?: string;
    fontWeight?: string;
    textAlign?: 'left' | 'center' | 'right';
    border?: string;
  };
}

interface ExcelGridProps {
  data: string[][];
  onCellChange?: (row: number, col: number, value: string) => void;
  onSelectionChange?: (selection: CellSelection) => void;
  formatData?: Record<string, any>; // Format data from Excel
}

interface CellSelection {
  start: { row: number; col: number };
  end: { row: number; col: number };
}

interface ColumnWidth {
  [col: number]: number;
}

const ExcelGrid: React.FC<ExcelGridProps> = ({ data, onCellChange, onSelectionChange, formatData = {} }) => {
  // Configuration
  const CELL_HEIGHT = 25;
  const DEFAULT_CELL_WIDTH = 100;
  const MIN_CELL_WIDTH = 50;
  const ROW_HEADER_WIDTH = 60;
  const VISIBLE_BUFFER = 5; // Extra rows/cols to render outside viewport

  // États
  const [columnWidths, setColumnWidths] = useState<ColumnWidth>({});
  const [selection, setSelection] = useState<CellSelection | null>(null);
  const [isSelecting, setIsSelecting] = useState(false);
  const [editingCell, setEditingCell] = useState<{ row: number; col: number } | null>(null);
  const [editValue, setEditValue] = useState('');
  const [scrollPosition, setScrollPosition] = useState({ x: 0, y: 0 });
  const [resizingColumn, setResizingColumn] = useState<number | null>(null);
  const [resizeStartX, setResizeStartX] = useState(0);
  const [resizeStartWidth, setResizeStartWidth] = useState(0);

  // Refs
  const gridRef = useRef<HTMLDivElement>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // Dimensions
  const totalRows = data.length;
  const totalCols = Math.max(...data.map(row => row.length), 0);

  // Calculer les dimensions totales
  const totalHeight = totalRows * CELL_HEIGHT;
  const totalWidth = useMemo(() => {
    let width = 0;
    for (let col = 0; col < totalCols; col++) {
      width += columnWidths[col] || DEFAULT_CELL_WIDTH;
    }
    return width;
  }, [columnWidths, totalCols]);

  // Calculer les cellules visibles
  const visibleRange = useMemo(() => {
    if (!scrollContainerRef.current) {
      return { startRow: 0, endRow: 50, startCol: 0, endCol: 20 };
    }

    const containerHeight = scrollContainerRef.current.clientHeight - CELL_HEIGHT; // Moins l'en-tête
    const containerWidth = scrollContainerRef.current.clientWidth - ROW_HEADER_WIDTH;

    const startRow = Math.max(0, Math.floor(scrollPosition.y / CELL_HEIGHT) - VISIBLE_BUFFER);
    const endRow = Math.min(
      totalRows,
      Math.ceil((scrollPosition.y + containerHeight) / CELL_HEIGHT) + VISIBLE_BUFFER
    );

    // Calculer les colonnes visibles
    let startCol = 0;
    let endCol = 0;
    let accumulatedWidth = 0;
    let startFound = false;

    for (let col = 0; col < totalCols; col++) {
      const colWidth = columnWidths[col] || DEFAULT_CELL_WIDTH;
      
      if (!startFound && accumulatedWidth + colWidth > scrollPosition.x) {
        startCol = Math.max(0, col - VISIBLE_BUFFER);
        startFound = true;
      }
      
      if (accumulatedWidth > scrollPosition.x + containerWidth + VISIBLE_BUFFER * DEFAULT_CELL_WIDTH) {
        endCol = col;
        break;
      }
      
      accumulatedWidth += colWidth;
    }

    if (endCol === 0) endCol = totalCols;

    return { startRow, endRow, startCol, endCol };
  }, [scrollPosition, totalRows, totalCols, columnWidths]);

  // Fonctions utilitaires
  const getColumnName = (index: number): string => {
    let name = '';
    let num = index;
    while (num >= 0) {
      name = String.fromCharCode(65 + (num % 26)) + name;
      num = Math.floor(num / 26) - 1;
    }
    return name;
  };

  const getCellRef = (row: number, col: number): string => {
    return `${getColumnName(col)}${row + 1}`;
  };

  const getColumnLeft = (col: number): number => {
    let left = 0;
    for (let i = 0; i < col; i++) {
      left += columnWidths[i] || DEFAULT_CELL_WIDTH;
    }
    return left;
  };

  // Gestion du scroll
  const handleScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
    const target = e.currentTarget;
    setScrollPosition({
      x: target.scrollLeft,
      y: target.scrollTop
    });
  }, []);

  // Gestion de la sélection
  const handleMouseDown = useCallback((row: number, col: number, e: React.MouseEvent) => {
    e.preventDefault();
    
    if (editingCell) {
      saveEdit();
    }

    const newSelection = { start: { row, col }, end: { row, col } };
    setSelection(newSelection);
    setIsSelecting(true);
    
    if (onSelectionChange) {
      onSelectionChange(newSelection);
    }
  }, [editingCell, onSelectionChange]);

  const handleMouseMove = useCallback((row: number, col: number) => {
    if (isSelecting && selection) {
      const newSelection = {
        ...selection,
        end: { row, col }
      };
      setSelection(newSelection);
      
      if (onSelectionChange) {
        onSelectionChange(newSelection);
      }
    }
  }, [isSelecting, selection, onSelectionChange]);

  const handleMouseUp = useCallback(() => {
    setIsSelecting(false);
  }, []);

  // Gestion de l'édition
  const handleDoubleClick = useCallback((row: number, col: number) => {
    setEditingCell({ row, col });
    setEditValue(data[row]?.[col] || '');
  }, [data]);

  const handleEditChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    setEditValue(e.target.value);
  }, []);

  const saveEdit = useCallback(() => {
    if (editingCell && onCellChange) {
      onCellChange(editingCell.row, editingCell.col, editValue);
    }
    setEditingCell(null);
    setEditValue('');
  }, [editingCell, editValue, onCellChange]);

  const handleEditKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      saveEdit();
    } else if (e.key === 'Escape') {
      setEditingCell(null);
      setEditValue('');
    }
  }, [saveEdit]);

  // Gestion du redimensionnement des colonnes
  const handleColumnResizeStart = useCallback((col: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setResizingColumn(col);
    setResizeStartX(e.clientX);
    setResizeStartWidth(columnWidths[col] || DEFAULT_CELL_WIDTH);
  }, [columnWidths]);

  const handleColumnResize = useCallback((e: MouseEvent) => {
    if (resizingColumn !== null) {
      const diff = e.clientX - resizeStartX;
      const newWidth = Math.max(MIN_CELL_WIDTH, resizeStartWidth + diff);
      setColumnWidths(prev => ({
        ...prev,
        [resizingColumn]: newWidth
      }));
    }
  }, [resizingColumn, resizeStartX, resizeStartWidth]);

  const handleColumnResizeEnd = useCallback(() => {
    setResizingColumn(null);
  }, []);

  // Effets
  useEffect(() => {
    if (resizingColumn !== null) {
      document.addEventListener('mousemove', handleColumnResize);
      document.addEventListener('mouseup', handleColumnResizeEnd);
      return () => {
        document.removeEventListener('mousemove', handleColumnResize);
        document.removeEventListener('mouseup', handleColumnResizeEnd);
      };
    }
  }, [resizingColumn, handleColumnResize, handleColumnResizeEnd]);

  useEffect(() => {
    if (editingCell && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [editingCell]);

  // Vérifier si une cellule est dans la sélection
  const isCellSelected = useCallback((row: number, col: number) => {
    if (!selection) return false;
    
    const minRow = Math.min(selection.start.row, selection.end.row);
    const maxRow = Math.max(selection.start.row, selection.end.row);
    const minCol = Math.min(selection.start.col, selection.end.col);
    const maxCol = Math.max(selection.start.col, selection.end.col);
    
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  }, [selection]);

  // Rendu des cellules
  const renderCell = (row: number, col: number) => {
    const isEditing = editingCell?.row === row && editingCell?.col === col;
    const isSelected = isCellSelected(row, col);
    const value = data[row]?.[col] || '';
    const cellRef = getCellRef(row, col);
    
    // Obtenir le style de la cellule depuis formatData
    const cellFormat = formatData[cellRef] || {};
    
    return (
      <div
        key={`${row}-${col}`}
        className={`excel-cell ${isSelected ? 'selected' : ''} ${isEditing ? 'editing' : ''}`}
        style={{
          position: 'absolute',
          left: getColumnLeft(col),
          top: row * CELL_HEIGHT,
          width: columnWidths[col] || DEFAULT_CELL_WIDTH,
          height: CELL_HEIGHT,
          ...cellFormat.style
        }}
        onMouseDown={(e) => handleMouseDown(row, col, e)}
        onMouseMove={() => handleMouseMove(row, col)}
        onDoubleClick={() => handleDoubleClick(row, col)}
      >
        {isEditing ? (
          <input
            ref={inputRef}
            type="text"
            value={editValue}
            onChange={handleEditChange}
            onKeyDown={handleEditKeyDown}
            onBlur={saveEdit}
            className="cell-editor"
          />
        ) : (
          <div className="cell-content">{value}</div>
        )}
      </div>
    );
  };

  return (
    <div className="excel-grid-container" ref={gridRef} onMouseUp={handleMouseUp}>
      {/* En-tête des colonnes - Fixe */}
      <div className="column-headers">
        <div className="corner-cell" style={{ width: ROW_HEADER_WIDTH }}>
          {selection && `${getCellRef(selection.start.row, selection.start.col)}`}
        </div>
        <div className="column-headers-scroll" style={{ left: ROW_HEADER_WIDTH, transform: `translateX(-${scrollPosition.x}px)` }}>
          {Array.from({ length: totalCols }, (_, col) => (
            <div
              key={col}
              className="column-header"
              style={{
                left: getColumnLeft(col),
                width: columnWidths[col] || DEFAULT_CELL_WIDTH
              }}
            >
              <span>{getColumnName(col)}</span>
              <div
                className="column-resize-handle"
                onMouseDown={(e) => handleColumnResizeStart(col, e)}
              />
            </div>
          ))}
        </div>
      </div>

      {/* Conteneur principal avec scroll */}
      <div className="excel-scroll-wrapper">
        {/* En-têtes des lignes - Fixe */}
        <div className="row-headers" style={{ transform: `translateY(-${scrollPosition.y}px)` }}>
          {Array.from({ length: totalRows }, (_, row) => (
            <div
              key={row}
              className="row-header"
              style={{ top: row * CELL_HEIGHT }}
            >
              {row + 1}
            </div>
          ))}
        </div>

        {/* Zone de scroll avec les cellules */}
        <div
          ref={scrollContainerRef}
          className="excel-scroll-container"
          onScroll={handleScroll}
          style={{ left: ROW_HEADER_WIDTH, top: CELL_HEIGHT }}
        >
          {/* Spacer pour créer la zone de scroll */}
          <div style={{ width: totalWidth, height: totalHeight, position: 'relative' }}>
            {/* Cellules visibles uniquement */}
            {Array.from(
              { length: visibleRange.endRow - visibleRange.startRow },
              (_, i) => visibleRange.startRow + i
            ).map(row =>
              Array.from(
                { length: visibleRange.endCol - visibleRange.startCol },
                (_, j) => visibleRange.startCol + j
              ).map(col => renderCell(row, col))
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default ExcelGrid;