import React, { useState } from 'react';
import { X, Plus, Trash2, FileSpreadsheet, Clipboard, AlertTriangle, Upload } from 'lucide-react';
import { Purchase } from '../App';
import * as XLSX from 'xlsx';

interface BulkAddItemsModalProps {
  show: boolean;
  onClose: () => void;
  onSave: (items: Partial<Purchase>[]) => void;
}

export const BulkAddItemsModal: React.FC<BulkAddItemsModalProps> = ({ show, onClose, onSave }) => {
  const [items, setItems] = useState<Partial<Purchase>[]>([
    { name: '', priceYuan: 0, quantity: 1 }
  ]);
  const [textInput, setTextInput] = useState('');
  const [previewData, setPreviewData] = useState<any[][]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [startRow, setStartRow] = useState<number>(1);
  const [columnMapping, setColumnMapping] = useState<Record<string, number | null>>(() => {
    const saved = localStorage.getItem('excelColumnMapping');
    return saved ? JSON.parse(saved) : {
      name: null,
      priceYuan: null,
      quantity: null,
      sellingPrice: null,
      platform: null,
      link: null,
      barcode: null
    };
  });

  if (!show) return null;

  const addItem = () => {
    setItems([...items, { name: '', priceYuan: 0, quantity: 1 }]);
  };

  const removeItem = (index: number) => {
    setItems(items.filter((_, i) => i !== index));
  };

  const updateItem = (index: number, field: keyof Purchase, value: any) => {
    const newItems = [...items];
    newItems[index] = { ...newItems[index], [field]: value };
    setItems(newItems);
  };

  const handleAddFromText = () => {
    const rows = textInput.split('\n').filter(row => row.trim() !== '');
    const newItems = rows.map(row => {
      const cols = row.split('|').map(c => c.trim()).filter(c => c !== '');
      if (cols.length < 3) return null;
      return { 
        name: cols[1], 
        priceYuan: parseFloat(cols[2]) || 0, 
        quantity: parseInt(cols[3]) || 0 
      };
    }).filter(item => item !== null) as Partial<Purchase>[];
    setItems([...items, ...newItems]);
    setTextInput('');
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
      setPreviewData(data);
    };
    reader.readAsBinaryString(file);
  };

  const processData = () => {
    const dataToProcess = previewData.slice(startRow - 1);
    console.log('Data to process:', dataToProcess);
    const newItems = dataToProcess.map(row => {
      return { 
        name: columnMapping.name !== null ? String(row[columnMapping.name] || '') : '',
        priceYuan: columnMapping.priceYuan !== null ? parseFloat(row[columnMapping.priceYuan]) || 0 : 0,
        quantity: columnMapping.quantity !== null ? parseInt(row[columnMapping.quantity]) || 0 : 0,
        sellingPrice: columnMapping.sellingPrice !== null ? parseFloat(row[columnMapping.sellingPrice]) || 0 : 0,
        platform: columnMapping.platform !== null ? String(row[columnMapping.platform] || '') : '',
        link: columnMapping.link !== null ? String(row[columnMapping.link] || '') : '',
        barcode: columnMapping.barcode !== null ? String(row[columnMapping.barcode] || '') : ''
      };
    }).filter(item => item.name);
    console.log('Processed new items:', newItems);
    setItems([...items, ...newItems]);
    setPreviewData([]);
    setFileName('');
  };

  const mappingOptions = [
    { key: '', label: '-- Ігнорувати --' },
    { key: 'name', label: 'Назва товару*' },
    { key: 'priceYuan', label: 'Ціна ¥' },
    { key: 'quantity', label: 'К-сть' },
    { key: 'sellingPrice', label: 'Ціна 1 шт (грн)' },
    { key: 'platform', label: 'Сайт/Опис' },
    { key: 'link', label: 'Посилання на товар' },
    { key: 'barcode', label: 'Штрих-код' }
  ];

  const handleMappingChange = (colIndex: number, value: string) => {
    const newMapping = { ...columnMapping };
    // Remove this colIndex from any other key
    Object.keys(newMapping).forEach(key => {
      if (newMapping[key] === colIndex) {
        newMapping[key] = null;
      }
    });
    if (value) {
      newMapping[value] = colIndex;
    }
    setColumnMapping(newMapping);
    localStorage.setItem('excelColumnMapping', JSON.stringify(newMapping));
  };

  const getSelectedValueForCol = (colIndex: number) => {
    const found = Object.entries(columnMapping).find(([_, index]) => index === colIndex);
    return found ? found[0] : '';
  };

  // Get max columns from the first few rows to ensure we render enough dropdowns
  const maxCols = previewData.length > 0 ? Math.max(...previewData.slice(0, 10).map(row => row.length), 5) : 0;

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[100] p-4">
      <div className="bg-white rounded-3xl p-8 max-w-6xl w-full shadow-2xl border border-gray-200 max-h-[90vh] overflow-y-auto">
        <div className="flex justify-between items-center mb-6">
          <h3 className="text-2xl font-black text-gray-900 uppercase tracking-tight">Імпорт Прайсу (Excel/CSV)</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-red-500 transition-colors">
            <X className="w-6 h-6" />
          </button>
        </div>

        {!previewData.length ? (
          <div className="mb-6 space-y-6">
            <div className="p-6 bg-gray-50 rounded-2xl border border-gray-200 flex flex-col items-center justify-center min-h-[200px] gap-4">
              <div className="w-16 h-16 bg-yellow-400/10 rounded-full flex items-center justify-center">
                <FileSpreadsheet className="w-8 h-8 text-yellow-500" />
              </div>
              <div className="text-center">
                <h4 className="text-lg font-bold text-gray-900 mb-2">Завантажте файл Excel</h4>
                <p className="text-sm text-gray-500 mb-4">Підтримуються формати .xlsx, .xls</p>
                <label className="inline-flex items-center gap-2 px-8 py-4 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-xs hover:bg-yellow-500 transition-all cursor-pointer shadow-lg shadow-yellow-400/20">
                  <Upload className="w-4 h-4" />
                  Вибрати файл
                  <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" />
                </label>
              </div>
            </div>
          </div>
        ) : (
          <div className="mb-6 space-y-4">
            <div className="flex flex-wrap items-center gap-4 bg-gray-50 p-4 rounded-xl border border-gray-200">
              <div className="flex items-center gap-2 bg-white px-4 py-2 rounded-lg border border-gray-200 shadow-sm">
                <FileSpreadsheet className="w-4 h-4 text-emerald-500" />
                <span className="text-sm font-bold text-gray-900">{fileName}</span>
              </div>
              
              <div className="flex items-center gap-3 ml-auto">
                <AlertTriangle className="w-4 h-4 text-yellow-500" />
                <span className="text-sm font-bold text-gray-600">Старт з рядка:</span>
                <input 
                  type="number" 
                  min="1"
                  value={startRow || 1}
                  onChange={(e) => setStartRow(Math.max(1, parseInt(e.target.value) || 1))}
                  className="w-20 bg-white border border-gray-200 text-gray-900 rounded-lg px-3 py-2 text-sm font-bold focus:outline-none focus:border-yellow-400 shadow-sm"
                />
              </div>
            </div>

            <div className="overflow-x-auto rounded-xl border border-gray-200 bg-white shadow-sm">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="p-3 border-r border-gray-200 text-center w-12 text-gray-500 font-black">#</th>
                    {Array.from({ length: maxCols }).map((_, colIndex) => (
                      <th key={colIndex} className="p-2 border-r border-gray-200 min-w-[200px]">
                        <select 
                          value={getSelectedValueForCol(colIndex)}
                          onChange={(e) => handleMappingChange(colIndex, e.target.value)}
                          className="w-full bg-white border border-gray-200 text-gray-900 rounded-lg px-3 py-2 text-xs font-bold focus:outline-none focus:border-yellow-400 cursor-pointer shadow-sm"
                        >
                          {mappingOptions.map(opt => (
                            <option key={opt.key} value={opt.key}>{opt.label}</option>
                          ))}
                        </select>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewData.slice(startRow - 1, startRow - 1 + 10).map((row, rowIndex) => (
                    <tr key={rowIndex} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                      <td className="p-3 border-r border-gray-100 text-center text-gray-500 font-mono text-xs">
                        {startRow + rowIndex}
                      </td>
                      {Array.from({ length: maxCols }).map((_, colIndex) => (
                        <td key={colIndex} className="p-3 border-r border-gray-100 text-gray-700 text-xs truncate max-w-[200px]">
                          {row[colIndex] !== undefined ? String(row[colIndex]) : ''}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            
            <div className="flex justify-between items-center mt-4">
              <p className="text-xs text-gray-500 font-bold">Показано {Math.min(10, previewData.length - startRow + 1)} з {previewData.length - startRow + 1} рядків</p>
              <div className="flex gap-3">
                <button 
                  onClick={() => {
                    setPreviewData([]);
                    setFileName('');
                  }} 
                  className="px-6 py-3 bg-white text-gray-700 rounded-xl font-black uppercase tracking-widest text-xs hover:bg-gray-50 transition-all border border-gray-200 shadow-sm"
                >
                  Скасувати
                </button>
                <button 
                  onClick={processData} 
                  className="px-8 py-3 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-xs hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-400/20 flex items-center gap-2"
                >
                  <Upload className="w-4 h-4" />
                  Завантажити в базу
                </button>
              </div>
            </div>
          </div>
        )}

        {items.length > 1 && (
          <div className="mt-8 pt-8 border-t border-gray-200">
            <h4 className="text-lg font-black text-gray-900 uppercase tracking-tight mb-4">Підготовлені до збереження ({items.length - 1})</h4>
            <div className="overflow-x-auto rounded-xl border border-gray-200 bg-white shadow-sm max-h-[400px]">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-50 z-10">
                  <tr className="border-b border-gray-200">
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Назва</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Ціна ¥</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">К-сть</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Ціна 1 шт (грн)</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Сайт/Опис</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Посилання</th>
                    <th className="p-4 text-[10px] font-black text-gray-500 uppercase tracking-widest">Штрих-код</th>
                    <th className="p-4"></th>
                  </tr>
                </thead>
                <tbody>
                  {items.slice(1).map((item, index) => (
                    <tr key={index} className="border-b border-gray-100 hover:bg-gray-50">
                      <td className="p-4">
                        <input
                          type="text"
                          value={item.name || ''}
                          onChange={(e) => updateItem(index + 1, 'name', e.target.value)}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                          placeholder="Назва товару"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="number"
                          value={item.priceYuan || 0}
                          onChange={(e) => updateItem(index + 1, 'priceYuan', parseFloat(e.target.value))}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="number"
                          value={item.quantity || 0}
                          onChange={(e) => updateItem(index + 1, 'quantity', parseInt(e.target.value))}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="number"
                          value={item.sellingPrice || 0}
                          onChange={(e) => updateItem(index + 1, 'sellingPrice', parseFloat(e.target.value))}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="text"
                          value={item.platform || ''}
                          onChange={(e) => updateItem(index + 1, 'platform', e.target.value)}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                          placeholder="Сайт/Опис"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="text"
                          value={item.link || ''}
                          onChange={(e) => updateItem(index + 1, 'link', e.target.value)}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                          placeholder="Посилання"
                        />
                      </td>
                      <td className="p-4">
                        <input
                          type="text"
                          value={item.barcode || ''}
                          onChange={(e) => updateItem(index + 1, 'barcode', e.target.value)}
                          className="w-full p-2 bg-white border border-gray-200 rounded-lg text-sm font-bold text-gray-900 focus:border-yellow-400 outline-none shadow-sm"
                          placeholder="Штрих-код"
                        />
                      </td>
                      <td className="p-4 text-right">
                        <button onClick={() => removeItem(index + 1)} className="text-gray-400 hover:text-red-500 transition-colors p-2">
                          <Trash2 className="w-5 h-5" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="flex justify-end gap-4 mt-6">
              <button 
                onClick={() => onSave(items.slice(1))}
                className="px-8 py-4 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-xs hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-400/20"
              >
                Зберегти всі товари
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
