import React, { useState } from 'react';
import { X, Plus, Trash2, Upload, Clipboard } from 'lucide-react';
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

  const handlePaste = () => {
    navigator.clipboard.readText().then(text => {
      const rows = text.split('\n').filter(row => row.trim() !== '');
      const newItems = rows.map(row => {
        const cols = row.split('|').map(c => c.trim()).filter(c => c !== '');
        if (cols.length < 3) return null;
        // Based on user example: | 149 | Розвальцовка труб | 98.0 | 1 | 98.0 | 617.00
        return { name: cols[1], priceYuan: parseFloat(cols[2]), quantity: parseInt(cols[3]) };
      }).filter(item => item !== null) as Partial<Purchase>[];
      setItems([...items, ...newItems]);
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
      const newItems = data.slice(1).map(row => {
        // Based on Excel screenshot: 
        // 1: Name, 2: Price, 3: Quantity
        return { name: row[1], priceYuan: parseFloat(row[2]), quantity: parseInt(row[3]) };
      }).filter(item => item.name);
      setItems([...items, ...newItems]);
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[100] p-4">
      <div className="bg-white rounded-3xl p-8 max-w-4xl w-full shadow-2xl border border-gray-100 max-h-[90vh] overflow-y-auto">
        <div className="flex justify-between items-center mb-6">
          <h3 className="text-2xl font-black text-black uppercase tracking-tight">Масове додавання товарів</h3>
          <div className="flex gap-2">
            <button onClick={handlePaste} className="p-2 text-gray-500 hover:text-black">
              <Clipboard className="w-6 h-6" />
            </button>
            <label className="p-2 text-gray-500 hover:text-black cursor-pointer">
              <Upload className="w-6 h-6" />
              <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" />
            </label>
            <button onClick={onClose} className="text-gray-400 hover:text-red-500 transition-colors">
              <X className="w-6 h-6" />
            </button>
          </div>
        </div>
        <table className="w-full mb-6">
          <thead>
            <tr className="text-left text-[10px] font-black text-gray-400 uppercase tracking-widest">
              <th className="pb-2">Товар</th>
              <th className="pb-2">Ціна ¥</th>
              <th className="pb-2">К-сть</th>
              <th className="pb-2"></th>
            </tr>
          </thead>
          <tbody>
            {items.map((item, index) => (
              <tr key={index}>
                <td className="pr-4 pb-2">
                  <input
                    type="text"
                    value={item.name}
                    onChange={(e) => updateItem(index, 'name', e.target.value)}
                    className="w-full p-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold"
                    placeholder="Назва товару"
                  />
                </td>
                <td className="pr-4 pb-2">
                  <input
                    type="number"
                    value={item.priceYuan}
                    onChange={(e) => updateItem(index, 'priceYuan', parseFloat(e.target.value))}
                    className="w-full p-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold"
                  />
                </td>
                <td className="pr-4 pb-2">
                  <input
                    type="number"
                    value={item.quantity}
                    onChange={(e) => updateItem(index, 'quantity', parseInt(e.target.value))}
                    className="w-full p-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold"
                  />
                </td>
                <td className="pb-2">
                  <button onClick={() => removeItem(index)} className="text-red-500 hover:text-red-700">
                    <Trash2 className="w-5 h-5" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="flex gap-4">
          <button onClick={addItem} className="flex items-center gap-2 text-yellow-500 font-bold text-sm">
            <Plus className="w-4 h-4" /> Додати рядок
          </button>
          <button 
            onClick={() => onSave(items)}
            className="flex-1 py-4 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-xs hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-100"
          >
            Зберегти
          </button>
        </div>
      </div>
    </div>
  );
};
