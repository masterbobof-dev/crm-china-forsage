/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo, useEffect } from 'react';
import { BulkAddItemsModal } from './components/BulkAddItemsModal';
import { QrReader } from 'react-qr-reader';
import { ErrorBoundary } from './components/ErrorBoundary';
import JsBarcode from 'jsbarcode';
import { trackParcel, getDocumentList } from './services/novaposhta';
import Cropper from 'react-easy-crop';
import { domToPng } from 'modern-screenshot';
import { jsPDF } from 'jspdf';
import { 
  Calculator, 
  Package, 
  Plane, 
  Ship, 
  Train, 
  Truck, 
  Info, 
  ChevronLeft,
  ChevronRight,
  Scale, 
  Maximize,
  Layers,
  ShieldCheck,
  Scissors,
  Minimize2,
  Mail,
  MessageCircle,
  Globe,
  ExternalLink,
  ShoppingCart,
  Warehouse,
  Merge as Combine,
  Home,
  BarChart3,
  LayoutDashboard,
  Search,
  Plus,
  Clock,
  Flag,
  X,
  Upload,
  FileText,
  ClipboardList,
  PlusCircle,
  ArrowRight,
  Calendar,
  MapPin,
  DollarSign,
  CheckCircle2,
  ChevronDown,
  Trash2,
  Edit2,
  Download,
  Receipt,
  UserCircle,
  Printer,
  LayoutGrid,
  List,
  Store,
  RotateCcw,
  Settings,
  CheckSquare,
  Square,
  MinusSquare,
  Zap,
  Copy,
  Camera,
  Smartphone
} from 'lucide-react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { motion } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { supabase } from './supabase';
import { useSupabaseSync, toSnakeCase } from './hooks/useSupabaseSync';

// Utility for Tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---

interface DensityTier {
  min: number;
  max: number | null;
  price: number;
  unit: 'kg' | 'm3';
}

interface Tariff {
  id: string;
  name: string;
  iconName: 'Ship' | 'Plane' | 'Train' | 'Truck' | 'Zap';
  pricePerKg?: number;
  volumetricFactor?: number;
  densityTiers?: DensityTier[];
  deliveryDays: string;
  description: string;
  localDeliveryPrice?: number; // $/kg
  minWeight?: number;
  minVolume?: number;
  minCost?: number; // $
  insuranceRate?: number; // % (e.g. 2 for 2%)
  packagingCost?: number; // $/kg
  packagingCostPerM3?: number; // $/m3
  customsFee?: number; // $
  handlingFee?: number; // $
  fuelSurcharge?: number; // %
}

interface ParcelData {
  weight: number; // kg
  length: number; // cm
  width: number; // cm
  height: number; // cm
  volume?: number; // m3
  declaredValue?: number; // USD
  isFabric?: boolean;
  isPressed?: boolean;
  isInsured?: boolean;
  density?: number; // kg/m3
}

// --- Constants ---

const THEME_BLACK = "#000000";
const THEME_YELLOW = "#facc15";
const TURBO_RED = "#facc15"; // THEME_YELLOW
const TURBO_GREEN = "#000000"; // THEME_BLACK

const DEFAULT_TARIFFS: Tariff[] = [
  {
    id: 'turbo-sea',
    name: 'Turboavia More',
    iconName: 'Ship',
    densityTiers: [
      { min: 0, max: 100, price: 480, unit: 'm3' },
      { min: 100, max: 200, price: 580, unit: 'm3' },
      { min: 200, max: 250, price: 610, unit: 'm3' },
      { min: 250, max: 300, price: 640, unit: 'm3' },
      { min: 300, max: 350, price: 690, unit: 'm3' },
      { min: 350, max: 400, price: 740, unit: 'm3' },
      { min: 400, max: 450, price: 790, unit: 'm3' },
      { min: 450, max: 500, price: 840, unit: 'm3' },
      { min: 500, max: 700, price: 2.1, unit: 'kg' },
      { min: 700, max: 800, price: 2.3, unit: 'kg' },
      { min: 800, max: null, price: 2.5, unit: 'kg' },
    ],
    localDeliveryPrice: 0.4,
    deliveryDays: '65-75 днів',
    description: 'Морська доставка Turboavia з тарифікацією за щільністю.',
    insuranceRate: 2,
    minWeight: 10,
    packagingCostPerM3: 30
  },
  {
    id: 'air-fast',
    name: 'СПЕЦ. Turboavia Авіа',
    iconName: 'Plane',
    densityTiers: [
      { min: 0, max: 60, price: 1336, unit: 'm3' },
      { min: 60, max: 80, price: 16.0, unit: 'kg' },
      { min: 80, max: 100, price: 15.0, unit: 'kg' },
      { min: 100, max: 120, price: 14.0, unit: 'kg' },
      { min: 120, max: 140, price: 13.6, unit: 'kg' },
      { min: 140, max: 160, price: 13.3, unit: 'kg' },
      { min: 160, max: null, price: 13.0, unit: 'kg' },
    ],
    deliveryDays: '18-20 днів',
    description: 'Спеціальна авіа доставка Turboavia.',
    insuranceRate: 2,
    minCost: 15,
    packagingCost: 0.5
  },
  {
    id: 'train',
    name: 'Залізниця',
    iconName: 'Train',
    pricePerKg: 4.5,
    volumetricFactor: 5000,
    deliveryDays: '35-45 днів',
    description: 'Надійний спосіб для середніх вантажів.',
    insuranceRate: 2,
    minWeight: 15,
    packagingCost: 0.2
  }
];

const getTariffIcon = (name: string) => {
  switch (name) {
    case 'Ship': return <Ship className="w-5 h-5" />;
    case 'Plane': return <Plane className="w-5 h-5" />;
    case 'Train': return <Train className="w-5 h-5" />;
    case 'Truck': return <Truck className="w-5 h-5" />;
    default: return <Package className="w-5 h-5" />;
  }
};

// --- Components ---

type CalculatorType = 'international' | 'novaposhta' | 'transfer';

type CRMModule = 'dashboard' | 'purchases' | 'china_warehouse' | 'consolidation' | 'ua_warehouse' | 'local_delivery' | 'my_warehouse' | 'issue_to_store' | 'settings' | 'calculator' | 'pinduoduo';

interface CRMModuleConfig {
  id: CRMModule;
  title: string;
  icon: any;
  description: string;
}

interface NovaPoshtaData {
  weight: number;
  length: number;
  width: number;
  height: number;
  declaredValue: number;
  destination: 'city' | 'region' | 'ukraine';
}

interface MoneyTransferData {
  amount: number;
  method: 'card' | 'cash';
}

export interface Purchase {
  id: string;
  platform: string;
  name: string;
  link: string;
  priceYuan: number;
  exchangeRate: number;
  quantity: number;
  trackNumber: string;
  photo: string;
  comment: string;
  size?: string;
  width?: number;
  height?: number;
  length?: number;
  dimUnit?: 'cm' | 'm';
  weight?: number;
  weightUnit?: 'g' | 'kg';
  volume?: number;
  density?: number;
  isFabric?: boolean;
  isPressed?: boolean;
  isInsured?: boolean;
  declaredValue?: number;
  shippingCost?: number;
  status: 'purchased' | 'shipped_by_seller' | 'arrived_china' | 'at_china_warehouse' | 'shipped_to_ua' | 'arrived_ua' | 'local_delivery' | 'my_warehouse' | 'sold';
  arrivalDate?: string;
  batchId?: string;
  localDeliveryId?: string;
  deliveryCostPerItem?: number; // This is "Доставка Китай"
  ukraineDeliveryCost?: number; // This is "Доставка Україна" (local logistics)
  novaPoshtaCost?: number;      // This is "Доставка Нова Пошта"
  sellingPrice?: number;
  soldDate?: string;
  markup?: boolean;
  markupValue?: number;
  barcode?: string;
  saleDestination?: 'physical_store' | 'online_store' | 'personal_use';
  createdAt: string;
  brand?: string;
  radius?: string;
  season?: string;
  article?: string;
}

interface LocalDeliveryGroup {
  id: string;
  name: string;
  type: 'novaposhta' | 'pickup';
  trackingNumber?: string;
  cost: number;
  itemIds: string[];
  status: 'pending' | 'received';
  createdAt: string;
}

type UserRole = 'admin' | 'manager';

interface Batch {
  id: string;
  name: string;
  shipmentDate: string;
  warehouse: string;
  deliveryType: 'sea' | 'air';
  status: 'shipped' | 'arrived_ua';
  totalWeight?: number;
  deliveryCost?: number;
  pricePerKg?: number;
  itemIds: string[];
  createdAt: string;
}

const statusLabels: Record<string, string> = {
  purchased: 'Куплено',
  shipped_by_seller: 'Відправлено продавцем',
  arrived_china: 'Прибуло в Китай',
  at_china_warehouse: 'Склад Китай',
  shipped_to_ua: 'Відправлено з Китаю',
  arrived_ua: 'На складі Україна Київ',
  local_delivery: 'Відправлено Новою Поштою',
  my_warehouse: 'Прибуло на мій склад',
  sold: 'Видано на магазин'
};

const CropModal = ({ image, onCropComplete, onCancel }: { image: string, onCropComplete: (croppedImage: string) => void, onCancel: () => void }) => {
  const [crop, setCrop] = useState({ x: 0, y: 0 });
  const [zoom, setZoom] = useState(1);
  const [croppedAreaPixels, setCroppedAreaPixels] = useState<any>(null);

  const onCropChange = (crop: { x: number, y: number }) => setCrop(crop);
  const onZoomChange = (zoom: number) => setZoom(zoom);

  const createImage = (url: string): Promise<HTMLImageElement> =>
    new Promise((resolve, reject) => {
      const image = new Image();
      image.addEventListener('load', () => resolve(image));
      image.addEventListener('error', (error) => reject(error));
      image.setAttribute('crossOrigin', 'anonymous');
      image.src = url;
    });

  const getCroppedImg = async (imageSrc: string, pixelCrop: any): Promise<string> => {
    const image = await createImage(imageSrc);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');

    if (!ctx) return '';

    canvas.width = pixelCrop.width;
    canvas.height = pixelCrop.height;

    ctx.drawImage(
      image,
      pixelCrop.x,
      pixelCrop.y,
      pixelCrop.width,
      pixelCrop.height,
      0,
      0,
      pixelCrop.width,
      pixelCrop.height
    );

    return canvas.toDataURL('image/jpeg');
  };

  const handleCrop = async () => {
    try {
      const croppedImage = await getCroppedImg(image, croppedAreaPixels);
      onCropComplete(croppedImage);
    } catch (e) {
      console.error(e);
    }
  };

  return (
    <div className="fixed inset-0 z-[400] flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
      <div className="bg-white rounded-3xl overflow-hidden max-w-2xl w-full flex flex-col h-[80vh] md:h-[70vh]">
        <div className="p-4 md:p-6 border-b flex justify-between items-center">
          <h3 className="text-lg md:text-xl font-black text-black uppercase tracking-tight">Обрізати фото</h3>
          <button onClick={onCancel} className="text-gray-400 hover:text-red-500 transition-colors">
            <X className="w-5 h-5 md:w-6 md:h-6" />
          </button>
        </div>
        <div className="relative flex-1 bg-gray-900">
          <Cropper
            image={image}
            crop={crop}
            zoom={zoom}
            aspect={1}
            onCropChange={onCropChange}
            onCropComplete={(_, pixels) => setCroppedAreaPixels(pixels)}
            onZoomChange={onZoomChange}
          />
        </div>
        <div className="p-4 md:p-6 bg-white border-t space-y-4">
          <div className="flex items-center gap-4">
            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Зум</span>
            <input
              type="range"
              value={zoom}
              min={1}
              max={3}
              step={0.1}
              aria-labelledby="Zoom"
              onChange={(e) => setZoom(Number(e.target.value))}
              className="flex-1 accent-yellow-400"
            />
          </div>
          <div className="flex gap-3 md:gap-4">
            <button
              onClick={onCancel}
              className="flex-1 py-3 md:py-4 bg-gray-50 text-gray-400 rounded-xl font-black uppercase tracking-widest text-[10px] md:text-xs hover:bg-gray-100 transition-all"
            >
              Скасувати
            </button>
            <button
              onClick={handleCrop}
              className="flex-1 py-3 md:py-4 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-[10px] md:text-xs hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-100"
            >
              Застосувати
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
const DeleteConfirmModal = ({ show, title, message, onConfirm, onCancel }: { show: boolean, title: string, message: string, onConfirm: () => void, onCancel: () => void }) => {
  if (!show) return null;
  return (
    <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-[100] p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl border border-gray-100"
      >
        <div className="w-16 h-16 bg-yellow-50 rounded-2xl flex items-center justify-center mb-6">
          <Trash2 className="w-8 h-8 text-yellow-500" />
        </div>
        <h3 className="text-2xl font-black text-black mb-2 uppercase tracking-tight">{title}</h3>
        <p className="text-gray-500 font-medium mb-8 leading-relaxed">{message}</p>
        <div className="flex gap-4">
          <button 
            onClick={onCancel}
            className="flex-1 py-4 bg-gray-50 text-gray-400 rounded-xl font-black uppercase tracking-widest text-xs hover:bg-gray-100 transition-all"
          >
            Скасувати
          </button>
          <button 
            onClick={onConfirm}
            className="flex-1 py-4 bg-yellow-400 text-black rounded-xl font-black uppercase tracking-widest text-xs hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-100"
          >
            Видалити
          </button>
        </div>
      </motion.div>
    </div>
  );
};

export default function App() {
  const [view, setView] = useState<'home' | 'calculator' | 'crm'>('crm');
  const [crmModule, setCrmModule] = useState<CRMModule>('purchases');
  const [tariffs, setTariffs] = useState<Tariff[]>(() => {
    const saved = localStorage.getItem('tariffs');
    return saved ? JSON.parse(saved) : DEFAULT_TARIFFS;
  });
  const [activeTab, setActiveTab] = useState<CalculatorType>('international');
  const [inputMethod, setInputMethod] = useState<'dims' | 'density'>('dims');
  const [parcel, setParcel] = useState<ParcelData>({
    weight: 0,
    length: 0,
    width: 0,
    height: 0,
    volume: 0,
    declaredValue: 0,
    isFabric: false,
    isPressed: false,
    isInsured: false,
    density: 0
  });
  const [selectedTariffId, setSelectedTariffId] = useState<string>(DEFAULT_TARIFFS[0].id);
  const [npData, setNpData] = useState({
    weight: 0,
    length: 0,
    width: 0,
    height: 0,
    declaredValue: 0,
    cityRecipient: '',
    citySender: '',
    destination: 'ukraine'
  });
  const [transferData, setTransferData] = useState({
    amount: 0,
    currency: 'UAH' as 'UAH' | 'USD' | 'EUR',
    method: 'card'
  });

  const selectedTariff = useMemo(() => tariffs.find(t => t.id === selectedTariffId) || tariffs[0], [tariffs, selectedTariffId]);

  const { density, finalVolumeM3 } = useMemo(() => {
    let volM3 = 0;
    if (inputMethod === 'dims') {
      volM3 = (parcel.length * parcel.width * parcel.height) / 1000000;
    } else {
      volM3 = parcel.volume;
    }
    const d = volM3 > 0 ? parcel.weight / volM3 : 0;
    return { density: d, finalVolumeM3: volM3 };
  }, [parcel, inputMethod]);

  const [userRole, setUserRole] = useState<UserRole>('admin');
  const [searchQuery, setSearchQuery] = useState('');
  const [purchaseSearch, setPurchaseSearch] = useState('');
  const [chinaWarehouseSearch, setChinaWarehouseSearch] = useState('');
  const [uaWarehouseSearch, setUaWarehouseSearch] = useState('');
  const [salesSearch, setSalesSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [notifications, setNotifications] = useState<{id: string, text: string, time: string, type?: 'success' | 'info' | 'error'}[]>([]);
  
  const addNotification = (text: string, type: 'success' | 'info' | 'error' = 'info') => {
    setNotifications(prev => [{ id: Math.random().toString(36).substr(2, 9), text, time: 'Щойно', type }, ...prev].slice(0, 5));
  };
  
  const [showAssignTrackModal, setShowAssignTrackModal] = useState(false);
  const [assignTrackNumber, setAssignTrackNumber] = useState('');

  const handleBulkSave = async (items: Partial<Purchase>[]) => {
    const newPurchases = items.map(item => ({
      id: Math.random().toString(36).substr(2, 9),
      name: item.name || '',
      platform: item.platform || 'Taobao',
      link: item.link || '',
      priceYuan: item.priceYuan || 0,
      quantity: item.quantity || 1,
      sellingPrice: item.sellingPrice || 0,
      exchangeRate: 5.5,
      trackNumber: '',
      photo: '',
      comment: '',
      status: 'purchased' as Purchase['status'],
      createdAt: new Date().toISOString()
    }));
    
    const snakeCasePurchases = toSnakeCase(newPurchases);
    // Save to Supabase
    const { error } = await supabase.from('purchases').insert(snakeCasePurchases);
    if (error) {
      console.error('Supabase insert error:', error);
      addNotification('Помилка при збереженні: ' + error.message, 'error');
    } else {
      addNotification('Товари успішно додані', 'success');
      setPurchases([...newPurchases as Purchase[], ...purchases]);
      setShowBulkAddModal(false);
    }
  };

  useEffect(() => {
    const checkSupabase = async () => {
      try {
        const { error } = await supabase.from('purchases').select('*').limit(1);
        if (error) {
          console.error('Supabase connection error:', error.message);
          addNotification('Помилка з\'єднання з Supabase: ' + error.message, 'error');
        } else {
          console.log('Successfully connected to Supabase!');
          addNotification('Успішне з\'єднання з сервером Supabase!', 'success');
        }
      } catch (err: any) {
        console.error('Supabase connection exception:', err);
        addNotification('Помилка з\'єднання з Supabase: ' + err.message, 'error');
      }
    };
    checkSupabase();
  }, []);
  const [dashboardStats, setDashboardStats] = useState({
    inTransitToChina: 12,
    atChinaWarehouse: 45,
    inTransitToUA: 28,
    atUAWarehouse: 15
  });
  const [showAddPurchaseModal, setShowAddPurchaseModal] = useState(false);
  const [showBulkAddModal, setShowBulkAddModal] = useState(false);
  const [selectedTrackNumber, setSelectedTrackNumber] = useState<string | null>(null);
  const [showImportTracksModal, setShowImportTracksModal] = useState(false);
  const [bulkImportText, setBulkImportText] = useState('');
  const [confirmModal, setConfirmModal] = useState<{show: boolean, title: string, message: string, onConfirm: () => void} | null>(null);
  const [showCreateBatchModal, setShowCreateBatchModal] = useState(false);
  const [showCreateLocalDeliveryModal, setShowCreateLocalDeliveryModal] = useState(false);
  const [localDeliveryForm, setLocalDeliveryForm] = useState({
    name: '',
    type: 'novaposhta' as 'novaposhta' | 'pickup',
    trackingNumber: '',
    cost: 0
  });
  const [editingLocalDeliveryId, setEditingLocalDeliveryId] = useState<string | null>(null);
  const [showCostModal, setShowCostModal] = useState<{show: boolean, batchId?: string}>({show: false});
  const [cropImage, setCropImage] = useState<string | null>(null);
  const [costForm, setCostForm] = useState({
    totalWeight: 0,
    deliveryCost: 0,
    volume: 0,
    declaredValue: 0,
    isInsured: false,
    isFabric: false,
    isPressed: false,
    tariffId: 'turbo-sea'
  });
  const [purchases, setPurchases, loadingPurchases] = useSupabaseSync<Purchase>('purchases', []);
  const [showTariffModal, setShowTariffModal] = useState<{show: boolean, tariffId: string | null}>({ show: false, tariffId: null });
  const [tariffForm, setTariffForm] = useState<Partial<Tariff>>({
    name: '',
    iconName: 'Ship',
    deliveryDays: '',
    description: '',
    pricePerKg: 0,
    volumetricFactor: 0,
    localDeliveryPrice: 0,
    minWeight: 0,
    minVolume: 0,
    insuranceRate: 2,
    packagingCost: 0,
    customsFee: 0,
    handlingFee: 0,
    fuelSurcharge: 0
  });

  const [batches, setBatches, loadingBatches] = useSupabaseSync<Batch>('batches', []);
  const [localDeliveries, setLocalDeliveries] = useState<LocalDeliveryGroup[]>(() => {
    const saved = localStorage.getItem('localDeliveries');
    return saved ? JSON.parse(saved) : [];
  });
  const [batchForm, setBatchForm] = useState({
    name: '',
    shipmentDate: new Date().toISOString().split('T')[0],
    warehouse: 'Guangzhou',
    deliveryType: 'sea' as 'sea' | 'air'
  });
  const [cnyToUah, setCnyToUah] = useState(() => parseFloat(localStorage.getItem('cnyToUah') || '5.5'));
  const [usdToUah, setUsdToUah] = useState(() => parseFloat(localStorage.getItem('usdToUah') || '40'));
  const [novaPoshtaApiKey, setNovaPoshtaApiKey] = useState(() => localStorage.getItem('novaPoshtaApiKey') || '');
  const [priceListMargin, setPriceListMargin] = useState(() => parseFloat(localStorage.getItem('priceListMargin') || '0'));

  const handleBulkDelete = async (ids: string[]) => {
    const { error } = await supabase.from('purchases').delete().in('id', ids);
    if (error) {
      addNotification('Помилка при видаленні: ' + error.message, 'error');
    } else {
      addNotification('Товари успішно видалені', 'success');
      setPurchases(purchases.filter(p => !ids.includes(p.id)));
      setSelectedPurchaseIds([]);
    }
  };

  useEffect(() => {
    localStorage.setItem('cnyToUah', cnyToUah.toString());
    localStorage.setItem('usdToUah', usdToUah.toString());
    localStorage.setItem('novaPoshtaApiKey', novaPoshtaApiKey);
    localStorage.setItem('tariffs', JSON.stringify(tariffs));
    localStorage.setItem('localDeliveries', JSON.stringify(localDeliveries));
    localStorage.setItem('priceListMargin', priceListMargin.toString());
  }, [cnyToUah, usdToUah, novaPoshtaApiKey, tariffs, localDeliveries, priceListMargin]);

  const internationalDetails = useMemo(() => {
    let shippingCost = 0;
    
    if (selectedTariff.densityTiers && selectedTariff.densityTiers.length > 0) {
      const tier = selectedTariff.densityTiers.find(t => density >= t.min && (t.max === null || density < t.max));
      if (tier) {
        if (tier.unit === 'kg') {
          shippingCost = parcel.weight * tier.price;
        } else {
          shippingCost = finalVolumeM3 * tier.price;
        }
      } else {
        const lastTier = selectedTariff.densityTiers[selectedTariff.densityTiers.length - 1];
        if (lastTier.unit === 'kg') {
          shippingCost = parcel.weight * lastTier.price;
        } else {
          shippingCost = finalVolumeM3 * lastTier.price;
        }
      }
    } else if (selectedTariff.pricePerKg) {
      const volumetricWeight = selectedTariff.volumetricFactor ? (finalVolumeM3 * 1000000) / selectedTariff.volumetricFactor : 0;
      const chargeableWeight = Math.max(parcel.weight, volumetricWeight);
      shippingCost = chargeableWeight * selectedTariff.pricePerKg;
    }
    
    if (selectedTariff.minCost && shippingCost < selectedTariff.minCost) {
      shippingCost = selectedTariff.minCost;
    }

    const insurance = parcel.isInsured && parcel.declaredValue ? (parcel.declaredValue * (selectedTariff.insuranceRate || 2)) / 100 : 0;
    const fabricSurcharge = parcel.isFabric ? parcel.weight * 0.5 : 0;
    const pressingCost = parcel.isPressed ? parcel.weight * 0.2 : 0;
    const packagingCost = (selectedTariff.packagingCost ? parcel.weight * selectedTariff.packagingCost : 0) + 
                          (selectedTariff.packagingCostPerM3 ? finalVolumeM3 * selectedTariff.packagingCostPerM3 : 0);
    const customsFee = selectedTariff.customsFee || 0;
    const handlingFee = selectedTariff.handlingFee || 0;
    const fuelSurcharge = selectedTariff.fuelSurcharge ? (shippingCost * selectedTariff.fuelSurcharge) / 100 : 0;
    const localDelivery = selectedTariff.localDeliveryPrice ? parcel.weight * selectedTariff.localDeliveryPrice : 0;

    const total = shippingCost + insurance + fabricSurcharge + pressingCost + packagingCost + customsFee + handlingFee + fuelSurcharge + localDelivery;

    return {
      shippingCost,
      insurance,
      fabricSurcharge,
      pressingCost,
      packagingCost,
      customsFee,
      handlingFee,
      fuelSurcharge,
      localDelivery,
      total,
      shippingCostUah: shippingCost * usdToUah,
      insuranceUah: insurance * usdToUah,
      fabricSurchargeUah: fabricSurcharge * usdToUah,
      pressingCostUah: pressingCost * usdToUah,
      packagingCostUah: packagingCost * usdToUah,
      customsFeeUah: customsFee * usdToUah,
      handlingFeeUah: handlingFee * usdToUah,
      fuelSurchargeUah: fuelSurcharge * usdToUah,
      localDeliveryUah: localDelivery * usdToUah,
      totalUah: total * usdToUah
    };
  }, [parcel, selectedTariff, density, finalVolumeM3, usdToUah]);

  const npDetails = useMemo(() => {
    const volWeight = (npData.length * npData.width * npData.height) / 4000;
    const chargeableWeight = Math.max(npData.weight, volWeight);
    const baseCost = 70 + (chargeableWeight * 5);
    const insurance = npData.declaredValue > 500 ? (npData.declaredValue * 0.005) : 0;
    const totalUah = baseCost + insurance;
    
    return {
      basePrice: baseCost,
      insurance,
      chargeableWeight,
      volumetricWeight: volWeight,
      totalUah,
      total: totalUah / usdToUah
    };
  }, [npData, usdToUah]);

  const transferDetails = useMemo(() => {
    let fee = 0;
    if (transferData.method === 'card') {
      fee = transferData.amount * 0.01 + 5;
    } else if (transferData.method === 'cash') {
      fee = transferData.amount * 0.02 + 20;
    }
    
    const total = transferData.amount + fee;
    
    let totalUah = 0;
    if (transferData.currency === 'UAH') {
      totalUah = total;
    } else if (transferData.currency === 'USD') {
      totalUah = total * usdToUah;
    } else {
      totalUah = total * usdToUah * 1.1;
    }
    
    return {
      fee,
      totalUah,
      total: transferData.currency === 'USD' ? total : totalUah / usdToUah
    };
  }, [transferData, usdToUah]);
  const [purchaseForm, setPurchaseForm] = useState({
    platform: 'Taobao',
    name: '',
    link: '',
    priceYuan: 0,
    exchangeRate: 5.5,
    quantity: 1,
    trackNumber: '',
    photo: '',
    comment: '',
    size: '',
    width: 0,
    height: 0,
    length: 0,
    dimUnit: 'cm' as 'cm' | 'm',
    weight: 0,
    weightUnit: 'kg' as 'g' | 'kg',
    volume: 0,
    density: 0,
    isFabric: false,
    isPressed: false,
    isInsured: false,
    declaredValue: 0,
    shippingCost: 0,
    status: 'purchased' as Purchase['status'],
    brand: '',
    radius: '',
    season: '',
    article: ''
  });
  const [showSaleModal, setShowSaleModal] = useState<{show: boolean, purchaseId: string | null}>({ show: false, purchaseId: null });
  const [saleForm, setSaleForm] = useState({
    sellingPrice: 0,
    novaPoshtaCost: 0,
    ukraineDeliveryCost: 0,
    markup: false,
    markupValue: 0,
    saleDestination: 'physical_store' as 'physical_store' | 'online_store' | 'personal_use'
  });
  const [trackingModal, setTrackingModal] = useState<{show: boolean, data: any | null, loading: boolean, error: string}>({ show: false, data: null, loading: false, error: '' });
  const [isUpdatingTracking, setIsUpdatingTracking] = useState(false);
  const [salesFilter, setSalesFilter] = useState<'all' | 'physical_store' | 'online_store' | 'personal_use'>('all');
  const [showStorePreview, setShowStorePreview] = useState(false);
  const [shippingPricePerKg, setShippingPricePerKg] = useState(12);
  const [selectedPurchaseIds, setSelectedPurchaseIds] = useState<string[]>([]);

  const [showImportWaybillModal, setShowImportWaybillModal] = useState(false);
  const [waybillImportText, setWaybillImportText] = useState('');

  const parseWaybillText = (text: string) => {
    const blocks = text.split(/\n\s*\n/).filter(b => b.trim().length > 0);
    const newPurchases: Purchase[] = [];
    
    blocks.forEach(block => {
      const lines = block.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      if (lines.length < 7) return;

      const trackNumber = lines[0];
      const arrivalDate = lines[1];
      const weight = parseFloat(lines[4].replace(',', '.'));
      const volume = parseFloat(lines[5].replace(',', '.'));
      const density = parseFloat(lines[6].replace(',', '.'));
      
      let shippingCost = 0;
      if (selectedTariff.densityTiers && selectedTariff.densityTiers.length > 0) {
        const tier = selectedTariff.densityTiers.find(t => density >= t.min && (t.max === null || density < t.max));
        if (tier) {
          shippingCost = tier.unit === 'kg' ? weight * tier.price : volume * tier.price;
        } else {
          const lastTier = selectedTariff.densityTiers[selectedTariff.densityTiers.length - 1];
          shippingCost = lastTier.unit === 'kg' ? weight * lastTier.price : volume * lastTier.price;
        }
      } else if (selectedTariff.pricePerKg) {
        const volumetricWeight = selectedTariff.volumetricFactor ? (volume * 1000000) / selectedTariff.volumetricFactor : 0;
        const chargeableWeight = Math.max(weight, volumetricWeight);
        shippingCost = chargeableWeight * selectedTariff.pricePerKg;
      }
      
      if (selectedTariff.minCost && shippingCost < selectedTariff.minCost) {
        shippingCost = selectedTariff.minCost;
      }

      const newPurchase: Purchase = {
        id: Math.random().toString(36).substr(2, 9),
        name: `Товар з накладної ${trackNumber}`,
        platform: 'Pinduoduo',
        link: '',
        priceYuan: 0,
        quantity: 1,
        exchangeRate: cnyToUah,
        trackNumber: trackNumber,
        photo: '',
        comment: `Імпортовано: ${arrivalDate}`,
        status: 'at_china_warehouse',
        createdAt: new Date().toISOString(),
        weight: weight,
        volume: volume,
        density: density,
        arrivalDate: arrivalDate,
        deliveryCostPerItem: shippingCost,
        shippingCost: shippingCost,
      };
      
      newPurchases.push(newPurchase);
    });
    
    return newPurchases;
  };

  const handleImportWaybills = () => {
    const imported = parseWaybillText(waybillImportText);
    if (imported.length === 0) {
      addNotification('Не вдалося розпізнати дані. Перевірте формат тексту.', 'error');
      return;
    }
    
    setPurchases([...purchases, ...imported]);
    setShowImportWaybillModal(false);
    setWaybillImportText('');
    addNotification(`Успішно імпортовано ${imported.length} накладних`, 'success');
  };

  const waybills = useMemo(() => {
    const filtered = purchases.filter(p => p.status === 'at_china_warehouse');
    const grouped = filtered.reduce((acc, p) => {
      const track = p.trackNumber || 'Без треку';
      if (!acc[track]) {
        acc[track] = {
          trackNumber: track,
          items: [],
          totalQuantity: 0,
          totalWeight: 0,
          totalCostYuan: 0,
          totalDeliveryCost: 0,
          arrivalDate: p.arrivalDate || '',
          status: p.status
        };
      }
      acc[track].items.push(p);
      acc[track].totalQuantity += p.quantity;
      acc[track].totalWeight += p.weight || 0;
      acc[track].totalCostYuan += (p.priceYuan * p.quantity);
      acc[track].totalDeliveryCost += (p.shippingCost || 0);
      return acc;
    }, {} as Record<string, any>);

    return Object.values(grouped).filter((w: any) => 
      w.trackNumber.toLowerCase().includes(chinaWarehouseSearch.toLowerCase()) ||
      w.items.some((i: any) => i.name.toLowerCase().includes(chinaWarehouseSearch.toLowerCase()))
    );
  }, [purchases, chinaWarehouseSearch]);

  const existingTrackNumbers = useMemo(() => {
    const tracks = new Set(purchases.map(p => p.trackNumber).filter(Boolean));
    return Array.from(tracks);
  }, [purchases]);

  const handleAssignTrack = (ids: string[], track: string) => {
    if (!track) {
      addNotification('Введіть трек-номер', 'error');
      return;
    }
    setPurchases(purchases.map(p => 
      ids.includes(p.id) 
        ? { ...p, trackNumber: track, status: 'at_china_warehouse' as Purchase['status'] } 
        : p
    ));
    addNotification(`Товари (${ids.length}) додано до накладної ${track}`, 'success');
    setSelectedPurchaseIds([]);
    setShowAssignTrackModal(false);
    setAssignTrackNumber('');
  };

  const handleEditWaybill = (oldTrack: string, newTrack: string) => {
    if (!newTrack || oldTrack === newTrack) return;
    setPurchases(purchases.map(p => 
      p.trackNumber === oldTrack ? { ...p, trackNumber: newTrack } : p
    ));
    addNotification(`Накладну ${oldTrack} змінено на ${newTrack}`, 'success');
  };
  const [showScanner, setShowScanner] = useState(false);
  const [scannerInput, setScannerInput] = useState('');
  const [isInventoryMode, setIsInventoryMode] = useState(false);
  const [uaWarehouseDateFilter, setUaWarehouseDateFilter] = useState('');
  const [uaWarehouseBatchFilter, setUaWarehouseBatchFilter] = useState('');
  const [priceListDateFilter, setPriceListDateFilter] = useState('');
  const [priceListBatchFilter, setPriceListBatchFilter] = useState('');
  const [priceListSearch, setPriceListSearch] = useState('');
  const [priceListView, setPriceListView] = useState<'grid' | 'table'>('grid');

  const exportToExcel = async (data: any[], fileName: string, mergeCells: boolean = true) => {
    addNotification('Генерація Excel...', 'info');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Price List');

    // Add Title Row
    const titleRow = worksheet.addRow(['CRM FORSAGE CHINA - ПРАЙС-ЛИСТ']);
    titleRow.font = { bold: true, size: 20, color: { argb: 'FF000000' } };
    worksheet.mergeCells('A1:O1');
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    titleRow.height = 40;

    // Add Date Row
    const dateRow = worksheet.addRow([`Дата генерації: ${new Date().toLocaleString()}`]);
    dateRow.font = { italic: true, size: 11, color: { argb: 'FF666666' } };
    worksheet.mergeCells('A2:O2');
    dateRow.alignment = { horizontal: 'center' };
    worksheet.addRow([]); // Empty row

    // Define columns
    worksheet.columns = [
      { header: 'Назва', key: 'name', width: 35 },
      { header: 'Опис / Коментар', key: 'comment', width: 40 },
      { header: 'Трек-номер', key: 'track', width: 25 },
      { header: 'ТТН НП', key: 'npTrack', width: 25 },
      { header: 'Платформа', key: 'platform', width: 15 },
      { header: 'Кількість', key: 'qty', width: 10 },
      { header: 'Ціна (¥)', key: 'priceYuan', width: 12 },
      { header: 'Курс (¥/₴)', key: 'rate', width: 12 },
      { header: 'Сума (₴)', key: 'sumUah', width: 15 },
      { header: 'Вага (кг)', key: 'weight', width: 12 },
      { header: 'Об\'єм (м³)', key: 'volume', width: 12 },
      { header: 'Щільність', key: 'density', width: 12 },
      { header: 'Доставка Китай (₴)', key: 'deliveryChina', width: 18 },
      { header: 'Доставка НП (₴)', key: 'deliveryNP', width: 18 },
      { header: 'Собівартість (₴)', key: 'cost', width: 15 },
      { header: 'Націнка (%)', key: 'markup', width: 12 },
      { header: 'Ціна продажу (₴)', key: 'sell', width: 18 },
      { header: 'Статус', key: 'status', width: 15 },
      { header: 'Партія', key: 'batch', width: 15 },
    ];

    // Add data
    const items = data.length > 0 && data[0].id ? data : purchases;
    
    const groupedItems: { [key: string]: typeof items } = {};
    const trackOrder: string[] = [];
    
    items.forEach((p) => {
      const track = p.trackNumber || `NO_TRACK_${p.id}`;
      if (!groupedItems[track]) {
        groupedItems[track] = [];
        trackOrder.push(track);
      }
      groupedItems[track].push(p);
    });

    let currentRow = 5; // Data starts at row 5

    trackOrder.forEach(track => {
      const group = groupedItems[track];
      const startRow = currentRow;
      
      let totalDeliveryForGroup = 0;
      let totalWeightForGroup = 0;

      group.forEach(p => {
        const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
        const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
        totalDeliveryForGroup += deliveryChinaUah + deliveryNPUah;
        const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
        totalWeightForGroup += actualWeight;
      });

      group.forEach((p, index) => {
        const costPriceYuan = p.priceYuan * p.quantity;
        const costPriceUah = costPriceYuan * p.exchangeRate;
        const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
        const actualVolume = p.volume || 0;
        const actualDensity = p.density || 0;
        
        let itemDeliveryUah = 0;
        if (group.length === 1) {
          itemDeliveryUah = totalDeliveryForGroup;
        } else if (totalWeightForGroup > 0) {
          itemDeliveryUah = totalDeliveryForGroup * (actualWeight / totalWeightForGroup);
        } else {
          itemDeliveryUah = totalDeliveryForGroup / group.length;
        }

        const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
        const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);

        const totalCostUah = costPriceUah + itemDeliveryUah;
        const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));

        const localDelivery = localDeliveries.find(ld => ld.id === p.localDeliveryId);
        const npTrack = localDelivery?.trackingNumber || '-';

        worksheet.addRow({
          name: p.name,
          comment: p.comment || '-',
          track: p.trackNumber,
          npTrack: npTrack,
          platform: p.platform,
          qty: p.quantity,
          priceYuan: p.priceYuan,
          rate: p.exchangeRate,
          sumUah: Math.round(costPriceUah),
          weight: actualWeight || '-',
          volume: actualVolume || '-',
          density: actualDensity || '-',
          deliveryChina: Math.round(deliveryChinaUah),
          deliveryNP: Math.round(deliveryNPUah),
          cost: Math.round(totalCostUah),
          markup: p.markup ? 'Індивідуальна' : `${priceListMargin}%`,
          sell: Math.round(sellingPriceUah),
          status: statusLabels[p.status] || p.status,
          batch: p.batchId || '-'
        });
        currentRow++;
      });

      if (mergeCells && group.length > 1 && !track.startsWith('NO_TRACK_')) {
        worksheet.mergeCells(`C${startRow}:C${currentRow - 1}`);
      }
    });

    // Style header
    const headerRow = worksheet.getRow(4);
    headerRow.height = 30;
    headerRow.font = { bold: true, color: { argb: 'FF000000' }, size: 11 };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFACC15' } // Yellow theme
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

    // Add borders to all cells
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber >= 4) {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
          };
          if (rowNumber > 4) {
            cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true, indent: 1 };
            if (cell.value && typeof cell.value === 'number') {
              cell.alignment = { vertical: 'middle', horizontal: 'center' };
            }
            if (rowNumber % 2 === 0) {
              cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF9F9F9' }
              };
            }
          }
        });
      }
    });

    // Generate buffer
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${fileName}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
    addNotification('Excel успішно згенеровано', 'success');
  };

  const exportToExcelWithPhotos = async (data: Purchase[], fileName: string, mergeCells: boolean = true) => {
    addNotification('Генерація Excel з фото...', 'info');
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Purchases');

    // Add Title Row
    const titleRow = worksheet.addRow(['CRM FORSAGE CHINA - ПРАЙС-ЛИСТ']);
    titleRow.font = { bold: true, size: 20, color: { argb: 'FF000000' } };
    worksheet.mergeCells('A1:Q1');
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    titleRow.height = 40;

    // Add Date Row
    const dateRow = worksheet.addRow([`Дата генерації: ${new Date().toLocaleString()}`]);
    dateRow.font = { italic: true, size: 11, color: { argb: 'FF666666' } };
    worksheet.mergeCells('A2:Q2');
    dateRow.alignment = { horizontal: 'center' };
    worksheet.addRow([]); // Empty row

    // Define columns
    worksheet.columns = [
      { header: 'Фото', key: 'photo', width: 25 },
      { header: 'Назва', key: 'name', width: 35 },
      { header: 'Опис / Коментар', key: 'comment', width: 35 },
      { header: 'Трек-номер', key: 'track', width: 25 },
      { header: 'ТТН НП', key: 'npTrack', width: 25 },
      { header: 'Платформа / Сайт', key: 'platform', width: 20 },
      { header: 'Кількість', key: 'qty', width: 10 },
      { header: 'Ціна (¥)', key: 'priceYuan', width: 15 },
      { header: 'Ціна (₴)', key: 'priceUah', width: 15 },
      { header: 'Сума (₴)', key: 'sumUah', width: 15 },
      { header: 'Вага (кг)', key: 'weight', width: 10 },
      { header: 'Об\'єм (м³)', key: 'volume', width: 15 },
      { header: 'Щільність (кг/м³)', key: 'density', width: 15 },
      { header: 'Доставка Китай (₴)', key: 'deliveryChina', width: 18 },
      { header: 'Доставка НП (₴)', key: 'deliveryNP', width: 18 },
      { header: 'Собівартість (₴)', key: 'cost', width: 15 },
      { header: 'Націнка (%)', key: 'markup', width: 12 },
      { header: 'Ціна продажу (₴)', key: 'sell', width: 15 },
    ];

    // Add data
    const groupedItems: { [key: string]: typeof data } = {};
    const trackOrder: string[] = [];
    
    data.forEach((p) => {
      const track = p.trackNumber || `NO_TRACK_${p.id}`;
      if (!groupedItems[track]) {
        groupedItems[track] = [];
        trackOrder.push(track);
      }
      groupedItems[track].push(p);
    });

    let currentRowIndex = 3; // 0-indexed for images, corresponds to row 5 in Excel (1: title, 2: date, 3: empty, 4: header)

    for (const track of trackOrder) {
      const group = groupedItems[track];
      const startRow = currentRowIndex + 2; // Excel row number (1-indexed, header is 4, so data starts at 5)

      
      let totalDeliveryForGroup = 0;
      let totalWeightForGroup = 0;

      group.forEach(p => {
        const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
        const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
        totalDeliveryForGroup += deliveryChinaUah + deliveryNPUah;
        const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
        totalWeightForGroup += actualWeight;
      });

      for (let i = 0; i < group.length; i++) {
        const p = group[i];
        const costPriceUah = (p.priceYuan * p.quantity) * p.exchangeRate;
        const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
        const actualVolume = p.volume || 0;
        const actualDensity = p.density || 0;
        
        let itemDeliveryUah = 0;
        if (group.length === 1) {
          itemDeliveryUah = totalDeliveryForGroup;
        } else if (totalWeightForGroup > 0) {
          itemDeliveryUah = totalDeliveryForGroup * (actualWeight / totalWeightForGroup);
        } else {
          itemDeliveryUah = totalDeliveryForGroup / group.length;
        }

        const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
        const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);

        const totalCostUah = costPriceUah + itemDeliveryUah;
        const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));

        const localDelivery = localDeliveries.find(ld => ld.id === p.localDeliveryId);
        const npTrack = localDelivery?.trackingNumber || '-';

        const row = worksheet.addRow({
          name: p.name,
          comment: p.comment || '-',
          track: p.trackNumber,
          npTrack: npTrack,
          platform: p.platform,
          qty: p.quantity,
          priceYuan: p.priceYuan,
          priceUah: (p.priceYuan * p.exchangeRate).toFixed(2),
          sumUah: costPriceUah.toFixed(0),
          weight: actualWeight || '-',
          volume: actualVolume || '-',
          density: actualDensity || '-',
          deliveryChina: Math.round(deliveryChinaUah),
          deliveryNP: Math.round(deliveryNPUah),
          cost: totalCostUah.toFixed(0),
          markup: p.markup ? 'Індивідуальна' : `${priceListMargin}%`,
          sell: sellingPriceUah.toFixed(0),
        });

        row.height = 100; // Set row height for images
        row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        if (p.photo) {
          try {
            if (p.photo.includes('base64,')) {
              const base64Data = p.photo.split('base64,')[1];
              const imageId = workbook.addImage({
                base64: base64Data,
                extension: 'jpeg',
              });
              worksheet.addImage(imageId, {
                tl: { col: 0, row: currentRowIndex + 1 },
                ext: { width: 120, height: 120 },
                editAs: 'oneCell'
              });
            }
          } catch (e) {
            console.error('Error adding image to Excel:', e);
          }
        }
        currentRowIndex++;
      }

      if (mergeCells && group.length > 1 && !track.startsWith('NO_TRACK_')) {
        worksheet.mergeCells(`D${startRow}:D${startRow + group.length - 1}`); // Track Number
      }
    }

    // Style header
    const headerRow = worksheet.getRow(4);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF003d2b' } // TURBO_GREEN
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height = 30;

    // Add borders to all cells
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // Generate buffer
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${fileName}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
    addNotification('Excel з фото успішно згенеровано', 'success');
  };

  const toggleSelectAll = (ids: string[]) => {
    if (selectedPurchaseIds.length === ids.length) {
      setSelectedPurchaseIds([]);
    } else {
      setSelectedPurchaseIds(ids);
    }
  };

  const toggleSelectOne = (id: string) => {
    if (selectedPurchaseIds.includes(id)) {
      setSelectedPurchaseIds(selectedPurchaseIds.filter(i => i !== id));
    } else {
      setSelectedPurchaseIds([...selectedPurchaseIds, id]);
    }
  };

  const handleBulkStatusChange = (newStatus: Purchase['status']) => {
    if (confirm(`Ви впевнені, що хочете змінити статус для ${selectedPurchaseIds.length} товарів на "${statusLabels[newStatus]}"?`)) {
      setPurchases(purchases.map(p => 
        selectedPurchaseIds.includes(p.id) ? { ...p, status: newStatus } : p
      ));
      setSelectedPurchaseIds([]);
      addNotification(`Статус успішно змінено для ${selectedPurchaseIds.length} товарів`, 'success');
    }
  };

  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
    addNotification(`Трек-номер ${text} скопійовано`, 'success');
  };

  useEffect(() => {
    const handlePaste = (event: ClipboardEvent) => {
      if (!showAddPurchaseModal) return;
      
      const items = event.clipboardData?.items;
      if (!items) return;

      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
          const blob = items[i].getAsFile();
          if (blob) {
            const reader = new FileReader();
            reader.onload = (e) => {
              setCropImage(e.target?.result as string);
            };
            reader.readAsDataURL(blob);
          }
        }
      }
    };

    window.addEventListener('paste', handlePaste);
    return () => window.removeEventListener('paste', handlePaste);
  }, [showAddPurchaseModal]);

  const filteredPurchases = useMemo(() => {
    return purchases.filter(p => {
      const matchesSearch = 
        (p.trackNumber || '').toLowerCase().includes(purchaseSearch.toLowerCase()) ||
        (p.name || '').toLowerCase().includes(purchaseSearch.toLowerCase()) ||
        (p.platform || '').toLowerCase().includes(purchaseSearch.toLowerCase());
      
      const matchesStatus = statusFilter === 'all' || p.status === statusFilter;
      
      return matchesSearch && matchesStatus;
    });
  }, [purchases, purchaseSearch, statusFilter]);

  const handleSale = () => {
    if (!showSaleModal.purchaseId) return;
    
    setPurchases(purchases.map(p => 
      p.id === showSaleModal.purchaseId 
        ? { 
            ...p, 
            status: 'sold', 
            sellingPrice: saleForm.sellingPrice,
            novaPoshtaCost: saleForm.novaPoshtaCost,
            ukraineDeliveryCost: saleForm.ukraineDeliveryCost,
            markup: saleForm.markup,
            markupValue: saleForm.markupValue,
            saleDestination: saleForm.saleDestination,
            soldDate: new Date().toISOString()
          } 
        : p
    ));
    
    setShowSaleModal({ show: false, purchaseId: null });
    setSaleForm({ sellingPrice: 0, novaPoshtaCost: 0, ukraineDeliveryCost: 0, markup: false, markupValue: 0, saleDestination: 'physical_store' });
    addNotification('Товар видано на магазин', 'success');
  };

  const handleDeleteSale = (id: string) => {
    if (confirm('Ви впевнені, що хочете видалити цей продаж? Товар повернеться на "Прибуло на мій склад".')) {
      setPurchases(purchases.map(p => p.id === id ? { 
        ...p, 
        status: 'my_warehouse', 
        sellingPrice: undefined, 
        soldDate: undefined,
        markup: undefined,
        markupValue: undefined,
        saleDestination: undefined,
        novaPoshtaCost: undefined,
        ukraineDeliveryCost: undefined
      } : p));
      addNotification('Продаж скасовано', 'info');
    }
  };

  const handleEditSale = (p: Purchase) => {
    setSaleForm({
      sellingPrice: p.sellingPrice || 0,
      novaPoshtaCost: p.novaPoshtaCost || 0,
      ukraineDeliveryCost: p.ukraineDeliveryCost || 0,
      markup: p.markup || false,
      markupValue: p.markupValue || 0,
      saleDestination: p.saleDestination || 'physical_store'
    });
    setShowSaleModal({ show: true, purchaseId: p.id });
  };
  const [editingPurchaseId, setEditingPurchaseId] = useState<string | null>(null);
  const [editingBatchId, setEditingBatchId] = useState<string | null>(null);

  const [isExportingPDF, setIsExportingPDF] = useState(false);

  const isIframe = useMemo(() => {
    try {
      return window.self !== window.top;
    } catch (e) {
      return true;
    }
  }, []);

  const handleExportTrackPDF = async (trackNumber: string) => {
    const element = document.getElementById('track-modal-content');
    if (!element) {
      addNotification('Помилка: Контейнер для друку не знайдено', 'error');
      return;
    }

    setIsExportingPDF(true);
    addNotification('Генерація PDF...', 'info');

    // Wait for React to re-render the DOM with isExportingPDF = true
    await new Promise(resolve => setTimeout(resolve, 100));

    try {
      const imgData = await domToPng(element, {
        scale: 1.5,
        backgroundColor: '#ffffff',
        width: 1400,
        filter: (node) => {
          if (node instanceof HTMLElement && node.classList.contains('no-print')) {
            return false;
          }
          return true;
        }
      });
      
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps = pdf.getImageProperties(imgData);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
      
      let heightLeft = pdfHeight;
      let position = 0;
      const pageHeight = pdf.internal.pageSize.getHeight();

      pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
      heightLeft -= pageHeight;

      while (heightLeft >= 0) {
        position = heightLeft - pdfHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
        heightLeft -= pageHeight;
      }

      pdf.save(`forsage-china-track-${trackNumber}-${new Date().toISOString().split('T')[0]}.pdf`);
      addNotification('PDF успішно згенеровано', 'success');
    } catch (err) {
      console.error('PDF export error:', err);
      addNotification('Помилка при генерації PDF. Спробуйте звичайний друк.', 'error');
    } finally {
      setIsExportingPDF(false);
    }
  };

  const handleExportPDF = async () => {
    const element = document.getElementById('price-list-container');
    if (!element) {
      addNotification('Помилка: Контейнер для друку не знайдено', 'error');
      return;
    }

    setIsExportingPDF(true);
    addNotification('Генерація PDF...', 'info');

    // Wait for React to re-render the DOM with isExportingPDF = true
    await new Promise(resolve => setTimeout(resolve, 100));

    try {
      // modern-screenshot handles images and modern CSS (like oklch) much better than html2canvas
      const imgData = await domToPng(element, {
        scale: 1.5,
        backgroundColor: '#ffffff',
        width: 1400,
        filter: (node) => {
          if (node instanceof HTMLElement && node.classList.contains('no-print')) {
            return false;
          }
          return true;
        }
      });
      
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps = pdf.getImageProperties(imgData);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
      
      let heightLeft = pdfHeight;
      let position = 0;
      const pageHeight = pdf.internal.pageSize.getHeight();

      pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
      heightLeft -= pageHeight;

      while (heightLeft >= 0) {
        position = heightLeft - pdfHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
        heightLeft -= pageHeight;
      }

      pdf.save(`forsage-china-price-list-${new Date().toISOString().split('T')[0]}.pdf`);
      addNotification('PDF успішно згенеровано', 'success');
    } catch (err) {
      console.error('PDF export error:', err);
      addNotification('Помилка при генерації PDF. Спробуйте звичайний друк.', 'error');
      window.focus();
      window.print();
    } finally {
      setIsExportingPDF(false);
    }
  };

  const handleAddToTrack = (purchase: Purchase) => {
    setPurchaseForm({
      platform: purchase.platform,
      name: '',
      link: '',
      priceYuan: 0,
      exchangeRate: purchase.exchangeRate,
      quantity: 1,
      trackNumber: purchase.trackNumber,
      photo: '',
      comment: '',
      size: '',
      width: 0,
      height: 0,
      length: 0,
      dimUnit: 'cm',
      weight: 0,
      weightUnit: 'kg',
      volume: 0,
      density: 0,
      isFabric: false,
      isPressed: false,
      isInsured: false,
      declaredValue: 0,
      shippingCost: 0,
      status: purchase.status,
      brand: '',
      radius: '',
      season: '',
      article: ''
    });
    setEditingPurchaseId(null);
    setShowAddPurchaseModal(true);
  };

  const handleEditPurchase = (purchase: Purchase) => {
    setPurchaseForm({
      platform: purchase.platform,
      name: purchase.name,
      link: purchase.link,
      priceYuan: purchase.priceYuan,
      exchangeRate: purchase.exchangeRate,
      quantity: purchase.quantity,
      trackNumber: purchase.trackNumber,
      photo: purchase.photo,
      comment: purchase.comment,
      size: purchase.size || '',
      width: purchase.width || 0,
      height: purchase.height || 0,
      length: purchase.length || 0,
      dimUnit: purchase.dimUnit || 'cm',
      weight: purchase.weight || 0,
      weightUnit: purchase.weightUnit || 'kg',
      volume: purchase.volume || 0,
      density: purchase.density || 0,
      isFabric: purchase.isFabric || false,
      isPressed: purchase.isPressed || false,
      isInsured: purchase.isInsured || false,
      declaredValue: purchase.declaredValue || 0,
      shippingCost: purchase.shippingCost || 0,
      status: purchase.status,
      brand: purchase.brand || '',
      radius: purchase.radius || '',
      season: purchase.season || '',
      article: purchase.article || ''
    });
    setEditingPurchaseId(purchase.id);
    setShowAddPurchaseModal(true);
    setSelectedTrackNumber(null);
  };

  const handleDeletePurchase = (id: string) => {
    setConfirmModal({
      show: true,
      title: 'Видалити запис?',
      message: 'Ви впевнені, що хочете видалити цей запис? Цю дію неможливо скасувати.',
      onConfirm: () => {
        setPurchases(purchases.filter(p => p.id !== id));
        addNotification('Запис видалено', 'error');
        setConfirmModal(null);
      }
    });
  };

  const handleEditBatch = (batch: Batch) => {
    setBatchForm({
      name: batch.name,
      shipmentDate: batch.shipmentDate,
      warehouse: batch.warehouse,
      deliveryType: batch.deliveryType
    });
    setEditingBatchId(batch.id);
    setShowCreateBatchModal(true);
  };

  const handleDeleteBatch = (id: string) => {
    setConfirmModal({
      show: true,
      title: 'Видалити партію?',
      message: 'Ви впевнені, що хочете видалити цю партію? Всі товари в ній повернуться до статусу "Відправлено з Китаю" без партії.',
      onConfirm: () => {
        setPurchases(purchases.map(p => p.batchId === id ? { ...p, status: 'shipped_to_ua', batchId: undefined } : p));
        setBatches(batches.filter(b => b.id !== id));
        addNotification('Партію видалено', 'error');
        setConfirmModal(null);
      }
    });
  };

  const handleBulkImport = () => {
    if (!bulkImportText.trim()) {
      addNotification('Вставте текст для імпорту', 'error');
      return;
    }

    const lines = bulkImportText.split('\n').map(l => l.trim()).filter(l => l !== '');
    const newPurchases: Purchase[] = [];
    let addedCount = 0;

    const trackIndices: number[] = [];
    const trackRegex = /^[A-Z0-9-]{10,}$/i;
    
    for (let i = 0; i < lines.length; i++) {
      if (trackRegex.test(lines[i]) && !/^(\d{2})\.(\d{2})\.(\d{4})/.test(lines[i])) {
        trackIndices.push(i);
      }
    }

    for (let j = 0; j < trackIndices.length; j++) {
      const startIndex = trackIndices[j];
      const endIndex = j < trackIndices.length - 1 ? trackIndices[j + 1] : lines.length;
      
      const block = lines.slice(startIndex, endIndex);
      const trackNumber = block[0];
      
      let arrivalDate = new Date().toISOString().split('T')[0];
      let weight = 0;
      let volume = 0;
      let density = 0;

      const dateLine = block.find(l => /^(\d{2})\.(\d{2})\.(\d{4})/.test(l));
      if (dateLine) {
        const [, d, m, y] = dateLine.match(/^(\d{2})\.(\d{2})\.(\d{4})/) || [];
        arrivalDate = `${y}-${m}-${d}`;
      }

      const numberLines = block.slice(1).filter(l => {
        if (l === '?') return true;
        if (/^(\d{2})\.(\d{2})\.(\d{4})/.test(l)) return false;
        if (trackRegex.test(l)) return false;
        return /^[\d]+([.,][\d]+)?/.test(l) && l.length < 20;
      });

      const values = numberLines.map(l => {
        if (l === '?') return 0;
        const match = l.match(/^([\d]+(?:[.,][\d]+)?)/);
        return match ? parseFloat(match[1].replace(',', '.')) : 0;
      });

      if (values.length >= 1) weight = values[0];
      if (values.length >= 2) volume = values[1];
      if (values.length >= 3) density = values[2];

      newPurchases.push({
        id: Math.random().toString(36).substr(2, 9),
        platform: 'Інше',
        name: `Імпорт ${trackNumber}`,
        link: '',
        priceYuan: 0,
        exchangeRate: cnyToUah,
        quantity: 1,
        trackNumber,
        photo: '',
        comment: '',
        dimUnit: 'cm',
        weight,
        weightUnit: 'kg',
        volume,
        density,
        isFabric: false,
        isPressed: false,
        isInsured: false,
        declaredValue: 0,
        shippingCost: 0,
        status: 'purchased',
        arrivalDate,
        createdAt: new Date().toISOString()
      });
      addedCount++;
    }

    if (newPurchases.length > 0) {
      setPurchases(prev => [...newPurchases, ...prev]);
      addNotification(`Успішно імпортовано ${addedCount} трек-номерів`, 'success');
      setBulkImportText('');
      setShowImportTracksModal(false);
    } else {
      addNotification('Не знайдено нових трек-номерів або невірний формат', 'error');
    }
  };

  const handleSavePurchase = (addAnother = false) => {
    if (editingPurchaseId) {
      setPurchases(purchases.map(p => p.id === editingPurchaseId ? { ...p, ...purchaseForm } : p));
      addNotification(`Запис "${purchaseForm.name}" оновлено`);
      setEditingPurchaseId(null);
    } else {
      const newPurchase: Purchase = {
        ...purchaseForm,
        id: Math.random().toString(36).substr(2, 9),
        createdAt: new Date().toISOString()
      };
      setPurchases([newPurchase, ...purchases]);
      addNotification(`Товар "${newPurchase.name}" додано до бази`);
    }
    
    if (addAnother && !editingPurchaseId) {
      setPurchaseForm({
        ...purchaseForm,
        name: '',
        link: '',
        priceYuan: 0,
        quantity: 1,
        trackNumber: '',
        photo: '',
        comment: '',
        size: '',
        width: 0,
        height: 0,
        length: 0,
        dimUnit: 'cm',
        weight: 0,
        weightUnit: 'kg',
        volume: 0,
        density: 0,
        isFabric: false,
        isPressed: false,
        isInsured: false,
        declaredValue: 0,
        shippingCost: 0,
        status: 'purchased',
        brand: '',
        radius: '',
        season: '',
        article: ''
      });
    } else {
      setShowAddPurchaseModal(false);
      setPurchaseForm({
        platform: 'Taobao',
        name: '',
        link: '',
        priceYuan: 0,
        exchangeRate: 5.5,
        quantity: 1,
        trackNumber: '',
        photo: '',
        comment: '',
        size: '',
        width: 0,
        height: 0,
        length: 0,
        dimUnit: 'cm',
        weight: 0,
        weightUnit: 'kg',
        volume: 0,
        density: 0,
        isFabric: false,
        isPressed: false,
        isInsured: false,
        declaredValue: 0,
        shippingCost: 0,
        status: 'purchased',
        brand: '',
        radius: '',
        season: '',
        article: ''
      });
    }
  };

  const handleCompleteLocalDelivery = (id: string) => {
    setLocalDeliveries(localDeliveries.map(ld => 
      ld.id === id ? { ...ld, status: 'received' } : ld
    ));
    
    setPurchases(purchases.map(p => 
      p.localDeliveryId === id ? { ...p, status: 'my_warehouse' } : p
    ));
    addNotification(`Відправку отримано, товари переміщено на "Прибуло на мій склад"`);
  };

  const handleCreateLocalDelivery = () => {
    if (editingLocalDeliveryId) {
      setLocalDeliveries(localDeliveries.map(ld => ld.id === editingLocalDeliveryId ? { ...ld, ...localDeliveryForm } : ld));
      addNotification(`Відправку "${localDeliveryForm.name}" оновлено`);
      setEditingLocalDeliveryId(null);
      setShowCreateLocalDeliveryModal(false);
      setLocalDeliveryForm({
        name: '',
        type: 'novaposhta',
        trackingNumber: '',
        cost: 0
      });
      return;
    }

    const itemsToDeliver = purchases.filter(p => p.status === 'arrived_ua' && !p.localDeliveryId);
    if (itemsToDeliver.length === 0) {
      alert('Немає товарів зі статусом "На складі Україна Київ" без відправки');
      return;
    }

    const newLocalDelivery: LocalDeliveryGroup = {
      id: Math.random().toString(36).substr(2, 9),
      name: localDeliveryForm.name || `Відправка-${new Date().toISOString().split('T')[0]}`,
      type: localDeliveryForm.type,
      trackingNumber: localDeliveryForm.trackingNumber,
      cost: localDeliveryForm.cost,
      status: 'pending',
      itemIds: itemsToDeliver.map(p => p.id),
      createdAt: new Date().toISOString()
    };

    setLocalDeliveries([newLocalDelivery, ...localDeliveries]);
    
    setPurchases(purchases.map(p => 
      p.status === 'arrived_ua' && !p.localDeliveryId
        ? { ...p, status: 'local_delivery', localDeliveryId: newLocalDelivery.id } 
        : p
    ));

    setShowCreateLocalDeliveryModal(false);
    addNotification(`Створено відправку "${newLocalDelivery.name}"`, 'success');
  };

  const handleEditLocalDelivery = (ld: LocalDeliveryGroup) => {
    setLocalDeliveryForm({
      name: ld.name,
      type: ld.type,
      trackingNumber: ld.trackingNumber || '',
      cost: ld.cost
    });
    setEditingLocalDeliveryId(ld.id);
    setShowCreateLocalDeliveryModal(true);
  };

  const handleDeleteLocalDelivery = (id: string) => {
    setConfirmModal({
      show: true,
      title: 'Видалити відправку?',
      message: 'Ви впевнені, що хочете видалити цю відправку? Всі товари в ній повернуться до статусу "На складі Україна Київ".',
      onConfirm: () => {
        setPurchases(purchases.map(p => p.localDeliveryId === id ? { ...p, status: 'arrived_ua', localDeliveryId: undefined } : p));
        setLocalDeliveries(localDeliveries.filter(ld => ld.id !== id));
        addNotification('Відправку видалено', 'error');
        setConfirmModal(null);
      }
    });
  };

  const getLocalDeliveryCost = (item: Purchase): number => {
    if (!item.localDeliveryId) return 0;
    const localDelivery = localDeliveries.find(ld => ld.id === item.localDeliveryId);
    if (!localDelivery) return 0;
    
    const groupItems = purchases.filter(p => localDelivery.itemIds.includes(p.id));
    if (groupItems.length === 0) return 0;
    
    const totalGroupWeight = groupItems.reduce((sum, p) => {
      const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
      return sum + actualWeight;
    }, 0);
    if (totalGroupWeight > 0) {
      const itemActualWeight = item.weightUnit === 'g' ? (item.weight || 0) / 1000 : (item.weight || 0);
      return localDelivery.cost * (itemActualWeight / totalGroupWeight);
    }
    return localDelivery.cost / groupItems.length;
  };

  const handleCreateBatch = () => {
    if (editingBatchId) {
      setBatches(batches.map(b => b.id === editingBatchId ? { ...b, ...batchForm } : b));
      addNotification(`Партію "${batchForm.name}" оновлено`);
      setEditingBatchId(null);
      setShowCreateBatchModal(false);
      setBatchForm({
        name: '',
        shipmentDate: new Date().toISOString().split('T')[0],
        warehouse: 'Guangzhou',
        deliveryType: 'sea'
      });
      return;
    }

    const itemsToBatch = purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId);
    if (itemsToBatch.length === 0) {
      alert('Немає товарів зі статусом "Відправлено з Китаю" без партії');
      return;
    }

    const newBatch: Batch = {
      id: Math.random().toString(36).substr(2, 9),
      name: batchForm.name || `BATCH-${new Date().toISOString().split('T')[0]}`,
      shipmentDate: batchForm.shipmentDate,
      warehouse: batchForm.warehouse,
      deliveryType: batchForm.deliveryType,
      status: 'shipped',
      itemIds: itemsToBatch.map(p => p.id),
      createdAt: new Date().toISOString()
    };

    setBatches([newBatch, ...batches]);
    
    // Update purchase statuses
    setPurchases(purchases.map(p => 
      p.status === 'shipped_to_ua' && !p.batchId
        ? { ...p, batchId: newBatch.id } 
        : p
    ));

    setShowCreateBatchModal(false);
    setBatchForm({
      name: '',
      shipmentDate: new Date().toISOString().split('T')[0],
      warehouse: 'Guangzhou',
      deliveryType: 'sea'
    });
    addNotification(`Партію "${newBatch.name}" успішно створено`);
  };

  const handleBatchArrived = (batchId: string) => {
    setBatches(batches.map(b => 
      b.id === batchId ? { ...b, status: 'arrived_ua' } : b
    ));
    
    setPurchases(purchases.map(p => 
      p.batchId === batchId ? { ...p, status: 'arrived_ua' } : p
    ));
    addNotification(`Партія прибула на склад в Україну`);
  };

  const handleSaveCosts = () => {
    if (!showCostModal.batchId) return;
    
    const selectedTariff = tariffs.find(t => t.id === costForm.tariffId) || tariffs[0];
    
    let shippingCost = 0;
    if (selectedTariff.pricePerKg) {
      const volumetricWeight = (costForm.volume * 1000000) / (selectedTariff.volumetricFactor || 5000);
      shippingCost = selectedTariff.pricePerKg * Math.max(costForm.totalWeight, volumetricWeight);
    } else if (selectedTariff.densityTiers) {
      const density = costForm.volume > 0 ? costForm.totalWeight / costForm.volume : 0;
      const tier = selectedTariff.densityTiers.find(t => density >= t.min && (t.max === null || density < t.max));
      if (tier) {
        shippingCost = tier.unit === 'm3' ? tier.price * costForm.volume : tier.price * costForm.totalWeight;
      }
    }

    const insurance = (costForm.isInsured && costForm.declaredValue) ? Math.max(1, costForm.declaredValue * 0.02) : 0;
    const fabricSurcharge = costForm.isFabric ? costForm.totalWeight * 0.2 : 0;
    const pressingCost = costForm.isPressed ? 5 : 0;
    const localDelivery = (selectedTariff.localDeliveryPrice || 0) * costForm.totalWeight;
    
    const totalDeliveryCost = costForm.deliveryCost || (shippingCost + insurance + localDelivery + fabricSurcharge + pressingCost);
    const pricePerKg = costForm.totalWeight > 0 ? totalDeliveryCost / costForm.totalWeight : 0;
    
    setBatches(batches.map(b => 
      b.id === showCostModal.batchId 
        ? { ...b, totalWeight: costForm.totalWeight, deliveryCost: totalDeliveryCost, pricePerKg } 
        : b
    ));

    // Update individual items in the batch
    setPurchases(purchases.map(p => {
      if (p.batchId === showCostModal.batchId && p.weight) {
        const actualWeight = p.weightUnit === 'g' ? p.weight / 1000 : p.weight;
        return { ...p, deliveryCostPerItem: actualWeight * pricePerKg };
      }
      return p;
    }));

    setShowCostModal({show: false});
    setCostForm({ 
      totalWeight: 0, 
      deliveryCost: 0, 
      volume: 0, 
      declaredValue: 0, 
      isInsured: false, 
      isFabric: false, 
      isPressed: false, 
      tariffId: 'turbo-sea' 
    });
  };

  const handleTrackNumberChange = (value: string) => {
    setPurchaseForm({ ...purchaseForm, trackNumber: value });
  };

  const crmModules: CRMModuleConfig[] = [
    { id: 'dashboard', title: 'Дашборд', icon: BarChart3, description: 'Аналітика та статистика' },
    { id: 'purchases', title: 'Закупки', icon: ShoppingCart, description: 'Управління замовленнями та постачальниками' },
    { id: 'china_warehouse', title: 'Склад Китай', icon: Warehouse, description: 'Прийом та обробка вантажів у Китаї' },
    { id: 'consolidation', title: 'В дорозі з Китаю', icon: Combine, description: 'Формування партій та логістика в Україну' },
    { id: 'ua_warehouse', title: 'На складі Україна Київ', icon: Package, description: 'Прийом партій в Києві' },
    { id: 'local_delivery', title: 'Нова Пошта', icon: Truck, description: 'Доставка по Україні та самовивіз' },
    { id: 'my_warehouse', title: 'Прибуло на мій склад', icon: Store, description: 'Наявність товарів та прайс-листи' },
    { id: 'issue_to_store', title: 'Видача на магазин', icon: Store, description: 'Продажі та видача товарів' },
    { id: 'pinduoduo', title: 'Pinduoduo App', icon: Smartphone, description: 'Мобільний доступ до маркетплейсу' },
    { id: 'settings', title: 'Налаштування', icon: Settings, description: 'Курси валют та системні параметри' },
  ];

  return (
    <div className="min-h-screen bg-[#f8f9fa] text-[#1a1a1a] font-sans selection:bg-yellow-100">
      
      {/* Main Header */}
      <header className="bg-black text-white shadow-md sticky top-0 z-50 no-print border-b border-yellow-400/20">
        <div className="max-w-7xl mx-auto px-4 h-7 flex items-center justify-between">
          <div className="flex items-center gap-2 group cursor-pointer" onClick={() => setView('crm')}>
            <div className="w-6 h-6 bg-yellow-400 rounded-lg flex items-center justify-center rotate-3 group-hover:rotate-12 transition-transform shadow-lg shadow-yellow-400/20">
              <Zap className="w-3 h-3 text-black fill-current" />
            </div>
            <div className="flex flex-col -space-y-1">
              <h1 className="text-base font-black tracking-tighter leading-none flex items-center gap-1">
                FORSAGE <span className="text-yellow-400">CHINA</span>
              </h1>
              <p className="text-[6px] font-bold text-yellow-400/60 uppercase tracking-[0.2em] leading-none mt-0.5">CRM FORSAGE CHINA</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <button 
              onClick={() => setView('calculator')}
              className={cn(
                "text-[10px] font-black uppercase tracking-widest transition-colors",
                view === 'calculator' ? "text-yellow-400" : "text-white/60 hover:text-white"
              )}
            >
              Калькулятор
            </button>
            <button 
              onClick={() => setView('crm')}
              className={cn(
                "text-[10px] font-black uppercase tracking-widest transition-colors",
                view === 'crm' ? "text-yellow-400" : "text-white/60 hover:text-white"
              )}
            >
              CRM
            </button>
          </div>
        </div>
      </header>


      {view === 'calculator' ? (
        <>
          {/* Sub Navigation for Calculator */}
          <div className="bg-black border-t border-yellow-400/10 py-1.5">
            <div className="max-w-7xl mx-auto px-6 flex justify-center">
              <div className="flex bg-white/5 p-0.5 rounded-lg border border-white/10">
                <button 
                  onClick={() => setActiveTab('international')}
                  className={cn(
                    "px-4 py-1.5 rounded-md text-[9px] font-black uppercase tracking-widest transition-all flex items-center gap-1.5",
                    activeTab === 'international' ? "bg-yellow-400 text-black shadow-lg" : "text-white/60 hover:text-white"
                  )}
                >
                  <Globe className="w-2.5 h-2.5" />
                  Китай
                </button>
                <button 
                  onClick={() => setActiveTab('novaposhta')}
                  className={cn(
                    "px-4 py-1.5 rounded-md text-[9px] font-black uppercase tracking-widest transition-all flex items-center gap-1.5",
                    activeTab === 'novaposhta' ? "bg-yellow-400 text-black shadow-lg" : "text-white/60 hover:text-white"
                  )}
                >
                  <Truck className="w-2.5 h-2.5" />
                  Нова Пошта
                </button>
                <button 
                  onClick={() => setActiveTab('transfer')}
                  className={cn(
                    "px-4 py-1.5 rounded-md text-[9px] font-black uppercase tracking-widest transition-all flex items-center gap-1.5",
                    activeTab === 'transfer' ? "bg-yellow-400 text-black shadow-lg" : "text-white/60 hover:text-white"
                  )}
                >
                  <DollarSign className="w-2.5 h-2.5" />
                  Переказ
                </button>
              </div>
            </div>
          </div>

          {/* Hero Section */}
          <div className="bg-black py-16 text-center text-white relative overflow-hidden">
        <div className="absolute inset-0 opacity-20 pointer-events-none">
          <img 
            src="https://picsum.photos/seed/logistics/1920/1080?blur=2" 
            alt="Background" 
            className="w-full h-full object-cover"
            referrerPolicy="no-referrer"
          />
          <div className="absolute inset-0 bg-gradient-to-b from-black via-black/80 to-black" />
        </div>
        <div className="max-w-4xl mx-auto px-6 relative z-10">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6 }}
          >
            <h2 className="text-4xl md:text-6xl font-display font-black mb-4 tracking-tight leading-tight uppercase italic">
              Розрахунок <span className="text-yellow-400">вартості</span> доставки
            </h2>
            <p className="text-white/60 text-lg max-w-2xl mx-auto font-medium">
              Професійний калькулятор для точного прорахунку логістики з Китаю.
            </p>
          </motion.div>
        </div>
      </div>

      <main className="max-w-7xl mx-auto px-6 py-16 grid grid-cols-1 lg:grid-cols-12 gap-12 -mt-16">
        
        {/* Left Column: Inputs */}
        <div className="lg:col-span-7 space-y-10">
          
          {activeTab === 'international' && (
            <>
              {/* Manual Input Section */}
              <section className="bg-white rounded-2xl p-10 shadow-2xl border-b-8 border-yellow-400 relative overflow-hidden">
                <div className="absolute top-0 right-0 w-40 h-40 bg-gray-50 -mr-20 -mt-20 rounded-full" />
                
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-10 relative z-10">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-black rounded-xl flex items-center justify-center text-white shadow-lg shadow-gray-900/20">
                      <Package className="w-7 h-7" />
                    </div>
                    <div>
                      <h2 className="font-display font-black text-3xl uppercase tracking-tight text-black">Параметри вантажу</h2>
                      <p className="text-xs text-gray-400 font-bold uppercase tracking-widest mt-1">Вкажіть дані вашої посилки</p>
                    </div>
                  </div>

                  <div className="flex bg-gray-100 p-1 rounded-xl self-start sm:self-center">
                    <button 
                      onClick={() => setInputMethod('dims')}
                      className={cn(
                        "px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all",
                        inputMethod === 'dims' ? "bg-black text-white shadow-md" : "text-gray-400 hover:text-gray-600"
                      )}
                    >
                      Розміри
                    </button>
                    <button 
                      onClick={() => setInputMethod('density')}
                      className={cn(
                        "px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all",
                        inputMethod === 'density' ? "bg-black text-white shadow-md" : "text-gray-400 hover:text-gray-600"
                      )}
                    >
                      Щільність
                    </button>
                  </div>
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-10 relative z-10">
                  <div className="space-y-4">
                    <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                      <Scale className="w-4 h-4 text-yellow-500" /> Фактична вага (кг)
                    </label>
                    <div className="relative">
                      <input 
                        type="number" 
                        value={parcel.weight || ''} 
                        onChange={(e) => setParcel(p => ({ ...p, weight: parseFloat(e.target.value) || 0 }))}
                        placeholder="0.00"
                        className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-5 focus:border-black focus:bg-white outline-none font-black text-2xl transition-all placeholder:text-gray-200"
                      />
                      <div className="absolute right-6 top-1/2 -translate-y-1/2 text-gray-300 font-black">KG</div>
                    </div>
                  </div>

                  {inputMethod === 'dims' ? (
                    <div className="space-y-4">
                      <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                        <Layers className="w-4 h-4 text-yellow-500" /> Об'єм вантажу (м³)
                      </label>
                      <div className="relative">
                        <input 
                          type="number" 
                          value={parcel.volume || ''} 
                          onChange={(e) => setParcel(p => ({ ...p, volume: parseFloat(e.target.value) || 0 }))}
                          placeholder="0.000"
                          step="0.001"
                          className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-5 focus:border-black focus:bg-white outline-none font-black text-2xl transition-all placeholder:text-gray-200"
                        />
                        <div className="absolute right-6 top-1/2 -translate-y-1/2 text-gray-300 font-black">M³</div>
                      </div>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                        <Maximize className="w-4 h-4 text-yellow-500" /> Щільність (кг/м³)
                      </label>
                      <div className="relative">
                        <input 
                          type="number" 
                          value={parcel.density || ''} 
                          onChange={(e) => setParcel(p => ({ ...p, density: parseFloat(e.target.value) || 0 }))}
                          placeholder="0"
                          className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-5 focus:border-black focus:bg-white outline-none font-black text-2xl transition-all placeholder:text-gray-200"
                        />
                        <div className="absolute right-6 top-1/2 -translate-y-1/2 text-gray-300 font-black">KG/M³</div>
                      </div>
                    </div>
                  )}
                  
                  {inputMethod === 'dims' && (
                    <div className="space-y-4 sm:col-span-2">
                      <div className="flex justify-between items-center">
                        <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                          <Maximize className="w-4 h-4 text-yellow-500" /> Габарити (см)
                        </label>
                        <span className="text-[10px] text-gray-300 font-black uppercase tracking-widest">Об'єм розраховується автоматично</span>
                      </div>
                      <div className="grid grid-cols-3 gap-6">
                        <div className="relative">
                          <input 
                            type="number" 
                            value={parcel.length || ''} 
                            onChange={(e) => setParcel(p => ({ ...p, length: parseFloat(e.target.value) || 0 }))}
                            placeholder="Д"
                            className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-4 py-5 focus:border-black focus:bg-white outline-none text-center font-black text-xl placeholder:text-gray-200"
                          />
                          <span className="absolute bottom-1 left-1/2 -translate-x-1/2 text-[8px] text-gray-300 font-black uppercase">Довжина</span>
                        </div>
                        <div className="relative">
                          <input 
                            type="number" 
                            value={parcel.width || ''} 
                            onChange={(e) => setParcel(p => ({ ...p, width: parseFloat(e.target.value) || 0 }))}
                            placeholder="Ш"
                            className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-4 py-5 focus:border-black focus:bg-white outline-none text-center font-black text-xl placeholder:text-gray-200"
                          />
                          <span className="absolute bottom-1 left-1/2 -translate-x-1/2 text-[8px] text-gray-300 font-black uppercase">Ширина</span>
                        </div>
                        <div className="relative">
                          <input 
                            type="number" 
                            value={parcel.height || ''} 
                            onChange={(e) => setParcel(p => ({ ...p, height: parseFloat(e.target.value) || 0 }))}
                            placeholder="В"
                            className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-4 py-5 focus:border-black focus:bg-white outline-none text-center font-black text-xl placeholder:text-gray-200"
                          />
                          <span className="absolute bottom-1 left-1/2 -translate-x-1/2 text-[8px] text-gray-300 font-black uppercase">Висота</span>
                        </div>
                      </div>
                    </div>
                  )}

                  <div className="space-y-4 sm:col-span-2 pt-4">
                    <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                      <ShieldCheck className="w-4 h-4 text-yellow-500" /> Оголошена вартість ($)
                    </label>
                    <div className="relative">
                      <input 
                        type="number" 
                        value={parcel.declaredValue || ''} 
                        onChange={(e) => setParcel(p => ({ ...p, declaredValue: parseFloat(e.target.value) || 0 }))}
                        placeholder="0.00"
                        className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-5 focus:border-black focus:bg-white outline-none font-black text-2xl transition-all placeholder:text-gray-200"
                      />
                      <div className="absolute right-6 top-1/2 -translate-y-1/2 text-yellow-500 font-black">USD</div>
                    </div>
                  </div>

                  <div className="space-y-4 sm:col-span-2 pt-6 border-t border-gray-50">
                    <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em] flex items-center gap-2">
                      Додаткові параметри
                    </label>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                      <button 
                        onClick={() => setParcel(p => ({ ...p, isInsured: !p.isInsured }))}
                        className={cn(
                          "flex items-center justify-between p-5 rounded-xl border-2 transition-all group",
                          parcel.isInsured ? "border-black bg-black/5" : "border-gray-100 hover:border-gray-200"
                        )}
                      >
                        <div className="flex items-center gap-3">
                          <ShieldCheck className={cn("w-5 h-5 transition-colors", parcel.isInsured ? "text-yellow-500" : "text-gray-400 group-hover:text-yellow-500")} />
                          <span className={cn("text-[11px] font-black uppercase tracking-widest", parcel.isInsured ? "text-black" : "text-gray-400")}>Страхування (2%)</span>
                        </div>
                        <div className={cn("w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all", parcel.isInsured ? "border-black bg-black" : "border-gray-200")}>
                          {parcel.isInsured && <div className="w-1.5 h-1.5 bg-white rounded-full" />}
                        </div>
                      </button>
                      <button 
                        onClick={() => setParcel(p => ({ ...p, isFabric: !p.isFabric }))}
                        className={cn(
                          "flex items-center justify-between p-5 rounded-xl border-2 transition-all group",
                          parcel.isFabric ? "border-black bg-black/5" : "border-gray-100 hover:border-gray-200"
                        )}
                      >
                        <div className="flex items-center gap-3">
                          <Scissors className={cn("w-5 h-5 transition-colors", parcel.isFabric ? "text-yellow-500" : "text-gray-400 group-hover:text-yellow-500")} />
                          <span className={cn("text-[11px] font-black uppercase tracking-widest", parcel.isFabric ? "text-black" : "text-gray-400")}>Тканина (+0.2$/кг)</span>
                        </div>
                        <div className={cn("w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all", parcel.isFabric ? "border-black bg-black" : "border-gray-200")}>
                          {parcel.isFabric && <div className="w-1.5 h-1.5 bg-white rounded-full" />}
                        </div>
                      </button>
                      <button 
                        onClick={() => setParcel(p => ({ ...p, isPressed: !p.isPressed }))}
                        className={cn(
                          "flex items-center justify-between p-5 rounded-xl border-2 transition-all group",
                          parcel.isPressed ? "border-black bg-black/5" : "border-gray-100 hover:border-gray-200"
                        )}
                      >
                        <div className="flex items-center gap-3">
                          <Minimize2 className={cn("w-5 h-5 transition-colors", parcel.isPressed ? "text-yellow-500" : "text-gray-400 group-hover:text-yellow-500")} />
                          <span className={cn("text-[11px] font-black uppercase tracking-widest", parcel.isPressed ? "text-black" : "text-gray-400")}>Пресування (+$5)</span>
                        </div>
                        <div className={cn("w-5 h-5 rounded-full border-2 flex items-center justify-center transition-all", parcel.isPressed ? "border-black bg-black" : "border-gray-200")}>
                          {parcel.isPressed && <div className="w-1.5 h-1.5 bg-white rounded-full" />}
                        </div>
                      </button>
                    </div>
                  </div>
                </div>
              </section>

              {/* Tariffs Section */}
              <section className="bg-white rounded-2xl p-10 shadow-2xl">
                <div className="flex items-center gap-4 mb-10">
                  <div className="w-12 h-12 bg-[#f8f9fa] rounded-xl flex items-center justify-center text-black border border-gray-100">
                    <Truck className="w-7 h-7" />
                  </div>
                  <div>
                    <h2 className="font-display font-black text-3xl uppercase tracking-tight text-black">Спосіб доставки</h2>
                    <p className="text-xs text-gray-400 font-bold uppercase tracking-widest mt-1">Оберіть оптимальний варіант</p>
                  </div>
                </div>
                
                <div className="grid grid-cols-1 gap-6">
                  {tariffs.map((tariff) => (
                    <button
                      key={tariff.id}
                      onClick={() => setSelectedTariffId(tariff.id)}
                      className={cn(
                        "relative p-8 rounded-2xl border-2 text-left transition-all duration-300 flex items-center justify-between group overflow-hidden",
                        selectedTariffId === tariff.id 
                          ? "border-[#003d2b] bg-black/5 ring-4 ring-black/5" 
                          : "border-gray-50 hover:border-gray-200 hover:bg-gray-50/50"
                      )}
                    >
                      <div className="flex items-center gap-8 relative z-10">
                        <div className={cn(
                          "w-16 h-16 rounded-2xl flex items-center justify-center transition-all duration-300 shadow-lg",
                          selectedTariffId === tariff.id 
                            ? "bg-[#003d2b] text-white scale-110" 
                            : "bg-white text-gray-400 border border-gray-100 group-hover:text-[#003d2b]"
                        )}>
                          {getTariffIcon(tariff.iconName)}
                        </div>
                        <div>
                          <h3 className={cn(
                            "font-display font-black text-xl uppercase tracking-tight transition-colors",
                            selectedTariffId === tariff.id ? "text-[#003d2b]" : "text-gray-900"
                          )}>{tariff.name}</h3>
                          <div className="flex items-center gap-3 mt-1">
                            <span className="text-xs text-black font-black uppercase tracking-widest">{tariff.deliveryDays}</span>
                            <span className="text-gray-300">|</span>
                            <span className="text-xs text-gray-400 font-bold">{tariff.description}</span>
                          </div>
                        </div>
                      </div>
                    </button>
                  ))}
                </div>
              </section>
            </>
          )}

          {activeTab === 'novaposhta' && (
            <section className="bg-white rounded-2xl p-10 shadow-2xl border-b-8 border-[#facc15]">
              <div className="flex items-center gap-4 mb-10">
                <div className="w-12 h-12 bg-black rounded-xl flex items-center justify-center text-white">
                  <Truck className="w-7 h-7" />
                </div>
                <div>
                  <h2 className="font-display font-black text-3xl uppercase tracking-tight text-black">Нова Пошта</h2>
                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest mt-1">Доставка по Україні</p>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-8">
                <div className="space-y-4">
                  <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Вага (кг)</label>
                  <input 
                    type="number" 
                    value={npData.weight || ''} 
                    onChange={(e) => setNpData(p => ({ ...p, weight: parseFloat(e.target.value) || 0 }))}
                    className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-4 font-black text-xl"
                  />
                </div>
                <div className="space-y-4">
                  <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Напрямок</label>
                  <select 
                    value={npData.destination}
                    onChange={(e) => setNpData(p => ({ ...p, destination: e.target.value as any }))}
                    className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-4 font-black text-lg appearance-none"
                  >
                    <option value="city">По місту</option>
                    <option value="region">По області</option>
                    <option value="ukraine">По Україні</option>
                  </select>
                </div>
                <div className="sm:col-span-2 grid grid-cols-3 gap-4">
                  <div className="space-y-2">
                    <label className="text-[9px] font-black text-gray-400 uppercase">Д (см)</label>
                    <input type="number" value={npData.length || ''} onChange={(e) => setNpData(p => ({ ...p, length: parseFloat(e.target.value) || 0 }))} className="w-full bg-gray-50 border-2 border-gray-100 rounded-lg px-4 py-3 text-center font-bold" />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[9px] font-black text-gray-400 uppercase">Ш (см)</label>
                    <input type="number" value={npData.width || ''} onChange={(e) => setNpData(p => ({ ...p, width: parseFloat(e.target.value) || 0 }))} className="w-full bg-gray-50 border-2 border-gray-100 rounded-lg px-4 py-3 text-center font-bold" />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[9px] font-black text-gray-400 uppercase">В (см)</label>
                    <input type="number" value={npData.height || ''} onChange={(e) => setNpData(p => ({ ...p, height: parseFloat(e.target.value) || 0 }))} className="w-full bg-gray-50 border-2 border-gray-100 rounded-lg px-4 py-3 text-center font-bold" />
                  </div>
                </div>
                <div className="sm:col-span-2 space-y-4">
                  <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Оголошена вартість (грн)</label>
                  <input 
                    type="number" 
                    value={npData.declaredValue || ''} 
                    onChange={(e) => setNpData(p => ({ ...p, declaredValue: parseFloat(e.target.value) || 0 }))}
                    className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-4 font-black text-xl"
                  />
                </div>
              </div>
            </section>
          )}

          {activeTab === 'transfer' && (
            <section className="bg-white rounded-2xl p-10 shadow-2xl border-b-8 border-[#facc15]">
              <div className="flex items-center gap-4 mb-10">
                <div className="w-12 h-12 bg-black rounded-xl flex items-center justify-center text-white">
                  <Globe className="w-7 h-7" />
                </div>
                <div>
                  <h2 className="font-display font-black text-3xl uppercase tracking-tight text-black">Грошові перекази</h2>
                  <p className="text-xs text-gray-400 font-bold uppercase tracking-widest mt-1">По Україні</p>
                </div>
              </div>

              <div className="space-y-8">
                <div className="space-y-4">
                  <label className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Сума переказу (грн)</label>
                  <input 
                    type="number" 
                    value={transferData.amount || ''} 
                    onChange={(e) => setTransferData(p => ({ ...p, amount: parseFloat(e.target.value) || 0 }))}
                    className="w-full bg-gray-50 border-2 border-gray-100 rounded-xl px-6 py-5 font-black text-3xl text-black"
                    placeholder="0.00"
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <button 
                    onClick={() => setTransferData(p => ({ ...p, method: 'card' }))}
                    className={cn(
                      "p-6 rounded-xl border-2 font-black uppercase tracking-widest text-xs transition-all",
                      transferData.method === 'card' ? "border-black bg-black text-white" : "border-gray-100 text-gray-400 hover:border-gray-200"
                    )}
                  >
                    На карту
                  </button>
                  <button 
                    onClick={() => setTransferData(p => ({ ...p, method: 'cash' }))}
                    className={cn(
                      "p-6 rounded-xl border-2 font-black uppercase tracking-widest text-xs transition-all",
                      transferData.method === 'cash' ? "border-black bg-black text-white" : "border-gray-100 text-gray-400 hover:border-gray-200"
                    )}
                  >
                    Готівка
                  </button>
                </div>
              </div>
            </section>
          )}

          {/* Density Guide */}
          <section className="bg-black rounded-2xl p-10 text-white shadow-2xl relative overflow-hidden">
            <div className="absolute bottom-0 right-0 w-64 h-64 bg-white/5 -mb-32 -mr-32 rounded-full blur-3xl" />
            <div className="relative z-10">
              <h3 className="font-display font-black text-2xl uppercase mb-6 flex items-center gap-3">
                <Info className="w-6 h-6 text-[#facc15]" />
                Як розраховується щільність?
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                <div className="space-y-2">
                  <div className="text-[#facc15] font-black text-3xl">01</div>
                  <p className="text-sm text-white/70 font-medium">Вимірюємо вагу вантажу в кілограмах (кг).</p>
                </div>
                <div className="space-y-2">
                  <div className="text-[#facc15] font-black text-3xl">02</div>
                  <p className="text-sm text-white/70 font-medium">Вимірюємо об'єм вантажу в кубічних метрах (м³).</p>
                </div>
                <div className="space-y-2">
                  <div className="text-[#facc15] font-black text-3xl">03</div>
                  <p className="text-sm text-white/70 font-medium">Ділимо вагу на об'єм. Отримуємо кг/м³.</p>
                </div>
              </div>
              <div className="mt-10 p-6 bg-white/5 rounded-xl border border-white/10 flex items-center gap-6">
                <div className="text-4xl font-black text-[#facc15] italic">!</div>
                <p className="text-sm font-medium leading-relaxed">
                  Чим вища щільність вашого вантажу, тим вигідніша ставка за кілограм. Для морської доставки це ключовий показник вартості.
                </p>
              </div>
            </div>
          </section>
        </div>

        {/* Right Column: Results */}
        <div className="lg:col-span-5">
          <div className="sticky top-32 space-y-10">
            
            {/* Summary Card */}
            <motion.div 
              key={activeTab}
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className="bg-white rounded-3xl p-12 shadow-2xl relative overflow-hidden border-2 border-black"
            >
              <div className="absolute top-0 right-0 w-40 h-40 bg-black -mr-20 -mt-20 rounded-full" />
              
              <div className="relative z-10">
                <div className="flex items-center gap-2 mb-6">
                  <div className="w-2 h-2 bg-[#facc15] rounded-full animate-pulse" />
                  <p className="text-gray-400 text-[11px] font-black uppercase tracking-[0.3em]">
                    {activeTab === 'international' ? 'Доставка з Китаю' : activeTab === 'novaposhta' ? 'Нова Пошта' : 'Переказ коштів'}
                  </p>
                </div>
                
                <div className="flex flex-col mb-12">
                  <span className="text-8xl font-display font-black tracking-tighter text-black leading-none">
                    {activeTab === 'international' 
                      ? `${internationalDetails.totalUah.toFixed(0)} грн` 
                      : activeTab === 'novaposhta' 
                        ? `${npDetails.totalUah.toFixed(0)} грн`
                        : `${transferDetails.totalUah.toFixed(0)} грн`
                    }
                  </span>
                  <span className="text-xl font-black text-black mt-2 uppercase tracking-widest">
                    {activeTab === 'international' 
                      ? `До сплати (UAH) / $${internationalDetails.total.toFixed(2)}` 
                      : 'До сплати (UAH)'}
                  </span>
                </div>

                <div className="space-y-6 pt-10 border-t border-gray-100">
                  {activeTab === 'international' && (
                    <>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Логістика</span>
                        <div className="text-right">
                          <div className="font-black text-2xl text-black">{internationalDetails.shippingCostUah.toFixed(0)} грн</div>
                          <div className="text-[10px] font-bold text-gray-400">${internationalDetails.shippingCost.toFixed(2)}</div>
                        </div>
                      </div>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Страхування</span>
                        <div className="text-right">
                          <div className="font-black text-2xl text-black">{internationalDetails.insuranceUah.toFixed(0)} грн</div>
                          <div className="text-[10px] font-bold text-gray-400">${internationalDetails.insurance.toFixed(2)}</div>
                        </div>
                      </div>
                      {(parcel.isFabric || parcel.isPressed || internationalDetails.packagingCost > 0 || internationalDetails.customsFee > 0 || internationalDetails.handlingFee > 0 || internationalDetails.fuelSurcharge > 0) && (
                        <div className="flex justify-between items-center group">
                          <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Дод. послуги та збори</span>
                          <div className="text-right">
                            <div className="font-black text-2xl text-black">{(internationalDetails.fabricSurchargeUah + internationalDetails.pressingCostUah + internationalDetails.packagingCostUah + internationalDetails.customsFeeUah + internationalDetails.handlingFeeUah + internationalDetails.fuelSurchargeUah).toFixed(0)} грн</div>
                            <div className="text-[10px] font-bold text-gray-400">${(internationalDetails.fabricSurcharge + internationalDetails.pressingCost + internationalDetails.packagingCost + internationalDetails.customsFee + internationalDetails.handlingFee + internationalDetails.fuelSurcharge).toFixed(2)}</div>
                          </div>
                        </div>
                      )}
                      {internationalDetails.localDelivery > 0 && (
                        <div className="flex justify-between items-center group">
                          <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Доставка по UA</span>
                          <div className="text-right">
                            <div className="font-black text-2xl text-black">{internationalDetails.localDeliveryUah.toFixed(0)} грн</div>
                            <div className="text-[10px] font-bold text-gray-400">${internationalDetails.localDelivery.toFixed(2)}</div>
                          </div>
                        </div>
                      )}
                    </>
                  )}
                  {activeTab === 'novaposhta' && (
                    <>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Тариф</span>
                        <span className="font-black text-2xl text-black">{npDetails.basePrice.toFixed(0)} грн</span>
                      </div>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Страхування</span>
                        <span className="font-black text-2xl text-black">{npDetails.insurance.toFixed(0)} грн</span>
                      </div>
                    </>
                  )}
                  {activeTab === 'transfer' && (
                    <>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Сума</span>
                        <span className="font-black text-2xl text-black">{transferData.amount.toFixed(0)} грн</span>
                      </div>
                      <div className="flex justify-between items-center group">
                        <span className="text-gray-400 text-[11px] font-black uppercase tracking-widest group-hover:text-black transition-colors">Комісія</span>
                        <span className="font-black text-2xl text-black">{transferDetails.fee.toFixed(0)} грн</span>
                      </div>
                    </>
                  )}
                </div>

                <button 
                  onClick={() => setView('crm')}
                  className="w-full mt-8 bg-[#facc15] text-black font-black py-4 rounded-xl shadow-xl shadow-yellow-900/10 hover:bg-[#eab308] hover:translate-y-[-1px] active:translate-y-[0.5px] transition-all flex items-center justify-center gap-3 uppercase tracking-[0.15em] text-[11px]"
                >
                  Оформити заявку в CRM
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </motion.div>

            {/* Details Card */}
            <div className="bg-white rounded-2xl p-6 shadow-2xl border-l-8 border-black">
              <h3 className="font-display font-black text-gray-900 mb-6 uppercase tracking-tight flex items-center gap-3 text-sm">
                <div className="w-8 h-8 bg-[#f8f9fa] rounded-lg flex items-center justify-center text-black">
                  <Info className="w-5 h-5" />
                </div>
                Параметри розрахунку
              </h3>
              <div className="space-y-8">
                {activeTab === 'international' && (
                  <>
                    <div className="flex justify-between items-end border-b border-gray-100 pb-6">
                      <div>
                        <p className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Щільність</p>
                        <p className="text-4xl font-black text-black mt-1">{density.toFixed(0)} <span className="text-sm font-bold text-gray-400">кг/м³</span></p>
                      </div>
                      <div className="text-right">
                        <p className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Об'єм</p>
                        <p className="text-2xl font-black text-yellow-500 mt-1">{finalVolumeM3.toFixed(3)} <span className="text-xs text-gray-400 font-bold">м³</span></p>
                      </div>
                    </div>

                    {selectedTariff.densityTiers && (
                      <div className="mt-6">
                        <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-3">Тарифна сітка ({selectedTariff.name})</p>
                        <div className="bg-gray-50 rounded-xl overflow-hidden border border-gray-100">
                          <table className="w-full text-[10px]">
                            <thead>
                              <tr className="bg-gray-100 text-gray-400 uppercase font-black">
                                <th className="p-2 text-left">Щільність</th>
                                <th className="p-2 text-right">Ціна</th>
                              </tr>
                            </thead>
                            <tbody className="font-bold text-black">
                              {selectedTariff.densityTiers.map((tier, i) => (
                                <tr key={i} className={`border-t border-gray-100 ${density >= tier.min && (tier.max === null || density < tier.max) ? 'bg-yellow-50 text-yellow-700' : ''}`}>
                                  <td className="p-2">
                                    {tier.max ? `${tier.min}-${tier.max}` : `>${tier.min}`} кг/м³
                                  </td>
                                  <td className="p-2 text-right">
                                    ${tier.price}/{tier.unit === 'kg' ? 'кг' : 'м³'}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}

                    <div className="mt-6 grid grid-cols-2 gap-4">
                      {selectedTariff.minWeight > 0 && (
                        <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <p className="text-[8px] font-black text-gray-400 uppercase tracking-widest">Мін. вага</p>
                          <p className="text-sm font-black text-black">{selectedTariff.minWeight} кг</p>
                        </div>
                      )}
                      {selectedTariff.minCost > 0 && (
                        <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <p className="text-[8px] font-black text-gray-400 uppercase tracking-widest">Мін. вартість</p>
                          <p className="text-sm font-black text-black">${selectedTariff.minCost}</p>
                        </div>
                      )}
                      {selectedTariff.insuranceRate > 0 && (
                        <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <p className="text-[8px] font-black text-gray-400 uppercase tracking-widest">Страховка</p>
                          <p className="text-sm font-black text-black">{selectedTariff.insuranceRate}%</p>
                        </div>
                      )}
                      {selectedTariff.localDeliveryPrice > 0 && (
                        <div className="p-3 bg-gray-50 rounded-xl border border-gray-100">
                          <p className="text-[8px] font-black text-gray-400 uppercase tracking-widest">Доставка UA</p>
                          <p className="text-sm font-black text-black">${selectedTariff.localDeliveryPrice}/кг</p>
                        </div>
                      )}
                    </div>
                  </>
                )}
                {activeTab === 'novaposhta' && (
                  <>
                    <div className="flex justify-between items-end border-b border-gray-100 pb-6">
                      <div>
                        <p className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Розрахункова вага</p>
                        <p className="text-4xl font-black text-black mt-1">{npDetails.chargeableWeight.toFixed(1)} <span className="text-sm font-bold text-gray-400">кг</span></p>
                      </div>
                      <div className="text-right">
                        <p className="text-[11px] font-black text-gray-400 uppercase tracking-[0.2em]">Об'ємна вага</p>
                        <p className="text-2xl font-black text-yellow-500 mt-1">{npDetails.volumetricWeight.toFixed(1)} <span className="text-xs text-gray-400 font-bold">кг</span></p>
                      </div>
                    </div>
                  </>
                )}
                {activeTab === 'transfer' && (
                  <div className="bg-[#f8f9fa] p-6 rounded-xl border border-gray-100">
                    <p className="text-[10px] text-gray-400 font-black uppercase tracking-[0.3em] mb-4">Умови переказу</p>
                    <p className="text-sm text-gray-600 leading-relaxed font-medium">
                      Комісія за переказ {transferData.method === 'card' ? 'на карту' : 'готівкою'} складає 
                      <span className="text-yellow-600 font-black"> {transferData.method === 'card' ? '1% + 5 грн' : '2% + 20 грн'}</span>.
                    </p>
                  </div>
                )}
              </div>
            </div>

          </div>
        </div>
      </main>
      </>
      ) : (
        <main className="max-w-7xl mx-auto px-6 py-12">
          <div className="flex flex-col lg:flex-row gap-10">
            {/* CRM Sidebar */}
            <aside className="w-full lg:w-64 space-y-4 no-print">
              {/* Role Switcher */}
              <div className="bg-white rounded-2xl p-4 shadow-sm border border-gray-100 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <UserCircle className="w-5 h-5 text-black" />
                  <span className="text-xs font-bold text-gray-500 uppercase tracking-wider">Роль:</span>
                </div>
                <select 
                  value={userRole}
                  onChange={(e) => setUserRole(e.target.value as UserRole)}
                  className="text-xs font-black text-black bg-gray-50 px-3 py-1.5 rounded-lg border-none focus:ring-2 focus:ring-yellow-400 cursor-pointer outline-none"
                >
                  <option value="admin">Адміністратор</option>
                  <option value="manager">Менеджер</option>
                </select>
              </div>

              <div className="bg-white rounded-2xl p-3 shadow-xl border-b-4 border-black">
                <h2 className="text-base font-black text-black uppercase tracking-tight mb-3 flex items-center gap-2">
                  <LayoutDashboard className="w-3.5 h-3.5" />
                  CRM FORSAGE CHINA
                </h2>
                <div className="space-y-0.5">
                  {crmModules
                    .map((module) => (
                      <button
                        key={module.id}
                        onClick={() => setCrmModule(module.id)}
                        className={cn(
                          "w-full flex items-center gap-2.5 p-2.5 rounded-xl font-bold text-[11px] transition-all text-left group",
                          crmModule === module.id 
                            ? "bg-black text-yellow-400 shadow-lg translate-x-1" 
                            : "text-gray-500 hover:bg-gray-50 hover:text-black"
                        )}
                      >
                      <module.icon className={cn(
                        "w-3.5 h-3.5",
                        crmModule === module.id ? "text-yellow-400" : "group-hover:text-yellow-400"
                      )} />
                      {module.title}
                    </button>
                  ))}
                </div>
              </div>

              <div className="bg-black rounded-2xl p-6 text-white shadow-2xl relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-24 h-24 bg-white/5 -mr-12 -mt-12 rounded-full group-hover:scale-110 transition-transform" />
                <Search className="w-8 h-8 text-yellow-400 mb-3" />
                <h3 className="text-base font-black uppercase tracking-tight mb-1">Швидкий пошук</h3>
                <p className="text-white/60 text-[10px] font-medium mb-4 leading-relaxed">Введіть трек-номер для миттєвого пошуку товару в системі.</p>
                <div className="relative">
                  <input 
                    type="text" 
                    placeholder="Трек-номер..."
                    value={purchaseSearch}
                    onChange={(e) => {
                      setPurchaseSearch(e.target.value);
                      if (crmModule !== 'purchases') setCrmModule('purchases');
                    }}
                    className="w-full bg-white/10 border border-white/20 rounded-xl px-4 py-2.5 text-xs outline-none focus:bg-white/20 transition-all placeholder:text-white/30"
                  />
                  <button className="absolute right-2 top-1/2 -translate-y-1/2 w-7 h-7 bg-yellow-400 rounded-lg flex items-center justify-center hover:bg-yellow-500 transition-colors">
                    <ArrowRight className="w-4 h-4 text-black" />
                  </button>
                </div>
              </div>
            </aside>

            {/* CRM Content Area */}
            <div className="flex-1">
              <motion.div
                key={crmModule}
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="bg-white rounded-3xl p-4 shadow-2xl border-t-8 border-yellow-400 min-h-[600px] relative overflow-hidden"
              >
                <div className="absolute top-0 right-0 w-64 h-64 bg-gray-50 -mr-32 -mt-32 rounded-full opacity-50" />
                
                <div className="relative z-10">
                  <div className="flex items-center gap-4 mb-4">
                    <div className="w-10 h-10 bg-[#f8f9fa] rounded-2xl flex items-center justify-center text-black shadow-inner">
                      {(() => {
                        const Icon = crmModules.find(m => m.id === crmModule)?.icon || LayoutDashboard;
                        return <Icon className="w-5 h-5" />;
                      })()}
                    </div>
                    <div>
                      <h1 className="text-xl font-black text-black uppercase tracking-tight">
                        {crmModules.find(m => m.id === crmModule)?.title}
                      </h1>
                      <p className="text-gray-400 font-bold uppercase tracking-widest text-[10px] mt-0.5">
                        {crmModules.find(m => m.id === crmModule)?.description}
                      </p>
                    </div>
                  </div>

                  {crmModule === 'dashboard' ? (
                    <div className="space-y-8">
                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                        <div className="bg-white p-6 rounded-3xl border-2 border-gray-100 shadow-sm">
                          <div className="flex items-center gap-4 mb-4">
                            <div className="w-12 h-12 bg-blue-50 rounded-2xl flex items-center justify-center text-blue-500">
                              <ShoppingCart className="w-6 h-6" />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Всього товарів</p>
                              <p className="text-3xl font-black text-black">{purchases.length}</p>
                            </div>
                          </div>
                        </div>
                        <div className="bg-white p-6 rounded-3xl border-2 border-gray-100 shadow-sm">
                          <div className="flex items-center gap-4 mb-4">
                            <div className="w-12 h-12 bg-amber-50 rounded-2xl flex items-center justify-center text-amber-500">
                              <Combine className="w-6 h-6" />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">В дорозі (Китай)</p>
                              <p className="text-3xl font-black text-black">{purchases.filter(p => p.status === 'shipped_to_ua' || p.status === 'arrived_china' || p.status === 'at_china_warehouse').length}</p>
                            </div>
                          </div>
                        </div>
                        <div className="bg-white p-6 rounded-3xl border-2 border-gray-100 shadow-sm">
                          <div className="flex items-center gap-4 mb-4">
                            <div className="w-12 h-12 bg-emerald-50 rounded-2xl flex items-center justify-center text-emerald-500">
                              <Store className="w-6 h-6" />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">На складі (Готові)</p>
                              <p className="text-3xl font-black text-black">{purchases.filter(p => p.status === 'my_warehouse').length}</p>
                            </div>
                          </div>
                        </div>
                        <div className="bg-white p-6 rounded-3xl border-2 border-gray-100 shadow-sm">
                          <div className="flex items-center gap-4 mb-4">
                            <div className="w-12 h-12 bg-purple-50 rounded-2xl flex items-center justify-center text-purple-500">
                              <CheckCircle2 className="w-6 h-6" />
                            </div>
                            <div>
                              <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Видано на магазин</p>
                              <p className="text-3xl font-black text-black">{purchases.filter(p => p.status === 'sold').length}</p>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                        <div className="bg-white rounded-3xl p-8 border-2 border-gray-100 shadow-sm">
                          <h3 className="text-xl font-black text-black uppercase tracking-tight mb-6 flex items-center gap-3">
                            <DollarSign className="w-6 h-6 text-yellow-500" />
                            Фінансова аналітика
                          </h3>
                          <div className="space-y-6">
                            <div className="flex justify-between items-center p-4 bg-gray-50 rounded-2xl">
                              <span className="text-sm font-bold text-gray-600">Заморожено в товарі (в дорозі)</span>
                              <span className="text-xl font-black text-black">
                                {purchases
                                  .filter(p => p.status !== 'sold' && p.status !== 'my_warehouse')
                                  .reduce((sum, p) => sum + (p.priceYuan * p.quantity * p.exchangeRate), 0).toFixed(0)} ₴
                              </span>
                            </div>
                            <div className="flex justify-between items-center p-4 bg-emerald-50 rounded-2xl">
                              <span className="text-sm font-bold text-emerald-700">Вартість товарів на складі</span>
                              <span className="text-xl font-black text-emerald-600">
                                {purchases
                                  .filter(p => p.status === 'my_warehouse')
                                  .reduce((sum, p) => {
                                    const costPriceUah = p.priceYuan * p.quantity * p.exchangeRate;
                                    const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
                                    const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
                                    return sum + costPriceUah + deliveryChinaUah + deliveryNPUah;
                                  }, 0).toFixed(0)} ₴
                              </span>
                            </div>
                            <div className="flex justify-between items-center p-4 bg-blue-50 rounded-2xl">
                              <span className="text-sm font-bold text-blue-700">Очікуваний прибуток (Мій склад)</span>
                              <span className="text-xl font-black text-blue-600">
                                {purchases
                                  .filter(p => p.status === 'my_warehouse')
                                  .reduce((sum, p) => {
                                    const costPriceUah = p.priceYuan * p.quantity * p.exchangeRate;
                                    const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
                                    const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
                                    const totalCostUah = costPriceUah + deliveryChinaUah + deliveryNPUah;
                                    const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));
                                    return sum + (sellingPriceUah - totalCostUah);
                                  }, 0).toFixed(0)} ₴
                              </span>
                            </div>
                            <div className="flex justify-between items-center p-4 bg-purple-50 rounded-2xl">
                              <span className="text-sm font-bold text-purple-700">Загальний прибуток (Продано)</span>
                              <span className="text-xl font-black text-purple-600">
                                {purchases
                                  .filter(p => p.status === 'sold')
                                  .reduce((sum, p) => {
                                    const costPriceUah = p.priceYuan * p.quantity * p.exchangeRate;
                                    const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0);
                                    const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
                                    const totalCostUah = costPriceUah + deliveryChinaUah + deliveryNPUah;
                                    const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));
                                    return sum + (sellingPriceUah - totalCostUah);
                                  }, 0).toFixed(0)} ₴
                              </span>
                            </div>
                          </div>
                        </div>

                        <div className="bg-white rounded-3xl p-8 border-2 border-gray-100 shadow-sm">
                          <h3 className="text-xl font-black text-black uppercase tracking-tight mb-6 flex items-center gap-3">
                            <BarChart3 className="w-6 h-6 text-purple-500" />
                            Статуси товарів
                          </h3>
                          <div className="space-y-4">
                            {Object.entries(statusLabels).map(([status, label]) => {
                              const count = purchases.filter(p => p.status === status).length;
                              const percentage = purchases.length > 0 ? (count / purchases.length) * 100 : 0;
                              if (count === 0) return null;
                              return (
                                <div key={status} className="space-y-2">
                                  <div className="flex justify-between text-sm font-bold">
                                    <span className="text-gray-600">{label}</span>
                                    <span className="text-black">{count} шт</span>
                                  </div>
                                  <div className="h-2 bg-gray-100 rounded-full overflow-hidden">
                                    <div 
                                      className="h-full bg-black rounded-full"
                                      style={{ width: `${percentage}%` }}
                                    />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    </div>
                  ) : crmModule === 'purchases' ? (
                    <div className="space-y-8">
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        <div className="bg-[#f8f9fa] p-6 rounded-2xl border border-gray-100">
                          <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-2">Всього закупок</p>
                          <p className="text-3xl font-black text-black">{purchases.length}</p>
                        </div>
                        <div className="bg-[#f8f9fa] p-6 rounded-2xl border border-gray-100">
                          <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-2">Кількість товарів</p>
                          <p className="text-3xl font-black text-gray-400">
                            {purchases.reduce((acc, p) => acc + p.quantity, 0)}
                          </p>
                        </div>
                      </div>
                      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
                        <div className="flex flex-wrap gap-4 w-full lg:w-auto">
                          <div className="relative flex-1 lg:flex-none min-w-[200px]">
                            <input 
                              type="text" 
                              value={purchaseSearch}
                              onChange={(e) => setPurchaseSearch(e.target.value)}
                              placeholder="Пошук..." 
                              className="w-full pl-10 pr-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-black transition-all"
                            />
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                          </div>
                          <select 
                            value={statusFilter}
                            onChange={(e) => setStatusFilter(e.target.value)}
                            className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black cursor-pointer"
                          >
                            <option value="all">Всі статуси</option>
                            <option value="purchased">Куплено</option>
                          </select>
                          <button 
                            onClick={() => setShowBulkAddModal(true)}
                            className="bg-gray-800 text-white px-8 py-5 rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-3 hover:bg-gray-900 transition-all shadow-lg shadow-gray-100 whitespace-nowrap"
                          >
                            <Plus className="w-5 h-5" />
                            Нова закупка
                          </button>
                          {selectedPurchaseIds.length > 0 && (
                            <>
                              <button 
                                onClick={() => setShowAssignTrackModal(true)}
                                className="bg-blue-500 text-white px-8 py-5 rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-3 hover:bg-blue-600 transition-all shadow-lg shadow-blue-100 whitespace-nowrap"
                              >
                                <Layers className="w-5 h-5" />
                                Призначити накладну ({selectedPurchaseIds.length})
                              </button>
                              <button 
                                onClick={() => {
                                  if (confirm(`Ви впевнені, що хочете видалити ${selectedPurchaseIds.length} вибраних товарів?`)) {
                                    handleBulkDelete(selectedPurchaseIds);
                                  }
                                }}
                                className="bg-red-500 text-white px-8 py-5 rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-3 hover:bg-red-600 transition-all shadow-lg shadow-red-100 whitespace-nowrap"
                              >
                                <Trash2 className="w-5 h-5" />
                                Видалити вибрані
                              </button>
                            </>
                          )}
                        </div>
                      </div>

                      {purchases.filter(p => p.status === 'purchased').length > 0 ? (
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="border-b border-gray-100">
                                <th className="py-4 px-4 w-10">
                                  <button 
                                    onClick={() => {
                                      const unassignedIds = purchases
                                        .filter(p => p.status === 'purchased' && !p.trackNumber)
                                        .map(p => p.id);
                                      toggleSelectAll(unassignedIds);
                                    }}
                                    className="text-gray-400 hover:text-black transition-colors"
                                  >
                                    {selectedPurchaseIds.length > 0 ? <CheckSquare className="w-5 h-5 text-black" /> : <Square className="w-5 h-5" />}
                                  </button>
                                </th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Товар</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Платформа</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Трек-номер</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна (¥)</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">К-сть</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Разом (¥)</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Дії</th>
                              </tr>
                            </thead>
                            <tbody>
                              {purchases
                                .filter(p => {
                                  const matchesSearch = (p.name || '').toLowerCase().includes(purchaseSearch.toLowerCase());
                                  const isPurchased = p.status === 'purchased';
                                  return matchesSearch && isPurchased;
                                })
                                .map((p) => (
                                  <tr key={p.id} className={clsx(
                                    "border-b border-gray-50 hover:bg-gray-50 transition-colors",
                                    selectedPurchaseIds.includes(p.id) && "bg-yellow-50/30"
                                  )}>
                                    <td className="py-4 px-4">
                                      <button 
                                        onClick={() => toggleSelectOne(p.id)}
                                        className="text-gray-400 hover:text-black transition-colors"
                                      >
                                        {selectedPurchaseIds.includes(p.id) ? <CheckSquare className="w-5 h-5 text-black" /> : <Square className="w-5 h-5" />}
                                      </button>
                                    </td>
                                    <td className="py-4 px-4">
                                      <div className="flex items-center gap-3">
                                        <div className="w-10 h-10 rounded-xl bg-gray-50 overflow-hidden flex-shrink-0 border border-gray-100">
                                          {p.photo ? <img src={p.photo} alt={p.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <Package className="w-5 h-5 m-2.5 text-gray-300" />}
                                        </div>
                                        <div>
                                          <p className="text-sm font-bold text-black">{p.name}</p>
                                          {p.link && <a href={p.link} target="_blank" rel="noreferrer" className="text-[10px] text-blue-500 hover:underline">Посилання</a>}
                                        </div>
                                      </div>
                                    </td>
                                    <td className="py-4 px-4">
                                      <span className="text-[10px] font-black text-gray-500 uppercase tracking-widest bg-gray-100 px-2 py-0.5 rounded">
                                        {p.platform}
                                      </span>
                                    </td>
                                    <td className="py-4 px-4">
                                      <div className="relative group/track min-w-[120px]">
                                        <input 
                                          type="text"
                                          list="existing-tracks"
                                          placeholder="Призначити..."
                                          className="w-full bg-transparent border-b border-gray-100 hover:border-gray-300 focus:border-black focus:outline-none py-1 text-xs font-bold transition-colors"
                                          onKeyDown={(e) => {
                                            if (e.key === 'Enter') {
                                              const val = (e.target as HTMLInputElement).value;
                                              if (val) handleAssignTrack([p.id], val);
                                            }
                                          }}
                                          onBlur={(e) => {
                                            const val = e.target.value;
                                            if (val) handleAssignTrack([p.id], val);
                                          }}
                                        />
                                        <datalist id="existing-tracks">
                                          {existingTrackNumbers.map(track => (
                                            <option key={track} value={track} />
                                          ))}
                                        </datalist>
                                      </div>
                                    </td>
                                    <td className="py-4 px-4 text-sm font-bold text-gray-600">{p.priceYuan} ¥</td>
                                    <td className="py-4 px-4 text-sm font-bold text-gray-600">{p.quantity}</td>
                                    <td className="py-4 px-4 text-sm font-black text-black">{(p.priceYuan * p.quantity).toFixed(2)} ¥</td>
                                    <td className="py-4 px-4">
                                      <div className="flex items-center gap-2">
                                        <button 
                                          onClick={() => handleEditPurchase(p)}
                                          className="p-2 text-gray-400 hover:text-black hover:bg-gray-100 rounded-xl transition-colors"
                                        >
                                          <Edit2 className="w-4 h-4" />
                                        </button>
                                        <button 
                                          onClick={() => handleDeletePurchase(p.id)}
                                          className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-colors"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>
                      ) : (
                        <div className="border-2 border-dashed border-gray-100 rounded-3xl p-20 flex flex-col items-center justify-center text-center">
                          <div className="w-24 h-24 bg-gray-50 rounded-full flex items-center justify-center mb-6">
                            <ShoppingCart className="w-10 h-10 text-gray-200" />
                          </div>
                          <h3 className="text-xl font-black text-gray-300 uppercase tracking-tight mb-2">Закупок немає</h3>
                          <p className="text-gray-400 text-sm max-w-xs mx-auto">
                            Ви ще не додали жодної закупки. Натисніть кнопку "Нова закупка", щоб розпочати.
                          </p>
                        </div>
                      )}
                    </div>
                  ) : crmModule === 'china_warehouse' ? (
                    <div className="space-y-8">
                      <div className="flex justify-between items-center">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Накладні на складі Китай</h2>
                        <div className="flex gap-4">
                          <div className="relative">
                            <input 
                              type="text" 
                              value={chinaWarehouseSearch}
                              onChange={(e) => setChinaWarehouseSearch(e.target.value)}
                              placeholder="Пошук за треком..." 
                              className="pl-10 pr-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-black transition-all"
                            />
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                          </div>
                          <div className="flex gap-4">
                            <button 
                              onClick={() => setShowImportWaybillModal(true)}
                              className="bg-indigo-600 text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] flex items-center gap-2 hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100"
                            >
                              <ClipboardList className="w-4 h-4" />
                              Імпорт накладних
                            </button>
                            <button 
                              onClick={() => {
                                setCrmModule('purchases');
                                addNotification('Виберіть товари в закупках для створення нової накладної', 'info');
                              }}
                              className="bg-black text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] flex items-center gap-2 hover:bg-gray-900 transition-all shadow-lg shadow-gray-100"
                            >
                              <Plus className="w-4 h-4" />
                              Додати накладну
                            </button>
                          </div>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 gap-6">
                        {waybills.length > 0 ? (
                          waybills.map((w: any) => (
                            <div key={w.trackNumber} className="bg-white border border-gray-100 rounded-3xl p-8 shadow-sm hover:shadow-md transition-all group">
                              <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
                                <div className="space-y-2">
                                  <div className="flex items-center gap-3">
                                    <div className="w-12 h-12 bg-amber-50 rounded-2xl flex items-center justify-center">
                                      <Receipt className="w-6 h-6 text-amber-600" />
                                    </div>
                                    <div>
                                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Трек-номер</p>
                                      <button 
                                        onClick={() => setSelectedTrackNumber(w.trackNumber)}
                                        className="text-xl font-black text-black hover:text-blue-600 transition-colors flex items-center gap-2"
                                      >
                                        {w.trackNumber}
                                        <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-opacity" />
                                      </button>
                                    </div>
                                  </div>
                                </div>

                                <div className="grid grid-cols-2 md:grid-cols-5 gap-8 flex-1 px-0 lg:px-12">
                                  <div>
                                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Товари</p>
                                    <p className="text-sm font-bold text-black">{w.items.length} поз. ({w.totalQuantity} шт)</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вага</p>
                                    <p className="text-sm font-bold text-black">{w.totalWeight.toFixed(2)} кг</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вартість</p>
                                    <p className="text-sm font-bold text-black">¥{w.totalCostYuan.toFixed(2)}</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Доставка</p>
                                    <p className="text-sm font-bold text-indigo-600">${w.totalDeliveryCost.toFixed(2)}</p>
                                  </div>
                                  <div>
                                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Статус</p>
                                    <span className="text-[10px] font-black uppercase tracking-widest px-3 py-1.5 rounded-lg bg-amber-50 text-amber-600">
                                      🏬 Склад Китай
                                    </span>
                                  </div>
                                </div>

                                <div className="flex gap-3">
                                  <button 
                                    onClick={() => {
                                      const newTrack = prompt('Введіть новий трек-номер для цієї накладної:', w.trackNumber);
                                      if (newTrack) handleEditWaybill(w.trackNumber, newTrack);
                                    }}
                                    className="p-3 text-gray-400 hover:text-blue-500 hover:bg-blue-50 rounded-xl transition-colors"
                                    title="Редагувати трек-номер"
                                  >
                                    <Edit2 className="w-5 h-5" />
                                  </button>
                                  <button 
                                    onClick={() => {
                                      const ids = w.items.map((i: any) => i.id);
                                      setPurchases(purchases.map(p => ids.includes(p.id) ? { ...p, status: 'shipped_to_ua' } : p));
                                      addNotification(`Накладну ${w.trackNumber} відправлено в Україну`, 'success');
                                    }}
                                    className="bg-indigo-50 text-indigo-600 px-6 py-3 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-100 transition-colors flex items-center gap-2"
                                  >
                                    <Truck className="w-4 h-4" />
                                    Відправити
                                  </button>
                                  <button 
                                    onClick={() => {
                                      if (confirm(`Ви впевнені, що хочете видалити накладну ${w.trackNumber}? Всі товари в ній будуть видалені.`)) {
                                        const ids = w.items.map((i: any) => i.id);
                                        handleBulkDelete(ids);
                                      }
                                    }}
                                    className="p-3 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-colors"
                                  >
                                    <Trash2 className="w-5 h-5" />
                                  </button>
                                </div>
                              </div>
                              
                              <div className="mt-6 pt-6 border-t border-gray-50 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                                {w.items.slice(0, 3).map((item: any) => (
                                  <div key={item.id} className="flex items-center gap-3 bg-gray-50/50 p-3 rounded-xl border border-gray-100/50">
                                    <div className="w-10 h-10 rounded-lg bg-white overflow-hidden flex-shrink-0 border border-gray-100">
                                      {item.photo ? <img src={item.photo} alt={item.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <Package className="w-5 h-5 m-2.5 text-gray-300" />}
                                    </div>
                                    <div className="min-w-0">
                                      <p className="text-xs font-bold text-black truncate">{item.name}</p>
                                      <p className="text-[10px] font-black text-gray-400 uppercase">{item.quantity} шт • ¥{item.priceYuan}</p>
                                    </div>
                                  </div>
                                ))}
                                {w.items.length > 3 && (
                                  <button 
                                    onClick={() => setSelectedTrackNumber(w.trackNumber)}
                                    className="flex items-center justify-center gap-2 text-[10px] font-black text-gray-400 uppercase tracking-widest hover:text-black transition-colors"
                                  >
                                    Ще {w.items.length - 3} товарів...
                                  </button>
                                )}
                              </div>
                            </div>
                          ))
                        ) : (
                          <div className="bg-white border border-gray-100 rounded-3xl p-20 text-center">
                            <div className="w-20 h-20 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-6">
                              <Receipt className="w-10 h-10 text-gray-300" />
                            </div>
                            <h3 className="text-xl font-black text-gray-300 uppercase tracking-tight mb-2">Накладних немає</h3>
                            <p className="text-gray-400 text-sm max-w-xs mx-auto">
                              Тут з'являться товари, яким ви призначили трек-номер у розділі "Закупки".
                            </p>
                          </div>
                        )}
                      </div>
                    </div>
                  ) : crmModule === 'consolidation' ? (
                    <ErrorBoundary>
                      <div className="space-y-8">
                        <div className="flex justify-between items-center">
                          <h2 className="text-2xl font-black text-black uppercase tracking-tight">В дорозі з Китаю</h2>
                          <button 
                            onClick={() => setShowCreateBatchModal(true)}
                            className="bg-[#facc15] text-black px-8 py-4 rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-3 hover:bg-[#eab308] transition-all shadow-lg shadow-yellow-100"
                          >
                          <Truck className="w-5 h-5" />
                          Створити партію доставки
                        </button>
                      </div>

                      {purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId).length > 0 && (
                        <div className="bg-white border border-gray-100 rounded-3xl p-8 shadow-sm mb-8">
                          <h3 className="text-xl font-black text-black uppercase tracking-tight mb-6">Товари в дорозі (без партії)</h3>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="border-b border-gray-100">
                                  <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Трек номер</th>
                                  <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва товару</th>
                                  <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Кількість</th>
                                  <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Статус</th>
                                  <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Вага (кг)</th>
                                </tr>
                              </thead>
                              <tbody>
                                {purchases
                                  .filter(p => p.status === 'shipped_to_ua' && !p.batchId)
                                  .map((p) => (
                                    <tr key={p.id} className="border-b border-gray-50 hover:bg-gray-50 transition-colors">
                                      <td className="py-4 px-4">
                                        <div className="flex items-center gap-2">
                                          <button 
                                            onClick={() => setSelectedTrackNumber(p.trackNumber)}
                                            className="text-xs font-mono font-black text-blue-600 bg-blue-50 px-2 py-1 rounded inline-flex items-center gap-2 hover:bg-blue-100 transition-all group/track"
                                            title="Переглянути товари цього трек-номеру"
                                          >
                                            {p.trackNumber || '—'}
                                            <Layers className="w-3 h-3 opacity-0 group-hover/track:opacity-100 transition-opacity" />
                                          </button>
                                          <button 
                                            onClick={() => copyToClipboard(p.trackNumber)}
                                            className="text-gray-400 hover:text-blue-500 transition-colors"
                                            title="Скопіювати"
                                          >
                                            <Layers className="w-3 h-3" />
                                          </button>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4">
                                        <p className="text-sm font-bold text-black">{p.name}</p>
                                      </td>
                                      <td className="py-4 px-4 text-sm font-bold text-gray-600">
                                        <input
                                          type="number"
                                          className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                          defaultValue={p.quantity || 1}
                                          onBlur={(e) => {
                                            const val = parseInt(e.target.value, 10);
                                            if (!isNaN(val) && val !== p.quantity && val > 0) {
                                              setPurchases(purchases.map(item => item.id === p.id ? { ...item, quantity: val } : item));
                                              addNotification(`Кількість оновлено`, 'success');
                                            }
                                          }}
                                        />
                                      </td>
                                      <td className="py-4 px-4">
                                        <div className="flex items-center gap-2">
                                          <span className="text-[10px] font-black uppercase tracking-widest px-3 py-1.5 rounded-lg bg-indigo-50 text-indigo-600">
                                            ✈️ В дорозі
                                          </span>
                                          <button 
                                            onClick={() => {
                                              setPurchases(purchases.map(item => item.id === p.id ? { ...item, status: 'arrived_ua' } : item));
                                              addNotification(`Товар "${p.name}" прибув в Україну`, 'success');
                                            }}
                                            className="bg-emerald-50 text-emerald-600 px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest hover:bg-emerald-100 transition-colors flex items-center gap-1"
                                            title="Прибуло в Україну"
                                          >
                                            <CheckCircle2 className="w-3 h-3" />
                                            Прибуло
                                          </button>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4 text-sm font-bold text-black">
                                        <div className="flex items-center gap-1 group">
                                          <input
                                            key={p.weight}
                                            type="number"
                                            className="w-16 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-right"
                                            defaultValue={p.weight || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.weight) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, weight: val } : item));
                                                addNotification(`Вагу оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span className="text-gray-400">{p.weightUnit || 'кг'}</span>
                                        </div>
                                      </td>
                                    </tr>
                                  ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {batches.length > 0 ? (
                        <div className="grid grid-cols-1 gap-6">
                          {batches.map((batch) => (
                            <div key={batch.id} className="bg-white border border-gray-100 rounded-3xl p-8 shadow-sm hover:shadow-md transition-all">
                              <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 mb-8">
                                <div className="flex items-center gap-4">
                                  <div className={cn(
                                    "w-14 h-14 rounded-2xl flex items-center justify-center",
                                    batch.status === 'arrived_ua' ? "bg-emerald-50 text-emerald-600" : "bg-indigo-50 text-indigo-600"
                                  )}>
                                    <Package className="w-7 h-7" />
                                  </div>
                                  <div>
                                    <h3 className="text-xl font-black text-black uppercase tracking-tight">{batch.name}</h3>
                                    <div className="flex items-center gap-3 mt-1">
                                      <span className="text-xs font-bold text-gray-400 flex items-center gap-1">
                                        <Calendar className="w-3 h-3" /> {batch.shipmentDate}
                                      </span>
                                      <span className="text-xs font-bold text-gray-400 flex items-center gap-1">
                                        <MapPin className="w-3 h-3" /> {batch.warehouse}
                                      </span>
                                      <span className="text-[10px] font-black uppercase tracking-widest bg-gray-100 px-2 py-0.5 rounded text-gray-500">
                                        {batch.deliveryType === 'air' ? '✈️ Авіа' : '🚢 Море'}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="flex flex-wrap items-center gap-4">
                                  {batch.status === 'shipped' && (
                                    <>
                                      <button 
                                        onClick={() => setShowCostModal({ show: true, batchId: batch.id })}
                                        className="bg-amber-50 text-amber-600 px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-amber-100 transition-all flex items-center gap-2"
                                      >
                                        <DollarSign className="w-4 h-4" />
                                        Внести вартість
                                      </button>
                                      <button 
                                        onClick={() => handleBatchArrived(batch.id)}
                                        className="bg-emerald-50 text-emerald-600 px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-emerald-100 transition-all flex items-center gap-2"
                                      >
                                        <CheckCircle2 className="w-4 h-4" />
                                        Партія прибула
                                      </button>
                                    </>
                                  )}
                                  <div className={cn(
                                    "px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest",
                                    batch.status === 'arrived_ua' ? "bg-emerald-50 text-emerald-600" : "bg-indigo-50 text-indigo-600"
                                  )}>
                                    {batch.status === 'arrived_ua' ? '📥 Прибула в Україну' : '✈️ В дорозі'}
                                  </div>
                                  <div className="flex items-center gap-2">
                                    <button 
                                      onClick={() => handleEditBatch(batch)}
                                      className="p-2 text-gray-400 hover:text-black transition-colors"
                                      title="Редагувати партію"
                                    >
                                      <Edit2 className="w-4 h-4" />
                                    </button>
                                    <button 
                                      onClick={() => handleDeleteBatch(batch.id)}
                                      className="p-2 text-gray-400 hover:text-red-500 transition-colors"
                                      title="Видалити партію"
                                    >
                                      <Trash2 className="w-4 h-4" />
                                    </button>
                                  </div>
                                </div>
                              </div>

                              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вага партії</p>
                                  <p className="text-lg font-black text-black">{batch.totalWeight ? `${batch.totalWeight} кг` : '—'}</p>
                                </div>
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вартість доставки (грн)</p>
                                  <p className="text-lg font-black text-black">{batch.deliveryCost ? `${(batch.deliveryCost * usdToUah).toFixed(0)} грн` : '—'}</p>
                                </div>
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Ціна за кг (грн)</p>
                                  <p className="text-lg font-black text-amber-600">{batch.pricePerKg ? `${(batch.pricePerKg * usdToUah).toFixed(0)} грн` : '—'}</p>
                                </div>
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Товарів</p>
                                  <p className="text-lg font-black text-gray-400">{batch.itemIds?.length || 0}</p>
                                </div>
                              </div>

                              <div className="border-t border-gray-50 pt-6">
                                <details className="group">
                                  <summary className="flex items-center justify-between cursor-pointer list-none">
                                    <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest group-open:text-black transition-colors">Список товарів у партії</span>
                                    <ChevronDown className="w-4 h-4 text-gray-400 group-open:rotate-180 transition-transform" />
                                  </summary>
                                  <div className="mt-4 space-y-3">
                                    {purchases.filter(p => batch.itemIds?.includes(p.id)).map(p => (
                                      <div key={p.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl hover:bg-gray-100 transition-colors">
                                        <div className="flex items-center gap-3">
                                          <div className="w-8 h-8 bg-white rounded-lg flex items-center justify-center text-[10px] font-mono font-bold text-blue-500 shadow-sm">
                                            {p.trackNumber.slice(-4)}
                                          </div>
                                          <div>
                                            <p className="text-xs font-bold text-black">{p.name}</p>
                                            <p className="text-[10px] text-gray-400">{p.weight || 0} {p.weightUnit || 'кг'}</p>
                                          </div>
                                        </div>
                                        <div className="text-right">
                                          <p className="text-xs font-black text-black">
                                            {p.deliveryCostPerItem ? `${(p.deliveryCostPerItem * usdToUah).toFixed(0)} грн` : '—'}
                                          </p>
                                          <p className="text-[10px] text-gray-400 uppercase font-black tracking-widest">доставка</p>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </details>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="border-2 border-dashed border-gray-100 rounded-3xl p-20 flex flex-col items-center justify-center text-center">
                          <div className="w-24 h-24 bg-gray-50 rounded-full flex items-center justify-center mb-6">
                            <Truck className="w-10 h-10 text-gray-200" />
                          </div>
                          <h3 className="text-xl font-black text-gray-300 uppercase tracking-tight mb-2">Партій не створено</h3>
                          <p className="text-gray-400 text-sm max-w-xs mx-auto">
                            Ви ще не створили жодної партії доставки. Натисніть кнопку "Створити партію доставки", щоб розпочати консолідацію.
                          </p>
                        </div>
                      )}
                    </div>
                  </ErrorBoundary>
                  ) : crmModule === 'ua_warehouse' ? (
                    <div className="space-y-8">
                      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Товари на складі Україна (Київ)</h2>
                        <div className="flex flex-wrap gap-4 w-full lg:w-auto">
                          <div className="relative flex-1 lg:flex-none min-w-[200px]">
                            <input 
                              type="text" 
                              value={uaWarehouseSearch}
                              onChange={(e) => setUaWarehouseSearch(e.target.value)}
                              placeholder="Пошук за треком..." 
                              className="w-full pl-10 pr-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-black transition-all"
                            />
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                          </div>
                          
                          <input 
                            type="date" 
                            value={uaWarehouseDateFilter}
                            onChange={(e) => setUaWarehouseDateFilter(e.target.value)}
                            className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black cursor-pointer"
                          />

                          <select 
                            value={uaWarehouseBatchFilter}
                            onChange={(e) => setUaWarehouseBatchFilter(e.target.value)}
                            className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black cursor-pointer"
                          >
                            <option value="">Всі партії</option>
                            {Array.from(new Set(purchases.filter(p => p.batchId).map(p => p.batchId))).map(batchId => (
                              <option key={batchId} value={batchId}>{batchId}</option>
                            ))}
                          </select>

                          <button 
                            onClick={() => handleBulkStatusChange('arrived_ua')}
                            className="bg-emerald-500 text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] flex items-center gap-2 hover:bg-emerald-600 transition-all shadow-lg shadow-emerald-100"
                          >
                            <CheckCircle2 className="w-4 h-4" />
                            Прибуло в Україну ({selectedPurchaseIds.length})
                          </button>
                          <button 
                            onClick={() => {
                              const filtered = purchases.filter(p => {
                                const matchesStatus = p.status === 'arrived_ua';
                                const matchesSearch = (p.trackNumber || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase()) || 
                                                     (p.name || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase());
                                const matchesDate = !uaWarehouseDateFilter || (p.arrivalDate && p.arrivalDate.includes(uaWarehouseDateFilter));
                                const matchesBatch = !uaWarehouseBatchFilter || p.batchId === uaWarehouseBatchFilter;
                                return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                              });
                              
                              const dataToExport = selectedPurchaseIds.length > 0 
                                ? filtered.filter(p => selectedPurchaseIds.includes(p.id))
                                : filtered;
                              exportToExcel(dataToExport, 'UA_Warehouse_Export');
                            }}
                            className="bg-emerald-500 text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] flex items-center gap-2 hover:bg-emerald-600 transition-all shadow-lg shadow-emerald-100"
                          >
                            <Download className="w-4 h-4" />
                            Експорт Excel {selectedPurchaseIds.length > 0 ? `(${selectedPurchaseIds.length})` : ''}
                          </button>

                          <button 
                            onClick={() => {
                              const filtered = purchases.filter(p => {
                                const matchesStatus = p.status === 'arrived_ua';
                                const matchesSearch = (p.trackNumber || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase()) || 
                                                     (p.name || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase());
                                const matchesDate = !uaWarehouseDateFilter || (p.arrivalDate && p.arrivalDate.includes(uaWarehouseDateFilter));
                                const matchesBatch = !uaWarehouseBatchFilter || p.batchId === uaWarehouseBatchFilter;
                                return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                              });
                              
                              const dataToExport = selectedPurchaseIds.length > 0 
                                ? filtered.filter(p => selectedPurchaseIds.includes(p.id))
                                : filtered;
                              exportToExcelWithPhotos(dataToExport, 'UA_Warehouse_With_Photos');
                            }}
                            className="bg-indigo-500 text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] flex items-center gap-2 hover:bg-indigo-600 transition-all shadow-lg shadow-indigo-100"
                          >
                            <Download className="w-4 h-4" />
                            Excel з фото
                          </button>
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="border-b border-gray-100">
                              <th className="py-4 px-4 w-10">
                                <button 
                                  onClick={() => {
                                    const filteredIds = purchases
                                      .filter(p => {
                                        const matchesStatus = p.status === 'arrived_ua';
                                        const matchesSearch = (p.trackNumber || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase()) || 
                                                             (p.name || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase());
                                        const matchesDate = !uaWarehouseDateFilter || (p.arrivalDate && p.arrivalDate.includes(uaWarehouseDateFilter));
                                        const matchesBatch = !uaWarehouseBatchFilter || p.batchId === uaWarehouseBatchFilter;
                                        return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                                      })
                                      .map(p => p.id);
                                    toggleSelectAll(filteredIds);
                                  }}
                                  className="text-gray-400 hover:text-black transition-colors"
                                >
                                  {selectedPurchaseIds.length > 0 ? <CheckSquare className="w-5 h-5 text-black" /> : <Square className="w-5 h-5" />}
                                </button>
                              </th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center">Кількість</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Вага (кг)</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Собівартість (грн)</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Доставка (грн)</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Повна собівартість (грн)</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Статус</th>
                              <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Дії</th>
                            </tr>
                          </thead>
                          <tbody>
                            {purchases
                              .filter(p => {
                                const matchesStatus = p.status === 'arrived_ua';
                                const matchesSearch = (p.trackNumber || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase()) || 
                                                     (p.name || '').toLowerCase().includes(uaWarehouseSearch.toLowerCase());
                                const matchesDate = !uaWarehouseDateFilter || (p.arrivalDate && p.arrivalDate.includes(uaWarehouseDateFilter));
                                const matchesBatch = !uaWarehouseBatchFilter || p.batchId === uaWarehouseBatchFilter;
                                return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                              })
                              .map((p) => (
                                <tr key={p.id} className={clsx(
                                  "border-b border-gray-50 hover:bg-gray-50 transition-colors",
                                  selectedPurchaseIds.includes(p.id) && "bg-yellow-50/30"
                                )}>
                                  <td className="py-4 px-4">
                                    <button 
                                      onClick={() => toggleSelectOne(p.id)}
                                      className="text-gray-400 hover:text-black transition-colors"
                                    >
                                      {selectedPurchaseIds.includes(p.id) ? <CheckSquare className="w-5 h-5 text-black" /> : <Square className="w-5 h-5" />}
                                    </button>
                                  </td>
                                  <td className="py-4 px-4">
                                    <p className="text-sm font-bold text-black">{p.name}</p>
                                    <div className="flex items-center gap-2 mt-1">
                                      <button 
                                        onClick={() => setSelectedTrackNumber(p.trackNumber)}
                                        className="text-[10px] font-black text-blue-500 hover:text-blue-700 uppercase tracking-widest flex items-center gap-1 group"
                                        title="Переглянути товари цього трек-номеру"
                                      >
                                        {p.trackNumber}
                                        <Layers className="w-3 h-3 opacity-0 group-hover:opacity-100 transition-opacity" />
                                      </button>
                                      <button 
                                        onClick={() => copyToClipboard(p.trackNumber)}
                                        className="text-gray-400 hover:text-blue-500 transition-colors"
                                        title="Скопіювати"
                                      >
                                        <Layers className="w-3 h-3" />
                                      </button>
                                    </div>
                                  </td>
                                  <td className="py-4 px-4 text-sm font-bold text-gray-600 text-center">
                                    <input
                                      type="number"
                                      className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-center"
                                      defaultValue={p.quantity || 1}
                                      onBlur={(e) => {
                                        const val = parseInt(e.target.value, 10);
                                        if (!isNaN(val) && val !== p.quantity && val > 0) {
                                          setPurchases(purchases.map(item => item.id === p.id ? { ...item, quantity: val } : item));
                                          addNotification(`Кількість оновлено`, 'success');
                                        }
                                      }}
                                    />
                                  </td>
                                  <td className="py-4 px-4 text-sm font-bold text-black">
                                    <div className="flex items-center gap-1 group">
                                      <input
                                        key={p.weight}
                                        type="number"
                                        className="w-16 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-right"
                                        defaultValue={p.weight || ''}
                                        onBlur={(e) => {
                                          const val = parseFloat(e.target.value);
                                          if (!isNaN(val) && val !== p.weight) {
                                            setPurchases(purchases.map(item => item.id === p.id ? { ...item, weight: val } : item));
                                            addNotification(`Вагу оновлено`, 'success');
                                          }
                                        }}
                                      />
                                      <span className="text-gray-400">{p.weightUnit || 'кг'}</span>
                                    </div>
                                  </td>
                                  <td className="py-4 px-4 text-sm font-bold text-gray-600">
                                    {(p.priceYuan * p.exchangeRate * p.quantity).toFixed(0)} грн
                                  </td>
                                  <td className="py-4 px-4 text-sm font-bold text-amber-600">
                                    {((p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah).toFixed(0)} грн
                                  </td>
                                  <td className="py-4 px-4">
                                    <p className="text-sm font-black text-black">
                                      {(p.priceYuan * p.exchangeRate * p.quantity + (p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah).toFixed(0)} грн
                                    </p>
                                  </td>
                                  <td className="py-4 px-4">
                                    <span className="text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full bg-emerald-100 text-emerald-500">
                                      На складі Україна
                                    </span>
                                  </td>
                                  <td className="py-4 px-4">
                                    <div className="flex items-center gap-2">
                                      <button 
                                        onClick={() => handleEditPurchase(p)}
                                        className="p-2 text-gray-400 hover:text-blue-500 transition-colors"
                                        title="Редагувати"
                                      >
                                        <Edit2 className="w-4 h-4" />
                                      </button>
                                      <button 
                                        onClick={() => handleDeletePurchase(p.id)}
                                        className="p-2 text-gray-400 hover:text-red-500 transition-colors"
                                        title="Видалити"
                                      >
                                        <Trash2 className="w-4 h-4" />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                        {purchases.filter(p => p.status === 'arrived_ua').length === 0 && (
                          <div className="py-20 text-center">
                            <p className="text-gray-400 font-bold">На складі в Україні поки немає товарів</p>
                          </div>
                        )}
                      </div>
                    </div>
                  ) : crmModule === 'local_delivery' ? (
                    <div className="space-y-8">
                      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-6">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Нова Пошта / Самовивіз</h2>
                        <div className="flex gap-4">
                          <button 
                            onClick={async () => {
                              const apiKey = localStorage.getItem('novaPoshtaApiKey');
                              console.log('Button clicked, API Key from localStorage:', apiKey);
                              if (!apiKey) {
                                console.log('API Key is missing');
                                addNotification('Вкажіть API ключ Нової Пошти в налаштуваннях', 'error');
                                return;
                              }
                              console.log('API Key found, starting update');
                              setIsUpdatingTracking(true);
                              try {
                                const endDate = new Date().toISOString().split('T')[0].split('-').reverse().join('.');
                                const startDate = new Date();
                                startDate.setDate(startDate.getDate() - 30);
                                const startDateStr = startDate.toISOString().split('T')[0].split('-').reverse().join('.');
                                
                                const data = await getDocumentList(apiKey, startDateStr, endDate);
                                console.log('Nova Poshta API response (raw array):', data);
                                
                                if (Array.isArray(data) && data.length > 0) {
                                  addNotification(`Завантажено ${data.length} відправок`, 'success');
                                  console.log('Nova Poshta documents:', data);
                                  // TODO: Add logic to process these shipments
                                } else if (Array.isArray(data) && data.length === 0) {
                                  addNotification('Відправок не знайдено', 'info');
                                } else {
                                  addNotification('Неочікуваний формат відповіді від API', 'error');
                                  console.log('Unexpected API response format:', data);
                                }
                              } catch (err: any) {
                                addNotification(err.message, 'error');
                              } finally {
                                setIsUpdatingTracking(false);
                              }
                            }}
                            disabled={isUpdatingTracking}
                            className="bg-blue-500 text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-blue-600 transition-all shadow-lg shadow-blue-200 flex items-center gap-2 disabled:opacity-50"
                          >
                            <Clock className="w-4 h-4" />
                            {isUpdatingTracking ? 'Завантаження...' : 'Завантажити всі відправки'}
                          </button>
                          <button 
                            onClick={() => {
                              setEditingLocalDeliveryId(null);
                              setLocalDeliveryForm({
                                name: '',
                                type: 'novaposhta',
                                trackingNumber: '',
                                cost: 0
                              });
                              setShowCreateLocalDeliveryModal(true);
                            }}
                            className="bg-black text-white px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-gray-800 transition-all shadow-lg shadow-gray-200 flex items-center gap-2"
                          >
                            <Plus className="w-4 h-4" />
                            Створити відправку
                          </button>
                        </div>
                      </div>

                      {localDeliveries.length > 0 ? (
                        <div className="grid gap-6">
                          {localDeliveries.map(ld => (
                            <div key={ld.id} className="bg-white p-8 rounded-[32px] border-2 border-gray-100 hover:border-black transition-colors shadow-sm relative overflow-hidden group">
                              <div className="absolute top-0 right-0 w-32 h-32 bg-gray-50 rounded-bl-full -z-10 group-hover:bg-gray-100 transition-colors" />
                              
                              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8">
                                <div>
                                  <div className="flex items-center gap-3 mb-2">
                                    <h3 className="text-2xl font-black text-black uppercase tracking-tight">{ld.name}</h3>
                                    <span className={cn(
                                      "px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest",
                                      ld.type === 'novaposhta' ? "bg-red-50 text-red-600" : "bg-blue-50 text-blue-600"
                                    )}>
                                      {ld.type === 'novaposhta' ? 'Нова Пошта' : 'Самовивіз'}
                                    </span>
                                  </div>
                                  <p className="text-sm font-bold text-gray-400">
                                    Створено: {new Date(ld.createdAt).toLocaleDateString('uk-UA')}
                                  </p>
                                  {ld.trackingNumber && ld.type === 'novaposhta' && (
                                    <div className="mt-4 p-4 bg-white rounded-2xl border border-gray-100">
                                      <div className="flex justify-between items-center mb-2">
                                        <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Статус Нової Пошти</p>
                                        <button 
                                          disabled={trackingModal.loading}
                                          onClick={async () => {
                                            const apiKey = localStorage.getItem('novaPoshtaApiKey');
                                            if (!apiKey) {
                                              addNotification('Вкажіть API ключ Нової Пошти в налаштуваннях', 'error');
                                              return;
                                            }
                                            setTrackingModal({ show: true, data: null, loading: true, error: '' });
                                            try {
                                              const data = await trackParcel(apiKey, ld.trackingNumber!);
                                              setTrackingModal({ show: true, data: data[0], loading: false, error: '' });
                                            } catch (err: any) {
                                              setTrackingModal({ show: true, data: null, loading: false, error: err.message });
                                              addNotification(err.message, 'error');
                                            }
                                          }}
                                          className="text-[10px] font-black text-blue-500 hover:text-blue-700 uppercase tracking-widest disabled:opacity-50"
                                        >
                                          {trackingModal.loading ? '...' : 'Оновити'}
                                        </button>
                                      </div>
                                      <p className="text-sm font-black text-black">
                                        {ld.trackingNumber}
                                      </p>
                                    </div>
                                  )}
                                  {ld.trackingNumber && ld.type !== 'novaposhta' && (
                                    <p className="text-sm font-bold text-gray-600 mt-1">
                                      ТТН: <span className="text-black">{ld.trackingNumber}</span>
                                    </p>
                                  )}
                                </div>
                                
                                <div className="flex flex-wrap items-center gap-3">
                                  {ld.status === 'pending' && (
                                    <button 
                                      onClick={() => handleCompleteLocalDelivery(ld.id)}
                                      className="bg-emerald-50 text-emerald-600 px-6 py-3 rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-emerald-100 transition-all flex items-center gap-2"
                                    >
                                      <CheckCircle2 className="w-4 h-4" />
                                      Отримано
                                    </button>
                                  )}
                                  <div className={cn(
                                    "px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest",
                                    ld.status === 'received' ? "bg-emerald-50 text-emerald-600" : "bg-amber-50 text-amber-600"
                                  )}>
                                    {ld.status === 'received' ? '✅ Отримано' : '⏳ В процесі'}
                                  </div>
                                  <div className="flex items-center gap-2">
                                    <button 
                                      onClick={() => handleEditLocalDelivery(ld)}
                                      className="p-2 text-gray-400 hover:text-black transition-colors"
                                      title="Редагувати відправку"
                                    >
                                      <Edit2 className="w-4 h-4" />
                                    </button>
                                    <button 
                                      onClick={() => handleDeleteLocalDelivery(ld.id)}
                                      className="p-2 text-gray-400 hover:text-red-500 transition-colors"
                                      title="Видалити відправку"
                                    >
                                      <Trash2 className="w-4 h-4" />
                                    </button>
                                  </div>
                                </div>
                              </div>

                              <div className="grid grid-cols-2 md:grid-cols-3 gap-4 mb-8">
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вартість доставки</p>
                                  <p className="text-lg font-black text-black">{ld.cost} грн</p>
                                </div>
                                <div className="bg-gray-50 p-4 rounded-2xl">
                                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Товарів</p>
                                  <p className="text-lg font-black text-gray-400">{ld.itemIds.length}</p>
                                </div>
                              </div>

                              <div className="border-t border-gray-50 pt-6">
                                <details className="group">
                                  <summary className="flex items-center justify-between cursor-pointer list-none">
                                    <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest group-open:text-black transition-colors">Список товарів у відправці</span>
                                    <ChevronDown className="w-4 h-4 text-gray-400 group-open:rotate-180 transition-transform" />
                                  </summary>
                                  <div className="mt-4 space-y-3">
                                    {purchases.filter(p => ld.itemIds.includes(p.id)).map(p => (
                                      <div key={p.id} className="flex items-center justify-between p-3 bg-gray-50 rounded-xl hover:bg-gray-100 transition-colors">
                                        <div className="flex items-center gap-3">
                                          <div className="w-8 h-8 bg-white rounded-lg flex items-center justify-center text-[10px] font-mono font-bold text-blue-500 shadow-sm">
                                            {p.trackNumber.slice(-4)}
                                          </div>
                                          <div>
                                            <p className="text-xs font-bold text-black">{p.name}</p>
                                            <p className="text-[10px] text-gray-400">{p.weight || 0} {p.weightUnit || 'кг'}</p>
                                          </div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </details>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="border-2 border-dashed border-gray-100 rounded-3xl p-20 flex flex-col items-center justify-center text-center">
                          <div className="w-24 h-24 bg-gray-50 rounded-full flex items-center justify-center mb-6">
                            <Truck className="w-10 h-10 text-gray-200" />
                          </div>
                          <h3 className="text-xl font-black text-gray-300 uppercase tracking-tight mb-2">Відправок не створено</h3>
                          <p className="text-gray-400 text-sm max-w-xs mx-auto">
                            Ви ще не створили жодної відправки. Натисніть кнопку "Створити відправку", щоб розпочати.
                          </p>
                        </div>
                      )}
                    </div>
                  ) : crmModule === 'issue_to_store' ? (
                    <div className="space-y-8">
                      <div className="flex justify-between items-center">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Видача на магазин</h2>
                        <div className="flex gap-4">
                          <select
                            value={salesFilter}
                            onChange={(e) => setSalesFilter(e.target.value as any)}
                            className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-black transition-all"
                          >
                            <option value="all">Всі видачі</option>
                            <option value="physical_store">Фізичний магазин</option>
                            <option value="online_store">Інтернет-магазин</option>
                            <option value="personal_use">Власні потреби</option>
                          </select>
                          <div className="relative">
                            <input 
                              type="text" 
                              value={salesSearch}
                              onChange={(e) => setSalesSearch(e.target.value)}
                              placeholder="Пошук за назвою..." 
                              className="pl-10 pr-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-sm font-bold outline-none focus:ring-2 focus:ring-black transition-all"
                            />
                            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                          </div>
                        </div>
                      </div>

                      <div className="space-y-4">
                        <h3 className="text-xl font-black text-black uppercase tracking-tight">Доступні для видачі (Прибуло на мій склад)</h3>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="border-b border-gray-100">
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Товар</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center">Кількість</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Собівартість (грн)</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Дії</th>
                              </tr>
                            </thead>
                            <tbody>
                              {purchases
                                .filter(p => p.status === 'my_warehouse' && ((p.name || '').toLowerCase().includes(salesSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(salesSearch.toLowerCase())))
                                .map((p) => {
                                  const costPrice = (p.priceYuan * p.quantity) / p.exchangeRate;
                                  const chinaDelivery = p.deliveryCostPerItem || 0;
                                  const uaDelivery = p.ukraineDeliveryCost || 0;
                                  const npDelivery = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
                                  const totalCost = costPrice + chinaDelivery + uaDelivery + npDelivery;

                                  return (
                                    <tr key={p.id} className="border-b border-gray-50 hover:bg-gray-50 transition-colors">
                                      <td className="py-4 px-4">
                                        <p className="text-sm font-bold text-black">{p.name}</p>
                                        <div className="flex items-center gap-2 mt-1">
                                          <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">{p.trackNumber}</span>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4 text-sm font-bold text-gray-600 text-center">
                                        <input
                                          type="number"
                                          className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-center"
                                          defaultValue={p.quantity || 1}
                                          onBlur={(e) => {
                                            const val = parseInt(e.target.value, 10);
                                            if (!isNaN(val) && val !== p.quantity && val > 0) {
                                              setPurchases(purchases.map(item => item.id === p.id ? { ...item, quantity: val } : item));
                                              addNotification(`Кількість оновлено`, 'success');
                                            }
                                          }}
                                        />
                                      </td>
                                      <td className="py-4 px-4 text-sm font-black text-black">
                                        {totalCost.toFixed(0)} грн
                                      </td>
                                      <td className="py-4 px-4 text-right">
                                        <button 
                                          onClick={() => setShowSaleModal({ show: true, purchaseId: p.id })}
                                          className="px-4 py-2 bg-black text-white rounded-xl text-xs font-black uppercase tracking-widest hover:bg-gray-800 transition-colors"
                                        >
                                          Видати
                                        </button>
                                      </td>
                                    </tr>
                                  );
                                })}
                            </tbody>
                          </table>
                          {purchases.filter(p => p.status === 'my_warehouse').length === 0 && (
                            <div className="py-10 text-center">
                              <p className="text-gray-400 font-bold">Немає товарів для видачі</p>
                            </div>
                          )}
                        </div>
                      </div>

                      <div className="space-y-4">
                        <h3 className="text-xl font-black text-black uppercase tracking-tight">Видані товари</h3>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse">
                            <thead>
                              <tr className="border-b border-gray-100">
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Дата продажу</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Товар</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-center">Кількість</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Собівартість</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна продажу</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Прибуток</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Куди</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest">Статус</th>
                                <th className="py-4 px-4 text-[10px] font-black text-gray-400 uppercase tracking-widest text-right">Дії</th>
                              </tr>
                            </thead>
                            <tbody>
                              {purchases
                                .filter(p => p.status === 'sold' && 
                                  (salesFilter === 'all' || p.saleDestination === salesFilter) &&
                                  ((p.name || '').toLowerCase().includes(salesSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(salesSearch.toLowerCase())))
                                .map((p) => {
                                  const costPrice = (p.priceYuan * p.quantity) * p.exchangeRate;
                                  const chinaDelivery = (p.deliveryCostPerItem || 0) * usdToUah;
                                  const intDelivery = (p.shippingCost || 0) * usdToUah;
                                  const uaDelivery = p.ukraineDeliveryCost || 0;
                                  const npDelivery = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p);
                                  const totalCost = costPrice + chinaDelivery + intDelivery + uaDelivery + npDelivery;
                                  const sellingPriceUah = p.sellingPrice || 0;
                                  const profit = sellingPriceUah - totalCost;

                                  return (
                                    <tr key={p.id} className="border-b border-gray-50 hover:bg-gray-50 transition-colors">
                                      <td className="py-4 px-4 text-sm font-bold text-gray-600">
                                        {p.soldDate ? new Date(p.soldDate).toLocaleDateString() : '-'}
                                      </td>
                                      <td className="py-4 px-4">
                                        <p className="text-sm font-bold text-black">{p.name}</p>
                                        <div className="flex items-center gap-2 mt-1">
                                          <button 
                                            onClick={() => setSelectedTrackNumber(p.trackNumber)}
                                            className="text-[10px] font-black text-blue-500 hover:text-blue-700 uppercase tracking-widest flex items-center gap-1 group"
                                            title="Переглянути товари цього трек-номеру"
                                          >
                                            {p.trackNumber}
                                            <Layers className="w-3 h-3 opacity-0 group-hover:opacity-100 transition-opacity" />
                                          </button>
                                          <button 
                                            onClick={() => copyToClipboard(p.trackNumber)}
                                            className="text-gray-400 hover:text-blue-500 transition-colors"
                                            title="Скопіювати"
                                          >
                                            <Layers className="w-3 h-3" />
                                          </button>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4 text-sm font-bold text-gray-600 text-center">
                                        <input
                                          type="number"
                                          className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-center"
                                          defaultValue={p.quantity || 1}
                                          onBlur={(e) => {
                                            const val = parseInt(e.target.value, 10);
                                            if (!isNaN(val) && val !== p.quantity && val > 0) {
                                              setPurchases(purchases.map(item => item.id === p.id ? { ...item, quantity: val } : item));
                                              addNotification(`Кількість оновлено`, 'success');
                                            }
                                          }}
                                        />
                                      </td>
                                      <td className="py-4 px-4 text-sm font-black text-gray-500">
                                        {totalCost.toFixed(0)} грн
                                      </td>
                                      <td className="py-4 px-4 text-sm font-black text-black">
                                        <div className="flex items-center gap-1 group">
                                          <input
                                            type="number"
                                            className="w-20 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-sm font-black text-black"
                                            defaultValue={sellingPriceUah.toFixed(0)}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== sellingPriceUah) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, sellingPrice: val, markup: true } : item));
                                                addNotification(`Ціну продажу оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span>грн</span>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4 text-sm font-black">
                                        <span className={profit >= 0 ? 'text-emerald-500' : 'text-red-500'}>
                                          {profit > 0 ? '+' : ''}{profit.toFixed(0)} грн
                                        </span>
                                      </td>
                                      <td className="py-4 px-4">
                                        <span className={cn(
                                          "px-2 py-1 rounded-full text-[9px] font-black uppercase tracking-widest",
                                          p.saleDestination === 'physical_store' ? "bg-blue-50 text-blue-600" :
                                          p.saleDestination === 'online_store' ? "bg-purple-50 text-purple-600" :
                                          "bg-orange-50 text-orange-600"
                                        )}>
                                          {p.saleDestination === 'physical_store' ? 'Магазин' :
                                           p.saleDestination === 'online_store' ? 'Інтернет' : 'Власне'}
                                        </span>
                                      </td>
                                      <td className="py-4 px-4">
                                        <div className="flex items-center gap-2">
                                          <div className={cn(
                                            "w-2 h-2 rounded-full",
                                            profit >= 0 ? "bg-emerald-400" : "bg-red-400"
                                          )} />
                                          <span className="text-xs font-bold text-gray-600">Продано</span>
                                        </div>
                                      </td>
                                      <td className="py-4 px-4 text-right">
                                        <div className="flex justify-end gap-2">
                                          <button 
                                            onClick={() => handleEditSale(p)}
                                            className="p-2 text-gray-400 hover:text-black transition-colors"
                                            title="Редагувати продаж"
                                          >
                                            <FileText className="w-4 h-4" />
                                          </button>
                                          <button 
                                            onClick={() => handleDeleteSale(p.id)}
                                            className="p-2 text-gray-400 hover:text-red-500 transition-colors"
                                            title="Видалити продаж"
                                          >
                                            <Trash2 className="w-4 h-4" />
                                          </button>
                                        </div>
                                      </td>
                                    </tr>
                                  );
                                })}
                            </tbody>
                          </table>
                          {purchases.filter(p => p.status === 'sold').length === 0 && (
                            <div className="py-10 text-center">
                              <p className="text-gray-400 font-bold">Продажів поки немає</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  ) : crmModule === 'my_warehouse' ? (
                    <div id="price-list-container" className="space-y-8 p-8 bg-white border-2 border-gray-100 rounded-[40px]">
                      {/* PDF Header - Only visible in print/PDF */}
                      <div className={cn("mb-10 border-b-4 border-black pb-8", isExportingPDF ? "block" : "hidden print:block")}>
                        <div className="flex justify-between items-end">
                          <div>
                            <h1 className="text-5xl font-black text-black uppercase tracking-tighter mb-2">ПРАЙС-ЛИСТ</h1>
                            <p className="text-xl font-bold text-black uppercase tracking-widest">FORSAGE CHINA DELIVERY</p>
                          </div>
                          <div className="text-right">
                            <p className="text-sm font-black text-gray-400 uppercase tracking-widest mb-1">Дата формування</p>
                            <p className="text-2xl font-black text-black">{new Date().toLocaleDateString('uk-UA')}</p>
                          </div>
                        </div>
                      </div>

                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 no-print">
                        <div>
                          <h2 className="text-2xl font-black text-black uppercase tracking-tight">Формування прайс-листа</h2>
                          <p className="text-gray-500 text-sm font-bold">Товари на складі в Україні з розрахунком повної вартості</p>
                        </div>
                        <div className="flex flex-col gap-4 no-print">
                          {selectedPurchaseIds.length > 0 && (
                            <div className="flex items-center gap-2 p-2 bg-black text-white rounded-2xl no-print">
                              <span className="text-[10px] font-black uppercase tracking-widest px-2">Вибрано: {selectedPurchaseIds.length}</span>
                              <button 
                                onClick={() => {
                                  if (confirm(`Ви впевнені, що хочете перевести ${selectedPurchaseIds.length} товарів у статус "Прибуло в Україну"?`)) {
                                    setPurchases(purchases.map(p => selectedPurchaseIds.includes(p.id) ? { ...p, status: 'arrived_ua' } : p));
                                    setSelectedPurchaseIds([]);
                                    addNotification(`Статус змінено для ${selectedPurchaseIds.length} товарів`, 'success');
                                  }
                                }}
                                className="px-4 py-2 bg-white text-black rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-gray-200 transition-all"
                              >
                                Прибуло в Україну
                              </button>
                              <button 
                                onClick={() => {
                                  // Barcode generation logic
                                  selectedPurchaseIds.forEach(id => {
                                    const purchase = purchases.find(p => p.id === id);
                                    if (purchase) {
                                      JsBarcode(`#barcode-${id}`, purchase.trackNumber, {
                                        format: "CODE128",
                                        lineColor: "#000",
                                        width: 2,
                                        height: 50,
                                        displayValue: true
                                      });
                                    }
                                  });
                                  setTimeout(() => {
                                    window.print();
                                  }, 500);
                                }}
                                className="px-4 py-2 bg-yellow-400 text-black rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-yellow-500 transition-all"
                              >
                                Створити етикетки
                              </button>
                            </div>
                          )}
                          <div id="barcode-print-area" className="hidden print:block">
                            {selectedPurchaseIds.map(id => {
                              const purchase = purchases.find(p => p.id === id);
                              if (!purchase) return null;
                              return (
                                <div key={id} className="mb-4">
                                  <svg id={`barcode-${id}`}></svg>
                                  <p className="text-center text-xs font-bold">{purchase.trackNumber}</p>
                                </div>
                              );
                            })}
                          </div>
                          <div className="flex flex-wrap items-center gap-4">
                            <div className="relative">
                              <Search className="w-4 h-4 absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" />
                              <input 
                                type="text"
                                placeholder="Пошук..."
                                value={priceListSearch}
                                onChange={(e) => setPriceListSearch(e.target.value)}
                                className="pl-10 pr-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black w-48"
                              />
                            </div>
                            <div className="relative">
                              <Zap className="w-4 h-4 absolute left-4 top-1/2 -translate-y-1/2 text-yellow-500" />
                              <input 
                                type="text" 
                                placeholder="Сканер..." 
                                value={scannerInput}
                                onChange={(e) => {
                                  setScannerInput(e.target.value);
                                  // Add scanner logic here
                                  const found = purchases.find(p => p.trackNumber === e.target.value);
                                  if (found) {
                                    if (isInventoryMode) {
                                      addNotification(`Товар ${found.trackNumber} знайдено`, 'info');
                                    } else {
                                      setPurchases(purchases.map(p => p.id === found.id ? { ...p, status: 'arrived_ua' } : p));
                                      addNotification(`Товар ${found.trackNumber} переміщено на склад`, 'success');
                                    }
                                    setScannerInput('');
                                  }
                                }}
                                autoFocus
                                className="pl-12 pr-4 py-3 bg-yellow-50 rounded-xl border border-yellow-100 font-bold focus:outline-none focus:ring-2 focus:ring-yellow-500"
                              />
                            </div>
                            <button 
                              onClick={() => setShowScanner(true)}
                              className="p-3 bg-gray-100 rounded-xl hover:bg-gray-200 transition-colors"
                            >
                              <Camera className="w-4 h-4 text-gray-600" />
                            </button>
                            <button 
                              onClick={() => setIsInventoryMode(!isInventoryMode)}
                              className={cn(
                                "px-4 py-3 rounded-xl text-xs font-black uppercase tracking-widest transition-all",
                                isInventoryMode ? "bg-red-500 text-white" : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                              )}
                            >
                              {isInventoryMode ? 'Вийти з інвентаризації' : 'Інвентаризація'}
                            </button>
                            <input 
                              type="date" 
                              value={priceListDateFilter}
                              onChange={(e) => setPriceListDateFilter(e.target.value)}
                              className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black cursor-pointer"
                            />

                            <select 
                              value={priceListBatchFilter}
                              onChange={(e) => setPriceListBatchFilter(e.target.value)}
                              className="px-4 py-3 bg-gray-50 border border-gray-100 rounded-xl text-xs font-black uppercase tracking-widest outline-none focus:ring-2 focus:ring-black cursor-pointer"
                            >
                              <option value="">Всі партії</option>
                              {Array.from(new Set(purchases.filter(p => p.batchId).map(p => p.batchId))).map(batchId => (
                                <option key={batchId} value={batchId}>{batchId}</option>
                              ))}
                            </select>

                            <div className="flex items-center gap-4 bg-white p-2 rounded-2xl shadow-sm border border-gray-100">
                              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest px-2">Націнка (%)</label>
                              <select 
                                value={priceListMargin}
                                onChange={(e) => setPriceListMargin(parseInt(e.target.value) || 0)}
                                className="w-32 p-2 bg-gray-50 rounded-xl border border-gray-100 font-black text-black focus:outline-none focus:ring-2 focus:ring-black"
                              >
                                <option value={0}>Без націнки (0%)</option>
                                <option value={50}>50%</option>
                                <option value={100}>100%</option>
                                <option value={150}>150%</option>
                                <option value={200}>200%</option>
                                <option value={250}>250%</option>
                                <option value={300}>300%</option>
                              </select>
                            </div>
                          </div>
                          <div className="flex items-center gap-4">
                            <div className="flex items-center gap-4 bg-white p-2 rounded-2xl shadow-sm border border-gray-100">
                              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest px-2">Курс CNY/UAH</label>
                              <input 
                                type="number" 
                                value={cnyToUah}
                                onChange={(e) => setCnyToUah(parseFloat(e.target.value) || 0)}
                                className="w-20 p-2 bg-gray-50 rounded-xl border border-gray-100 font-black text-black text-center focus:outline-none"
                              />
                            </div>
                            <button 
                              className={cn(
                                "flex items-center gap-2 px-3 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all shadow-sm border",
                                showStorePreview ? "bg-black text-white border-black" : "bg-white text-gray-400 border-gray-200 hover:bg-gray-50"
                              )}
                            >
                              <LayoutGrid className="w-3 h-3" />
                              {showStorePreview ? 'Адмін вид' : 'Вид магазину'}
                            </button>
                            <div className="flex bg-gray-100 p-1 rounded-xl no-print">
                              <button 
                                onClick={() => setPriceListView('grid')}
                                className={cn(
                                  "px-3 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                                  priceListView === 'grid' ? "bg-black text-white shadow-md" : "text-gray-400 hover:text-gray-600"
                                )}
                              >
                                <LayoutGrid className="w-3 h-3" />
                              </button>
                              <button 
                                onClick={() => setPriceListView('table')}
                                className={cn(
                                  "px-3 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                                  priceListView === 'table' ? "bg-black text-white shadow-md" : "text-gray-400 hover:text-gray-600"
                                )}
                              >
                                <List className="w-3 h-3" />
                              </button>
                            </div>
                            <button 
                              onClick={() => {
                                const filtered = purchases.filter(p => {
                                  const matchesStatus = p.status === 'my_warehouse';
                                  const matchesSearch = !priceListSearch || (p.name || '').toLowerCase().includes(priceListSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(priceListSearch.toLowerCase());
                                  const matchesDate = !priceListDateFilter || (p.arrivalDate && p.arrivalDate.includes(priceListDateFilter)) || (p.createdAt && p.createdAt.includes(priceListDateFilter));
                                  const matchesBatch = !priceListBatchFilter || p.batchId === priceListBatchFilter;
                                  return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                                });
                                exportToExcel(filtered, 'Price_List_Export', false);
                              }}
                              className="bg-emerald-500 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-emerald-600 transition-all shadow-lg shadow-emerald-100 flex items-center gap-2"
                            >
                              <Download className="w-3 h-3" />
                              Експорт Excel
                            </button>
                            <button 
                              onClick={() => {
                                const filtered = purchases.filter(p => {
                                  const matchesStatus = p.status === 'my_warehouse';
                                  const matchesSearch = !priceListSearch || (p.name || '').toLowerCase().includes(priceListSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(priceListSearch.toLowerCase());
                                  const matchesDate = !priceListDateFilter || (p.arrivalDate && p.arrivalDate.includes(priceListDateFilter)) || (p.createdAt && p.createdAt.includes(priceListDateFilter));
                                  const matchesBatch = !priceListBatchFilter || p.batchId === priceListBatchFilter;
                                  return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                                });
                                exportToExcelWithPhotos(filtered, 'Price_List_With_Photos', false);
                              }}
                              className="bg-indigo-500 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-indigo-600 transition-all shadow-lg shadow-indigo-100 flex items-center gap-2"
                            >
                              <Download className="w-3 h-3" />
                              Excel з фото
                            </button>
                            <button 
                              onClick={handleExportPDF}
                              disabled={isExportingPDF}
                              className={cn(
                                "bg-yellow-400 text-black px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-100 flex items-center gap-2",
                                isExportingPDF && "opacity-50 cursor-not-allowed"
                              )}
                            >
                              {isExportingPDF ? (
                                <div className="w-3 h-3 border-2 border-black border-t-transparent rounded-full animate-spin" />
                              ) : (
                                <Download className="w-3 h-3" />
                              )}
                              {isExportingPDF ? 'Генерація...' : 'Зберегти як PDF'}
                            </button>
                            <button 
                              onClick={() => window.print()}
                              className="bg-gray-800 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-black transition-all shadow-lg shadow-gray-100 flex items-center gap-2 no-print"
                            >
                              <Printer className="w-3 h-3" />
                              Друк
                            </button>
                          </div>
                          {isIframe && (
                            <p className="text-[10px] text-amber-600 font-bold text-right no-print">
                              Порада: Якщо друк не працює, відкрийте додаток у новій вкладці
                            </p>
                          )}
                        </div>
                      </div>

                      {priceListView === 'grid' ? (
                        <div className={cn(
                          "grid gap-6 p-4 bg-white",
                          showStorePreview ? "grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4" : "grid-cols-1 md:grid-cols-2 xl:grid-cols-3"
                        )}>
                          {purchases
                            .filter(p => {
                              const matchesStatus = p.status === 'my_warehouse';
                              const matchesSearch = !priceListSearch || (p.name || '').toLowerCase().includes(priceListSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(priceListSearch.toLowerCase());
                              const matchesDate = !priceListDateFilter || (p.arrivalDate && p.arrivalDate.includes(priceListDateFilter)) || (p.createdAt && p.createdAt.includes(priceListDateFilter));
                              const matchesBatch = !priceListBatchFilter || p.batchId === priceListBatchFilter;
                              return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                            })
                            .map(p => {
                              const costPriceYuan = p.priceYuan * p.quantity;
                              const costPriceUah = costPriceYuan * p.exchangeRate;
                              
                              const trackPurchases = p.trackNumber ? purchases.filter(item => item.trackNumber === p.trackNumber) : [p];
                              let totalDeliveryForGroup = 0;
                              let totalWeightForGroup = 0;
                              let totalDeliveryChinaUah = 0;
                              let totalDeliveryNPUah = 0;
                              
                              trackPurchases.forEach(item => {
                                const deliveryChinaUah = (item.deliveryCostPerItem || 0) * usdToUah; 
                                const deliveryIntUah = (item.shippingCost || 0) * usdToUah; 
                                const deliveryUAUah = (item.ukraineDeliveryCost || 0); 
                                const deliveryNPUah = (item.novaPoshtaCost || 0) + getLocalDeliveryCost(item); 
                                totalDeliveryChinaUah += deliveryChinaUah + deliveryIntUah;
                                totalDeliveryNPUah += deliveryUAUah + deliveryNPUah;
                                totalDeliveryForGroup += deliveryChinaUah + deliveryIntUah + deliveryUAUah + deliveryNPUah;
                                const actualWeight = item.weightUnit === 'g' ? (item.weight || 0) / 1000 : (item.weight || 0);
                                totalWeightForGroup += actualWeight;
                              });

                              const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
                              let itemDeliveryUah = 0;
                              let itemDeliveryChinaUah = 0;
                              let itemDeliveryNPUah = 0;

                              if (trackPurchases.length === 1) {
                                itemDeliveryUah = totalDeliveryForGroup;
                                itemDeliveryChinaUah = totalDeliveryChinaUah;
                                itemDeliveryNPUah = totalDeliveryNPUah;
                              } else if (totalWeightForGroup > 0) {
                                const ratio = actualWeight / totalWeightForGroup;
                                itemDeliveryUah = totalDeliveryForGroup * ratio;
                                itemDeliveryChinaUah = totalDeliveryChinaUah * ratio;
                                itemDeliveryNPUah = totalDeliveryNPUah * ratio;
                              } else {
                                const ratio = 1 / trackPurchases.length;
                                itemDeliveryUah = totalDeliveryForGroup * ratio;
                                itemDeliveryChinaUah = totalDeliveryChinaUah * ratio;
                                itemDeliveryNPUah = totalDeliveryNPUah * ratio;
                              }
                              
                              const totalCostUah = costPriceUah + itemDeliveryUah;
                              const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));
                              const sellingPriceYuan = sellingPriceUah / p.exchangeRate;

                              const dimString = p.width ? `${p.width}x${p.height}x${p.length} ${p.dimUnit}` : p.size || '-';
                              const weightString = p.weight ? `${p.weight} ${p.weightUnit}` : '-';

                              return (
                                <motion.div 
                                  key={p.id}
                                  layout
                                  initial={{ opacity: 0, scale: 0.9 }}
                                  animate={{ opacity: 1, scale: 1 }}
                                  className={cn(
                                    "bg-white rounded-3xl overflow-hidden shadow-xl border flex flex-col group transition-all break-inside-avoid",
                                    isExportingPDF ? "border-2 border-gray-200 shadow-none" : "border-gray-100 hover:border-black print:border-2 print:border-gray-200 print:shadow-none"
                                  )}
                                >
                                  <div className={cn("h-64 relative overflow-hidden bg-gray-50 border-b", isExportingPDF ? "border-gray-200" : "border-gray-50 print:border-gray-200")}>
                                    {p.photo ? (
                                      <img src={p.photo} alt={p.name} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-500" referrerPolicy="no-referrer" />
                                    ) : (
                                      <div className="w-full h-full flex items-center justify-center">
                                        <Package className="w-12 h-12 text-gray-200" />
                                      </div>
                                    )}
                                      <div className="absolute top-4 left-4 no-print">
                                        <span className="bg-white/90 backdrop-blur-sm text-black px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest shadow-sm">
                                          {p.platform}
                                        </span>
                                      </div>
                                    {p.markup && (
                                      <div className="absolute top-4 right-4">
                                        <span className="bg-emerald-500 text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest shadow-lg">
                                          Націнка
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                  <div className="p-6 flex-1 flex flex-col">
                                      <div className="flex justify-between items-start gap-4 mb-4">
                                        <h4 className="text-lg font-black text-black line-clamp-2 leading-tight flex-1">{p.name}</h4>
                                        <div className="text-right flex flex-col items-end">
                                          <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Трек</p>
                                          <div className="flex items-center gap-2">
                                            <button 
                                              onClick={() => setSelectedTrackNumber(p.trackNumber)}
                                              className="text-xs font-black text-blue-500 hover:text-blue-700 flex items-center gap-1 group"
                                              title="Переглянути товари цього трек-номеру"
                                            >
                                              {p.trackNumber}
                                              <Layers className="w-3 h-3 opacity-0 group-hover:opacity-100 transition-opacity" />
                                            </button>
                                            <button 
                                              onClick={() => copyToClipboard(p.trackNumber)}
                                              className="text-gray-400 hover:text-blue-500 transition-colors"
                                              title="Скопіювати"
                                            >
                                              <Layers className="w-3 h-3" />
                                            </button>
                                          </div>
                                        </div>
                                      </div>
                                    
                                    {p.comment && (
                                      <div className={cn("mb-4 p-3 rounded-xl border", isExportingPDF ? "bg-white border-gray-200" : "bg-gray-50 border-gray-100 print:bg-white print:border-gray-200")}>
                                        <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Опис / Коментар</p>
                                        <p className="text-xs text-gray-600 font-medium leading-relaxed">{p.comment}</p>
                                      </div>
                                    )}

                                    <div className="space-y-3 mb-6 flex-1">
                                        <div className="flex justify-between items-center text-xs no-print">
                                          <span className="text-gray-400 font-bold uppercase tracking-wider">Ціна в Китаї</span>
                                          <span className="font-black text-black">{p.priceYuan.toFixed(2)} ¥</span>
                                        </div>
                                        <div className="flex justify-between items-center text-xs no-print">
                                          <span className="text-gray-400 font-bold uppercase tracking-wider">Ціна в Україні</span>
                                          <span className="font-black text-black">{costPriceUah.toFixed(0)} грн</span>
                                        </div>
                                        <div className="flex justify-between items-center text-xs no-print">
                                          <span className="text-gray-400 font-bold uppercase tracking-wider">Доставка з Китаю</span>
                                          <div className="text-right">
                                            <div className="font-black text-black">{itemDeliveryChinaUah.toFixed(0)} грн</div>
                                          </div>
                                        </div>
                                        <div className="flex justify-between items-center text-xs no-print">
                                          <span className="text-gray-400 font-bold uppercase tracking-wider">Доставка НП</span>
                                          <div className="text-right">
                                            <div className="font-black text-black">{itemDeliveryNPUah.toFixed(0)} грн</div>
                                          </div>
                                        </div>
                                      <div className="flex justify-between items-center text-[10px]">
                                        <span className="text-gray-400 font-bold uppercase tracking-wider">Вага / Об'єм / Щільність</span>
                                        <div className="flex items-center gap-1 font-black text-gray-600">
                                          <input
                                            type="number"
                                            className="w-10 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-right"
                                            defaultValue={p.weight || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.weight) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, weight: val } : item));
                                                addNotification(`Вагу оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span>{p.weightUnit || 'кг'} /</span>
                                          <input
                                            type="number"
                                            className="w-12 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-right"
                                            defaultValue={p.volume || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.volume) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, volume: val } : item));
                                                addNotification(`Об'єм оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span>м³ /</span>
                                          <input
                                            type="number"
                                            className="w-10 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-right"
                                            defaultValue={p.density || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.density) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, density: val } : item));
                                                addNotification(`Щільність оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span>кг/м³</span>
                                        </div>
                                      </div>
                                      {p.localDeliveryId && (
                                        <div className="flex justify-between items-center text-[10px] mt-1">
                                          <span className="text-gray-400 font-bold uppercase tracking-wider">ТТН НП</span>
                                          <span className="font-black text-gray-600">
                                            {localDeliveries.find(ld => ld.id === p.localDeliveryId)?.trackingNumber || '-'}
                                          </span>
                                        </div>
                                      )}
                                      <div className="h-px bg-gray-100 my-2 no-print" />
                                        <div className="flex justify-between items-center no-print">
                                          <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Повна собівартість</span>
                                          <div className="text-right">
                                            <div className="text-xl font-black text-black">{totalCostUah.toFixed(0)} грн</div>
                                            <div className="text-[10px] font-bold text-gray-400">{(totalCostUah / p.exchangeRate).toFixed(2)} ¥</div>
                                          </div>
                                        </div>
                                    </div>

                                    <div className="mt-auto pt-6 border-t border-gray-100">
                                      <div className="flex justify-between items-end">
                                          <div>
                                            <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Роздрібна ціна</p>
                                            <div className="flex items-baseline gap-2">
                                              <span className="text-3xl font-black text-black">{sellingPriceUah.toFixed(0)}</span>
                                              <span className="text-sm font-black text-black">грн</span>
                                            </div>
                                            <p className="text-[10px] font-bold text-gray-400 mt-1 no-print">{sellingPriceYuan.toFixed(2)} ¥</p>
                                          </div>
                                        <div className="flex flex-col items-end gap-2 no-print">
                                          <span className={cn(
                                            "px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest",
                                            p.markup ? "bg-emerald-50 text-emerald-700" : "bg-green-50 text-green-700"
                                          )}>
                                            {p.markup ? 'Індивідуальна' : `${priceListMargin}% націнка`}
                                          </span>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </motion.div>
                              );
                            })}
                        </div>
                      ) : (
                        <div className={cn("bg-white rounded-3xl overflow-hidden shadow-xl border", isExportingPDF ? "m-0 shadow-none border-2 border-gray-300" : "border-gray-200 m-4 print:m-0 print:shadow-none print:border-2 print:border-gray-300")}>
                          <table className="w-full text-left border-collapse border-spacing-0">
                            <thead className={cn(isExportingPDF ? "bg-gray-100 text-black" : "bg-black text-white print:bg-gray-100 print:text-black")}>
                              <tr>
                                <th className={cn("py-3 px-4 text-left w-12 border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>
                                  <button 
                                    onClick={() => {
                                      const filteredIds = purchases
                                        .filter(p => {
                                          const matchesStatus = p.status === 'my_warehouse';
                                          const matchesSearch = !priceListSearch || (p.name || '').toLowerCase().includes(priceListSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(priceListSearch.toLowerCase());
                                          const matchesDate = !priceListDateFilter || (p.arrivalDate && p.arrivalDate.includes(priceListDateFilter)) || (p.createdAt && p.createdAt.includes(priceListDateFilter));
                                          const matchesBatch = !priceListBatchFilter || p.batchId === priceListBatchFilter;
                                          return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                                        })
                                        .map(p => p.id);
                                      toggleSelectAll(filteredIds);
                                    }}
                                    className="text-gray-400 hover:text-white transition-colors print:text-black"
                                  >
                                    {selectedPurchaseIds.length > 0 ? <CheckSquare className="w-5 h-5 text-white print:text-black" /> : <Square className="w-5 h-5" />}
                                  </button>
                                </th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Фото</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Назва / Опис</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Трек / ТТН НП / Платформа</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Вага / Об'єм / Щільність</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Доставка (Китай / НП)</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Собівартість (Китай / Укр)</th>
                                <th className={cn("py-3 px-4 text-[9px] font-black uppercase tracking-widest border text-right", isExportingPDF ? "border-gray-300" : "border-white/10 print:border-gray-300")}>Ціна (грн)</th>
                              </tr>
                            </thead>
                            <tbody>
                              {purchases
                                .filter(p => {
                                  const matchesStatus = p.status === 'my_warehouse';
                                  const matchesSearch = !priceListSearch || (p.name || '').toLowerCase().includes(priceListSearch.toLowerCase()) || (p.trackNumber || '').toLowerCase().includes(priceListSearch.toLowerCase());
                                  const matchesDate = !priceListDateFilter || (p.arrivalDate && p.arrivalDate.includes(priceListDateFilter)) || (p.createdAt && p.createdAt.includes(priceListDateFilter));
                                  const matchesBatch = !priceListBatchFilter || p.batchId === priceListBatchFilter;
                                  return matchesStatus && matchesSearch && matchesDate && matchesBatch;
                                })
                                .map(p => {
                                  const costPriceYuan = p.priceYuan * p.quantity;
                                  const costPriceUah = costPriceYuan * p.exchangeRate;
                                  
                                  const trackPurchases = p.trackNumber ? purchases.filter(item => item.trackNumber === p.trackNumber) : [p];
                                  let totalDeliveryForGroup = 0;
                                  let totalWeightForGroup = 0;
                                  let totalDeliveryChinaUah = 0;
                                  let totalDeliveryNPUah = 0;
                                  
                                  trackPurchases.forEach(item => {
                                    const deliveryChinaUah = (item.deliveryCostPerItem || 0) * usdToUah; 
                                    const deliveryIntUah = (item.shippingCost || 0) * usdToUah; 
                                    const deliveryUAUah = (item.ukraineDeliveryCost || 0); 
                                    const deliveryNPUah = (item.novaPoshtaCost || 0) + getLocalDeliveryCost(item); 
                                    totalDeliveryChinaUah += deliveryChinaUah + deliveryIntUah;
                                    totalDeliveryNPUah += deliveryUAUah + deliveryNPUah;
                                    totalDeliveryForGroup += deliveryChinaUah + deliveryIntUah + deliveryUAUah + deliveryNPUah;
                                    const actualWeight = item.weightUnit === 'g' ? (item.weight || 0) / 1000 : (item.weight || 0);
                                    totalWeightForGroup += actualWeight;
                                  });

                                  const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
                                  let itemDeliveryUah = 0;
                                  let itemDeliveryChinaUah = 0;
                                  let itemDeliveryNPUah = 0;

                                  if (trackPurchases.length === 1) {
                                    itemDeliveryUah = totalDeliveryForGroup;
                                    itemDeliveryChinaUah = totalDeliveryChinaUah;
                                    itemDeliveryNPUah = totalDeliveryNPUah;
                                  } else if (totalWeightForGroup > 0) {
                                    const ratio = actualWeight / totalWeightForGroup;
                                    itemDeliveryUah = totalDeliveryForGroup * ratio;
                                    itemDeliveryChinaUah = totalDeliveryChinaUah * ratio;
                                    itemDeliveryNPUah = totalDeliveryNPUah * ratio;
                                  } else {
                                    const ratio = 1 / trackPurchases.length;
                                    itemDeliveryUah = totalDeliveryForGroup * ratio;
                                    itemDeliveryChinaUah = totalDeliveryChinaUah * ratio;
                                    itemDeliveryNPUah = totalDeliveryNPUah * ratio;
                                  }
                                  
                                  const totalCostUah = costPriceUah + itemDeliveryUah;
                                  const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));

                                  const dimString = p.width ? `${p.width}x${p.height}x${p.length} ${p.dimUnit}` : p.size || '-';
                                  const weightString = p.weight ? `${p.weight} ${p.weightUnit}` : '-';
                                  const volumeString = p.volume ? `${p.volume.toFixed(3)} м³` : '-';
                                  const densityString = p.density ? `${p.density.toFixed(0)} кг/м³` : '-';

                                  return (
                                    <tr key={p.id} className={cn("border-b hover:bg-gray-50 transition-colors group", isExportingPDF ? "border-gray-300" : "border-gray-200 print:border-gray-300", selectedPurchaseIds.includes(p.id) && "bg-yellow-50/30")}>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <button 
                                          onClick={() => toggleSelectOne(p.id)}
                                          className="text-gray-400 hover:text-black transition-colors print:hidden"
                                        >
                                          {selectedPurchaseIds.includes(p.id) ? <CheckSquare className="w-5 h-5 text-black" /> : <Square className="w-5 h-5" />}
                                        </button>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <div className="w-12 h-12 rounded-lg overflow-hidden bg-gray-100 border border-gray-200">
                                          {p.photo ? (
                                            <img src={p.photo} alt={p.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                                          ) : (
                                            <div className="w-full h-full flex items-center justify-center">
                                              <Package className="w-4 h-4 text-gray-200" />
                                            </div>
                                          )}
                                        </div>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <p className="text-[11px] font-black text-black mb-0.5">{p.name}</p>
                                        {p.comment && (
                                          <p className="text-[9px] text-gray-500 font-medium leading-tight max-w-[150px]">{p.comment}</p>
                                        )}
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <p className="text-[9px] text-gray-600 font-bold">{p.trackNumber}</p>
                                        {p.localDeliveryId && (
                                          <p className="text-[9px] text-blue-500 font-bold">
                                            НП: {localDeliveries.find(ld => ld.id === p.localDeliveryId)?.trackingNumber || '-'}
                                          </p>
                                        )}
                                        <p className="text-[8px] text-gray-400 uppercase tracking-widest">{p.platform}</p>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <div className="flex items-center gap-1 group mb-1">
                                          <input
                                            key={p.weight}
                                            type="number"
                                            className="w-12 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-[10px] font-black text-black text-right"
                                            defaultValue={p.weight || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.weight) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, weight: val } : item));
                                                addNotification(`Вагу оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span className="text-[10px] font-black text-black">{p.weightUnit || 'кг'}</span>
                                        </div>
                                        <div className="flex items-center gap-1 group mb-1">
                                          <input
                                            key={p.volume}
                                            type="number"
                                            className="w-12 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-[9px] font-bold text-gray-500 text-right"
                                            defaultValue={p.volume || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.volume) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, volume: val } : item));
                                                addNotification(`Об'єм оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span className="text-[9px] font-bold text-gray-500">м³</span>
                                        </div>
                                        <div className="flex items-center gap-1 group">
                                          <input
                                            key={p.density}
                                            type="number"
                                            className="w-12 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-[9px] font-bold text-gray-500 text-right"
                                            defaultValue={p.density || ''}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== p.density) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, density: val } : item));
                                                addNotification(`Щільність оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span className="text-[9px] font-bold text-gray-500">кг/м³</span>
                                        </div>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <p className="text-[10px] font-black text-black">{itemDeliveryUah.toFixed(0)} грн (Всього)</p>
                                        <p className="text-[9px] text-gray-500 font-bold">Китай: {itemDeliveryChinaUah.toFixed(0)} грн</p>
                                        <p className="text-[9px] text-gray-500 font-bold">НП: {itemDeliveryNPUah.toFixed(0)} грн</p>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <p className="text-[10px] font-black text-black">{costPriceYuan.toFixed(2)} ¥ (Китай)</p>
                                        <p className="text-[10px] font-black text-black">{costPriceUah.toFixed(0)} грн (Укр)</p>
                                        <p className="text-[10px] font-black text-black">{totalCostUah.toFixed(0)} грн (Повна)</p>
                                      </td>
                                      <td className={cn("py-3 px-4 align-top border text-right", isExportingPDF ? "border-gray-300" : "border-gray-100 print:border-gray-300")}>
                                        <div className="flex items-center justify-end gap-1 group">
                                          <input
                                            type="number"
                                            className="w-20 bg-transparent border-b border-transparent group-hover:border-gray-300 focus:border-black focus:outline-none transition-colors text-sm font-black text-black text-right"
                                            defaultValue={sellingPriceUah.toFixed(0)}
                                            onBlur={(e) => {
                                              const val = parseFloat(e.target.value);
                                              if (!isNaN(val) && val !== sellingPriceUah) {
                                                setPurchases(purchases.map(item => item.id === p.id ? { ...item, sellingPrice: val, markup: true } : item));
                                                addNotification(`Ціну продажу оновлено`, 'success');
                                              }
                                            }}
                                          />
                                          <span className="text-sm font-black text-black">грн</span>
                                        </div>
                                        <p className="text-[9px] text-gray-400 font-bold">{(sellingPriceUah / p.exchangeRate).toFixed(2)} ¥</p>
                                        <p className="text-[9px] text-emerald-500 font-bold mt-1">{p.markup ? 'Індивідуальна' : `Націнка: ${priceListMargin}%`}</p>
                                      </td>
                                    </tr>
                                  );
                                })}
                            </tbody>
                          </table>
                        </div>
                      )}
                      
                      {purchases.filter(p => p.status === 'my_warehouse').length === 0 && (
                        <div className="bg-white rounded-3xl p-20 text-center border-2 border-dashed border-gray-100">
                          <Package className="w-16 h-16 text-gray-200 mx-auto mb-6" />
                          <h3 className="text-xl font-black text-gray-400 uppercase tracking-tight">Склад порожній</h3>
                          <p className="text-gray-400 text-sm font-bold mt-2">Додайте товари з Нової Пошти, щоб сформувати прайс</p>
                        </div>
                      )}
                    </div>
                  ) : crmModule === 'pinduoduo' ? (
                    <div className="flex flex-col items-center justify-center space-y-8 py-10">
                      <div className="text-center mb-4">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Pinduoduo Simulator</h2>
                        <p className="text-gray-400 text-sm font-bold">Мобільна версія маркетплейсу в браузері</p>
                      </div>

                      {/* Phone Frame */}
                      <div className="relative w-[480px] h-[880px] bg-black rounded-[4rem] border-[12px] border-gray-900 shadow-2xl overflow-hidden flex flex-col">
                        {/* Navigation Bar */}
                        <div className="h-16 bg-gray-900 flex items-center justify-between px-6 z-20 border-b border-white/5">
                          <div className="flex items-center gap-2">
                            <button 
                              onClick={() => {
                                try {
                                  const frame = document.getElementById('pinduoduo-frame') as HTMLIFrameElement;
                                  frame.contentWindow?.history.back();
                                } catch (e) {
                                  console.log("Navigation restricted by browser security");
                                }
                              }}
                              className="p-2 hover:bg-white/10 rounded-lg transition-colors text-white/60 hover:text-white"
                              title="Назад"
                            >
                              <ChevronLeft className="w-5 h-5" />
                            </button>
                            <button 
                              onClick={() => {
                                try {
                                  const frame = document.getElementById('pinduoduo-frame') as HTMLIFrameElement;
                                  frame.contentWindow?.history.forward();
                                } catch (e) {
                                  console.log("Navigation restricted by browser security");
                                }
                              }}
                              className="p-2 hover:bg-white/10 rounded-lg transition-colors text-white/60 hover:text-white"
                              title="Вперед"
                            >
                              <ChevronRight className="w-5 h-5" />
                            </button>
                            <button 
                              onClick={() => {
                                const frame = document.getElementById('pinduoduo-frame') as HTMLIFrameElement;
                                if (frame) frame.src = frame.src;
                              }}
                              className="p-2 hover:bg-white/10 rounded-lg transition-colors text-white/60 hover:text-white"
                              title="Оновити"
                            >
                              <RotateCcw className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => {
                                const frame = document.getElementById('pinduoduo-frame') as HTMLIFrameElement;
                                if (frame) frame.src = "https://mobile.yangkeduo.com/";
                              }}
                              className="p-2 hover:bg-white/10 rounded-lg transition-colors text-white/60 hover:text-white"
                              title="На головну"
                            >
                              <Home className="w-4 h-4" />
                            </button>
                          </div>
                          
                          <div className="flex-1 mx-4">
                            <div className="bg-white/5 rounded-full px-4 py-1.5 text-[10px] text-white/30 truncate font-mono text-center border border-white/5">
                              mobile.yangkeduo.com
                            </div>
                          </div>

                          <div className="flex items-center gap-2">
                            <div className="w-2 h-2 bg-green-500 rounded-full animate-pulse shadow-[0_0_10px_rgba(34,197,94,0.5)]" />
                            <span className="text-[9px] font-black text-white/40 uppercase tracking-[0.2em]">Online</span>
                          </div>
                        </div>
                        
                        {/* Iframe Container */}
                        <div className="flex-1 bg-white relative z-10">
                          <iframe 
                            id="pinduoduo-frame"
                            src="https://mobile.yangkeduo.com/" 
                            className="w-full h-full border-none"
                            title="Pinduoduo Mobile"
                          />
                        </div>

                        {/* Home Bar */}
                        <div className="h-6 bg-gray-900 flex items-center justify-center z-20">
                          <div className="w-24 h-1 bg-white/20 rounded-full" />
                        </div>
                      </div>

                      <div className="max-w-md text-center space-y-4">
                        <div className="bg-yellow-50 border border-yellow-200 p-4 rounded-2xl">
                          <p className="text-xs text-yellow-800 font-bold leading-relaxed">
                            <strong>Порада:</strong> Через обмеження безпеки браузера (Cross-Origin), кнопки "Назад" та "Вперед" не можуть керувати історією всередині сайту Pinduoduo. Використовуйте навігацію безпосередньо в інтерфейсі додатка або кнопку "На головну".
                          </p>
                        </div>
                        <a 
                          href="https://mobile.yangkeduo.com/" 
                          target="_blank" 
                          rel="noreferrer"
                          className="inline-flex items-center gap-2 bg-black text-white px-8 py-4 rounded-2xl font-black uppercase tracking-widest text-xs hover:bg-gray-800 transition-all"
                        >
                          <ExternalLink className="w-4 h-4" />
                          Відкрити в новому вікні
                        </a>
                      </div>
                    </div>
                  ) : crmModule === 'settings' ? (
                    <div className="space-y-8">
                      <div className="flex justify-between items-center">
                        <h2 className="text-2xl font-black text-black uppercase tracking-tight">Налаштування системи</h2>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                        <div className="bg-white p-8 rounded-3xl shadow-xl border border-gray-100">
                          <div className="flex items-center gap-4 mb-6">
                            <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center">
                              <DollarSign className="w-6 h-6 text-black" />
                            </div>
                            <div>
                              <h3 className="text-lg font-black text-black uppercase tracking-tight">Курси валют</h3>
                              <p className="text-gray-400 text-xs font-bold">Фіксовані курси для розрахунків</p>
                            </div>
                          </div>

                          <div className="space-y-6">
                            <div className="space-y-2">
                              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Курс Юань до Гривні (CNY/UAH)</label>
                              <div className="flex items-center gap-4">
                                <input 
                                  type="number" 
                                  step="0.01"
                                  value={cnyToUah}
                                  onChange={(e) => setCnyToUah(parseFloat(e.target.value) || 0)}
                                  className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-black text-2xl text-black focus:outline-none focus:ring-2 focus:ring-black"
                                />
                                <div className="text-gray-400 font-black text-xl">UAH</div>
                              </div>
                              <p className="text-[10px] text-gray-400 font-bold italic">Цей курс використовується для розрахунку собівартості та формування прайс-листа.</p>
                            </div>

                            <div className="space-y-2">
                              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Курс Долара до Гривні (USD/UAH)</label>
                              <div className="flex items-center gap-4">
                                <input 
                                  type="number" 
                                  step="0.01"
                                  value={usdToUah}
                                  onChange={(e) => setUsdToUah(parseFloat(e.target.value) || 0)}
                                  className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-black text-2xl text-black focus:outline-none focus:ring-2 focus:ring-black"
                                />
                                <div className="text-gray-400 font-black text-xl">UAH</div>
                              </div>
                              <p className="text-[10px] text-gray-400 font-bold italic">Використовується для розрахунку вартості доставки в гривні.</p>
                            </div>
                          </div>
                        </div>

                        <div className="bg-white p-8 rounded-3xl shadow-xl border border-gray-100">
                          <div className="flex items-center justify-between mb-6">
                            <div className="flex items-center gap-4">
                              <div className="w-12 h-12 bg-blue-50 rounded-2xl flex items-center justify-center">
                                <Truck className="w-6 h-6 text-black" />
                              </div>
                              <div>
                                <h3 className="text-lg font-black text-black uppercase tracking-tight">Тарифи доставки</h3>
                                <p className="text-gray-400 text-xs font-bold">Налаштування калькулятора</p>
                              </div>
                            </div>
                            <button 
                              onClick={() => {
                                setTariffForm({
                                  name: '',
                                  iconName: 'Ship',
                                  deliveryDays: '',
                                  description: '',
                                  pricePerKg: 0,
                                  volumetricFactor: 0,
                                  localDeliveryPrice: 0,
                                  minWeight: 0,
                                  minVolume: 0,
                                  minCost: 0,
                                  insuranceRate: 2,
                                  packagingCost: 0,
                                  packagingCostPerM3: 0,
                                  customsFee: 0,
                                  handlingFee: 0,
                                  fuelSurcharge: 0,
                                  densityTiers: []
                                });
                                setShowTariffModal({ show: true, tariffId: null });
                              }}
                              className="p-2 bg-black text-white rounded-lg hover:bg-gray-800 transition-all"
                            >
                              <Plus className="w-4 h-4" />
                            </button>
                          </div>

                          <div className="space-y-4">
                            {tariffs.map(t => (
                              <div key={t.id} className="flex items-center justify-between p-4 bg-gray-50 rounded-2xl border border-gray-100">
                                <div className="flex items-center gap-3">
                                  {getTariffIcon(t.iconName)}
                                  <div>
                                    <p className="text-sm font-black text-black uppercase tracking-tight">{t.name}</p>
                                    <p className="text-[10px] text-gray-400 font-bold">{t.deliveryDays}</p>
                                  </div>
                                </div>
                                <div className="flex gap-2">
                                  <button 
                                    onClick={() => {
                                      setTariffForm(t);
                                      setShowTariffModal({ show: true, tariffId: t.id });
                                    }}
                                    className="p-2 text-gray-400 hover:text-black transition-colors"
                                  >
                                    <Edit2 className="w-4 h-4" />
                                  </button>
                                  <button 
                                    onClick={() => {
                                      if (tariffs.length > 1) {
                                        setConfirmModal({
                                          show: true,
                                          title: 'Видалити тариф?',
                                          message: `Ви впевнені, що хочете видалити ${t.name}?`,
                                          onConfirm: () => setTariffs(prev => prev.filter(item => item.id !== t.id))
                                        });
                                      } else {
                                        addNotification('Неможливо видалити останній тариф', 'error');
                                      }
                                    }}
                                    className="p-2 text-red-400 hover:text-red-600 transition-colors"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                  </button>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="bg-white p-8 rounded-3xl shadow-xl border border-gray-100">
                          <div className="flex items-center gap-4 mb-6">
                            <div className="w-12 h-12 bg-red-50 rounded-2xl flex items-center justify-center">
                              <Store className="w-6 h-6 text-black" />
                            </div>
                            <div>
                              <h3 className="text-lg font-black text-black uppercase tracking-tight">Інтеграції</h3>
                              <p className="text-gray-400 text-xs font-bold">Налаштування зовнішніх сервісів</p>
                            </div>
                          </div>

                          <div className="space-y-6">
                            <div className="space-y-2">
                              <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">API Ключ Нової Пошти</label>
                              <div className="flex items-center gap-4">
                                <input 
                                  type="password" 
                                  value={novaPoshtaApiKey}
                                  onChange={(e) => setNovaPoshtaApiKey(e.target.value)}
                                  placeholder="Введіть API ключ..."
                                  className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold text-sm text-black focus:outline-none focus:ring-2 focus:ring-black"
                                />
                              </div>
                              <p className="text-[10px] text-gray-400 font-bold italic">Використовується для автоматичного створення ТТН та розрахунку вартості доставки.</p>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  ) : null}
                </div>
              </motion.div>
            </div>
          </div>
        </main>
      )}

      {/* Footer */}
      <footer className="bg-black py-16 text-center text-white relative overflow-hidden">
        <div className="absolute inset-0 opacity-5 pointer-events-none">
          <div className="absolute top-0 left-0 w-full h-full bg-[radial-gradient(circle_at_center,_var(--tw-gradient-stops))] from-white via-transparent to-transparent" />
        </div>
        <div className="max-w-7xl mx-auto px-6 relative z-10">
          <div className="flex items-center justify-center gap-3 mb-6">
            <div className="w-10 h-10 bg-yellow-400 rounded-lg flex items-center justify-center rotate-3">
              <Calculator className="w-6 h-6 text-black" />
            </div>
            <span className="text-3xl font-black italic tracking-tighter text-white">FORSAGE<span className="text-yellow-400"> CHINA</span></span>
          </div>
          <p className="text-white/40 text-[10px] font-black uppercase tracking-[0.3em]">System Version 2.4.0</p>
        </div>
      </footer>

      {/* CRM Modals */}
      {selectedTrackNumber && (
        <div className={cn("fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6", isExportingPDF ? "absolute inset-0 block bg-white" : "print:absolute print:inset-0 print:block print:bg-white")}>
          <div className={cn("absolute inset-0 bg-black/90 backdrop-blur-sm", isExportingPDF ? "hidden" : "print:hidden")} onClick={() => setSelectedTrackNumber(null)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className={cn("bg-white w-full max-w-4xl overflow-hidden flex flex-col relative z-10", isExportingPDF ? "shadow-none max-h-none overflow-visible rounded-none" : "rounded-[2rem] max-h-[90vh] shadow-2xl print:shadow-none print:max-h-none print:overflow-visible print:rounded-none")}
          >
            <div className={cn("p-6 flex justify-between items-center", isExportingPDF ? "bg-white border-b-2 border-black" : "border-b border-gray-100 bg-gray-50/50 print:bg-white print:border-b-2 print:border-black")}>
              <div>
                <h3 className="text-xl font-black text-black">Трек-номер: {selectedTrackNumber}</h3>
                <p className="text-sm text-gray-500 font-medium mt-1">Товари в цьому відправленні</p>
              </div>
              <div className="flex items-center gap-2 no-print">
                <button 
                  onClick={() => {
                    const trackPurchases = purchases.filter(p => p.trackNumber === selectedTrackNumber);
                    exportToExcelWithPhotos(trackPurchases, `Track_${selectedTrackNumber}`);
                  }}
                  className="bg-emerald-500 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-emerald-600 transition-all shadow-lg shadow-emerald-100 flex items-center gap-2"
                  title="Експорт в Excel"
                >
                  <Download className="w-4 h-4" />
                  Excel
                </button>
                <button 
                  onClick={() => handleExportTrackPDF(selectedTrackNumber)}
                  disabled={isExportingPDF}
                  className={cn(
                    "bg-yellow-400 text-black px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-yellow-500 transition-all shadow-lg shadow-yellow-100 flex items-center gap-2",
                    isExportingPDF && "opacity-50 cursor-not-allowed"
                  )}
                  title="Експорт в PDF"
                >
                  <FileText className="w-4 h-4" />
                  PDF
                </button>
                <button 
                  onClick={() => window.print()}
                  className="bg-black text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-gray-800 transition-all shadow-lg flex items-center gap-2 no-print"
                  title="Друк"
                >
                  <Printer className="w-4 h-4" />
                  Друк
                </button>
                <button 
                  onClick={() => setSelectedTrackNumber(null)}
                  className="p-2 hover:bg-gray-200 rounded-full transition-colors ml-2"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
            </div>
            
            <div id="track-modal-content" className={cn("flex-1", isExportingPDF ? "overflow-visible bg-white p-0 mt-6" : "p-6 overflow-y-auto bg-gray-50/30 print:overflow-visible print:bg-white print:p-0 print:mt-6")}>
              <div className={cn(isExportingPDF ? "space-y-6" : "space-y-4 print:space-y-6")}>
                {(() => {
                  const trackPurchases = purchases.filter(p => p.trackNumber === selectedTrackNumber);
                  let totalDeliveryForGroup = 0;
                  let totalWeightForGroup = 0;
                  
                  trackPurchases.forEach(p => {
                    const deliveryChinaUah = (p.deliveryCostPerItem || 0) * usdToUah; 
                    const deliveryIntUah = (p.shippingCost || 0) * usdToUah; 
                    const deliveryUAUah = (p.ukraineDeliveryCost || 0); 
                    const deliveryNPUah = (p.novaPoshtaCost || 0) + getLocalDeliveryCost(p); 
                    totalDeliveryForGroup += deliveryChinaUah + deliveryIntUah + deliveryUAUah + deliveryNPUah;
                    const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
                    totalWeightForGroup += actualWeight;
                  });

                  return trackPurchases.map(p => {
                    const costPriceYuan = p.priceYuan * p.quantity;
                    const costPriceUah = costPriceYuan * p.exchangeRate;
                    const actualWeight = p.weightUnit === 'g' ? (p.weight || 0) / 1000 : (p.weight || 0);
                    
                    let itemDeliveryUah = 0;
                    if (trackPurchases.length === 1) {
                      itemDeliveryUah = totalDeliveryForGroup;
                    } else if (totalWeightForGroup > 0) {
                      itemDeliveryUah = totalDeliveryForGroup * (actualWeight / totalWeightForGroup);
                    } else {
                      itemDeliveryUah = totalDeliveryForGroup / trackPurchases.length;
                    }

                    const totalCostUah = costPriceUah + itemDeliveryUah;
                    const sellingPriceUah = p.markup ? (p.sellingPrice || totalCostUah) : (totalCostUah * (1 + priceListMargin / 100));
                    const dimString = p.width ? `${p.width}x${p.height}x${p.length} ${p.dimUnit}` : p.size || '';

                    return (
                      <div key={p.id} className={cn("bg-white rounded-2xl p-4 flex gap-4 items-center transition-shadow", isExportingPDF ? "border-2 border-gray-200 shadow-none break-inside-avoid" : "border border-gray-100 shadow-sm hover:shadow-md print:border-2 print:border-gray-200 print:shadow-none print:break-inside-avoid")}>
                        <div className={cn("w-20 h-20 rounded-xl overflow-hidden bg-gray-50 flex-shrink-0 border", isExportingPDF ? "border-gray-200" : "border-gray-100 print:border-gray-200")}>
                          {p.photo ? (
                            <img src={p.photo} alt={p.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                          ) : (
                            <div className="w-full h-full flex items-center justify-center">
                              <Package className="w-8 h-8 text-gray-300" />
                            </div>
                          )}
                        </div>
                        
                        <div className="flex-1 min-w-0">
                          <h4 className={cn("font-bold text-black text-base", isExportingPDF ? "whitespace-normal" : "truncate print:whitespace-normal")}>{p.name}</h4>
                          <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2">
                            <div className="flex items-center gap-1 text-xs text-gray-500">
                              <span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Ціна:</span>
                              <input
                                type="number"
                                className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                defaultValue={p.priceYuan || 0}
                                onBlur={(e) => {
                                  const val = parseFloat(e.target.value);
                                  if (!isNaN(val) && val !== p.priceYuan) {
                                    setPurchases(purchases.map(item => item.id === p.id ? { ...item, priceYuan: val } : item));
                                    addNotification(`Ціну оновлено`, 'success');
                                  }
                                }}
                              />
                              ¥
                            </div>
                            <div className="flex items-center gap-1 text-xs text-gray-500">
                              <span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">К-сть:</span>
                              <input
                                type="number"
                                className="w-12 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                defaultValue={p.quantity || 1}
                                onBlur={(e) => {
                                  const val = parseInt(e.target.value, 10);
                                  if (!isNaN(val) && val !== p.quantity && val > 0) {
                                    setPurchases(purchases.map(item => item.id === p.id ? { ...item, quantity: val } : item));
                                    addNotification(`Кількість оновлено`, 'success');
                                  }
                                }}
                              />
                              шт
                            </div>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Статус:</span> {statusLabels[p.status]}</p>
                            <div className="flex items-center gap-1 text-xs text-gray-500">
                              <span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Вага:</span>
                              <input
                                type="number"
                                className="w-12 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                defaultValue={p.weight || ''}
                                onBlur={(e) => {
                                  const val = parseFloat(e.target.value);
                                  if (!isNaN(val) && val !== p.weight) {
                                    setPurchases(purchases.map(item => item.id === p.id ? { ...item, weight: val } : item));
                                    addNotification(`Вагу оновлено`, 'success');
                                  }
                                }}
                              />
                              {p.weightUnit || 'кг'}
                            </div>
                            <div className="flex items-center gap-1 text-xs text-gray-500">
                              <span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Об'єм:</span>
                              <input
                                type="number"
                                className="w-16 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                defaultValue={p.volume || ''}
                                onBlur={(e) => {
                                  const val = parseFloat(e.target.value);
                                  if (!isNaN(val) && val !== p.volume) {
                                    setPurchases(purchases.map(item => item.id === p.id ? { ...item, volume: val } : item));
                                    addNotification(`Об'єм оновлено`, 'success');
                                  }
                                }}
                              />
                              м³
                            </div>
                            <div className="flex items-center gap-1 text-xs text-gray-500">
                              <span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Щільність:</span>
                              <input
                                type="number"
                                className="w-12 bg-transparent border-b border-transparent hover:border-gray-300 focus:border-black focus:outline-none transition-colors"
                                defaultValue={p.density || ''}
                                onBlur={(e) => {
                                  const val = parseFloat(e.target.value);
                                  if (!isNaN(val) && val !== p.density) {
                                    setPurchases(purchases.map(item => item.id === p.id ? { ...item, density: val } : item));
                                    addNotification(`Щільність оновлено`, 'success');
                                  }
                                }}
                              />
                              кг/м³
                            </div>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Доставка Китай:</span> {Math.round((p.deliveryCostPerItem || 0) * usdToUah + (p.shippingCost || 0) * usdToUah + (p.ukraineDeliveryCost || 0))} ₴</p>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Доставка НП:</span> {Math.round((p.novaPoshtaCost || 0) + getLocalDeliveryCost(p))} ₴</p>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Собівартість:</span> {Math.round(totalCostUah)} ₴</p>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Націнка:</span> {p.markup ? 'Індивідуальна' : `${priceListMargin}%`}</p>
                            <p className="text-xs text-gray-500"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Ціна продажу:</span> {Math.round(sellingPriceUah)} ₴</p>
                          </div>
                          {p.comment && <p className="text-xs text-gray-500 mt-1 italic"><span className="font-bold text-gray-400 uppercase tracking-widest text-[10px]">Коментар:</span> {p.comment}</p>}
                        </div>
                        
                        <div className="flex items-center gap-2 flex-shrink-0 no-print">
                          <button 
                            onClick={() => {
                              setSelectedTrackNumber(null);
                              handleEditPurchase(p);
                            }}
                            className="p-2 text-gray-400 hover:text-blue-500 hover:bg-blue-50 rounded-xl transition-colors"
                            title="Редагувати"
                          >
                            <Edit2 className="w-5 h-5" />
                          </button>
                          <button 
                            onClick={() => handleDeletePurchase(p.id)}
                            className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-xl transition-colors"
                            title="Видалити"
                          >
                            <Trash2 className="w-5 h-5" />
                          </button>
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
            </div>
            
            <div className="p-6 border-t border-gray-100 bg-white no-print">
              <button 
                onClick={() => {
                  const trackPurchases = purchases.filter(p => p.trackNumber === selectedTrackNumber);
                  const basePurchase = trackPurchases[0];
                  if (basePurchase) {
                    handleAddToTrack(basePurchase);
                  } else {
                    setPurchaseForm({
                      platform: 'Taobao',
                      name: '',
                      link: '',
                      priceYuan: 0,
                      exchangeRate: 5.5,
                      quantity: 1,
                      trackNumber: selectedTrackNumber,
                      photo: '',
                      comment: '',
                      size: '',
                      width: 0,
                      height: 0,
                      length: 0,
                      dimUnit: 'cm',
                      weight: 0,
                      weightUnit: 'kg',
                      volume: 0,
                      density: 0,
                      isFabric: false,
                      isPressed: false,
                      isInsured: false,
                      declaredValue: 0,
                      shippingCost: 0,
                      status: 'shipped_by_seller',
                      brand: '',
                      radius: '',
                      season: '',
                      article: ''
                    });
                    setShowAddPurchaseModal(true);
                  }
                  setSelectedTrackNumber(null);
                }}
                className="w-full py-4 bg-black text-white rounded-xl font-bold hover:bg-gray-800 transition-colors flex items-center justify-center gap-2"
              >
                <Plus className="w-5 h-5" />
                Додати товар до цього трек-номеру
              </button>
            </div>
          </motion.div>
        </div>
      )}
      {cropImage && (
        <CropModal 
          image={cropImage} 
          onCropComplete={(cropped) => {
            setPurchaseForm({ ...purchaseForm, photo: cropped });
            setCropImage(null);
          }}
          onCancel={() => setCropImage(null)}
        />
      )}
      <BulkAddItemsModal 
        show={showBulkAddModal} 
        onClose={() => setShowBulkAddModal(false)} 
        onSave={handleBulkSave} 
      />
      {showAddPurchaseModal && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowAddPurchaseModal(false)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-6xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => {
              setShowAddPurchaseModal(false);
              setEditingPurchaseId(null);
              setPurchaseForm({
                platform: 'Taobao',
                name: '',
                link: '',
                priceYuan: 0,
                exchangeRate: 5.5,
                quantity: 1,
                trackNumber: '',
                photo: '',
                comment: '',
                size: '',
                width: 0,
                height: 0,
                length: 0,
                dimUnit: 'cm',
                weight: 0,
                weightUnit: 'kg',
                volume: 0,
                density: 0,
                isFabric: false,
                isPressed: false,
                isInsured: false,
                declaredValue: 0,
                shippingCost: 0,
                status: 'purchased',
                brand: '',
                radius: '',
                season: '',
                article: ''
              });
            }} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-[#003d2b] uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <PlusCircle className="w-6 h-6 md:w-8 md:h-8 text-black" />
              Додати нову закупку
            </h3>
            
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Платформа / Постачальник</label>
                <select 
                  value={purchaseForm.platform}
                  onChange={(e) => setPurchaseForm({...purchaseForm, platform: e.target.value})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                >
                  <option value="Taobao">Taobao</option>
                  <option value="Pinduoduo">Pinduoduo</option>
                  <option value="1688">1688</option>
                  <option value="Інше">Інше</option>
                </select>
              </div>
              
              <div className="space-y-2 lg:col-span-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва товару</label>
                <input 
                  type="text" 
                  value={purchaseForm.name}
                  onChange={(e) => setPurchaseForm({...purchaseForm, name: e.target.value})}
                  placeholder="Назва товару" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>

              <div className="space-y-2 lg:col-span-3">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Посилання на товар</label>
                <input 
                  type="text" 
                  value={purchaseForm.link}
                  onChange={(e) => setPurchaseForm({...purchaseForm, link: e.target.value})}
                  placeholder="https://..." 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна за одиницю (¥)</label>
                <input 
                  type="number" 
                  value={purchaseForm.priceYuan || ''}
                  onChange={(e) => setPurchaseForm({...purchaseForm, priceYuan: parseFloat(e.target.value) || 0})}
                  placeholder="0.00" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Курс юаня (CNY/UAH)</label>
                <input 
                  type="number" 
                  step="0.01"
                  value={purchaseForm.exchangeRate || ''}
                  onChange={(e) => setPurchaseForm({...purchaseForm, exchangeRate: parseFloat(e.target.value) || 0})}
                  placeholder={cnyToUah.toString()} 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна в грн (Авто)</label>
                <div className="w-full p-4 bg-gray-100 rounded-xl border border-gray-100 font-black text-black">
                  {(purchaseForm.priceYuan * purchaseForm.exchangeRate).toFixed(2)} грн
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Кількість</label>
                <input 
                  type="number" 
                  value={purchaseForm.quantity || ''}
                  onChange={(e) => setPurchaseForm({...purchaseForm, quantity: parseInt(e.target.value) || 0})}
                  placeholder="1" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Загальна ціна</label>
                <div className="w-full p-4 bg-gray-100 rounded-xl border border-gray-100 font-black flex justify-between items-center">
                  <span className="text-black">{(purchaseForm.priceYuan * purchaseForm.quantity).toFixed(2)} ¥</span>
                  <span className="text-black">{((purchaseForm.priceYuan * purchaseForm.quantity) * purchaseForm.exchangeRate).toFixed(2)} грн</span>
                </div>
              </div>

              {/* Advanced Calculation Section */}
              <div className="lg:col-span-3 border-t border-gray-100 pt-6 mt-2">
                <h4 className="text-xs font-black text-[#003d2b] uppercase tracking-widest mb-4 flex items-center gap-2">
                  <Calculator className="w-4 h-4 text-black" />
                  Параметри доставки
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  <div className="space-y-2 lg:col-span-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Розміри (Ш x В x Д) в см</label>
                    <div className="flex gap-2">
                      <input 
                        type="number" 
                        value={purchaseForm.width || ''}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value) || 0;
                          const newForm = { ...purchaseForm, width: val };
                          
                          // Recalculate shipping cost
                          const w = newForm.weight || 0;
                          const actualWeight = newForm.weightUnit === 'kg' ? w : w / 1000;
                          
                          const factor = newForm.dimUnit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          // If dimensions are present, use them. Otherwise use volume/density if present.
                          let finalWeight = Math.max(actualWeight, volWeight);
                          
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            const volWeightFromVol = newForm.volume * 200; // Standard volumetric weight factor for m3
                            finalWeight = Math.max(actualWeight, volWeightFromVol);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5; // Example fabric surcharge
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        placeholder="Ш (см)" 
                        className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                      <input 
                        type="number" 
                        value={purchaseForm.height || ''}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value) || 0;
                          const newForm = { ...purchaseForm, height: val };
                          
                          const w = newForm.weight || 0;
                          const actualWeight = newForm.weightUnit === 'kg' ? w : w / 1000;
                          
                          const factor = newForm.dimUnit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          let finalWeight = Math.max(actualWeight, volWeight);
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            finalWeight = Math.max(actualWeight, newForm.volume * 200);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        placeholder="В (см)" 
                        className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                      <input 
                        type="number" 
                        value={purchaseForm.length || ''}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value) || 0;
                          const newForm = { ...purchaseForm, length: val };
                          
                          const w = newForm.weight || 0;
                          const actualWeight = newForm.weightUnit === 'kg' ? w : w / 1000;
                          
                          const factor = newForm.dimUnit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          let finalWeight = Math.max(actualWeight, volWeight);
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            finalWeight = Math.max(actualWeight, newForm.volume * 200);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        placeholder="Д (см)" 
                        className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                      <select 
                        value={purchaseForm.dimUnit}
                        onChange={(e) => {
                          const unit = e.target.value as 'cm' | 'm';
                          const newForm = { ...purchaseForm, dimUnit: unit };
                          
                          const w = newForm.weight || 0;
                          const actualWeight = newForm.weightUnit === 'kg' ? w : w / 1000;
                          
                          const factor = unit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          let finalWeight = Math.max(actualWeight, volWeight);
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            finalWeight = Math.max(actualWeight, newForm.volume * 200);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        className="p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                      >
                        <option value="cm">см</option>
                        <option value="m">м</option>
                      </select>
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Вага</label>
                    <div className="flex gap-2">
                      <input 
                        type="number" 
                        step="0.01"
                        value={purchaseForm.weight || ''}
                        onChange={(e) => {
                          const w = parseFloat(e.target.value) || 0;
                          const newForm = { ...purchaseForm, weight: w };
                          
                          const actualWeight = newForm.weightUnit === 'kg' ? w : w / 1000;
                          
                          const factor = newForm.dimUnit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          let finalWeight = Math.max(actualWeight, volWeight);
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            finalWeight = Math.max(actualWeight, newForm.volume * 200);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        placeholder="0.00" 
                        className="flex-1 p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                      <select 
                        value={purchaseForm.weightUnit}
                        onChange={(e) => {
                          const unit = e.target.value as 'g' | 'kg';
                          const newForm = { ...purchaseForm, weightUnit: unit };
                          
                          const w = newForm.weight;
                          const actualWeight = unit === 'kg' ? w : w / 1000;
                          
                          const factor = newForm.dimUnit === 'm' ? 100 : 1;
                          const volWeight = ((newForm.width * factor) * (newForm.height * factor) * (newForm.length * factor)) / 5000;
                          
                          let finalWeight = Math.max(actualWeight, volWeight);
                          if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                            finalWeight = Math.max(actualWeight, newForm.volume * 200);
                          }

                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;

                          setPurchaseForm(newForm);
                        }}
                        className="p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                      >
                        <option value="g">г</option>
                        <option value="kg">кг</option>
                      </select>
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Об'єм (м³)</label>
                    <input 
                      type="number" 
                      step="0.001"
                      value={purchaseForm.volume || ''}
                      onChange={(e) => {
                        const v = parseFloat(e.target.value) || 0;
                        const newForm = { ...purchaseForm, volume: v };
                        
                        // If dimensions are missing, calculate density if weight is present
                        const actualWeight = newForm.weightUnit === 'kg' ? newForm.weight : (newForm.weight || 0) / 1000;
                        if (v > 0 && actualWeight > 0) {
                          newForm.density = actualWeight / v;
                        }

                        // Recalculate shipping cost based on volume if dimensions are missing
                        if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0) {
                          const volWeight = v * 200;
                          const finalWeight = Math.max(actualWeight, volWeight);
                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;
                        }

                        setPurchaseForm(newForm);
                      }}
                      placeholder="0.000" 
                      className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Щільність (кг/м³)</label>
                    <input 
                      type="number" 
                      value={purchaseForm.density || ''}
                      onChange={(e) => {
                        const d = parseFloat(e.target.value) || 0;
                        const newForm = { ...purchaseForm, density: d };
                        
                        // If volume is missing but weight and density are present, calculate volume
                        const actualWeight = newForm.weightUnit === 'kg' ? newForm.weight : (newForm.weight || 0) / 1000;
                        if (d > 0 && actualWeight > 0 && (!newForm.volume || newForm.volume === 0)) {
                          newForm.volume = actualWeight / d;
                        }

                        // Recalculate shipping cost based on density if dimensions are missing
                        if (newForm.width === 0 && newForm.height === 0 && newForm.length === 0 && newForm.volume && newForm.volume > 0) {
                          const volWeight = newForm.volume * 200;
                          const finalWeight = Math.max(actualWeight, volWeight);
                          newForm.shippingCost = finalWeight * shippingPricePerKg;
                          if (newForm.isFabric) newForm.shippingCost += 5;
                          if (newForm.isInsured) newForm.shippingCost += (newForm.declaredValue || 0) * 0.02;
                        }

                        setPurchaseForm(newForm);
                      }}
                      placeholder="0" 
                      className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Вартість доставки (грн)</label>
                    <input 
                      type="number" 
                      step="1"
                      value={purchaseForm.shippingCost ? Math.round(purchaseForm.shippingCost * usdToUah) : ''}
                      onChange={(e) => setPurchaseForm({...purchaseForm, shippingCost: (parseFloat(e.target.value) || 0) / usdToUah})}
                      placeholder="0" 
                      className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                    />
                  </div>
                </div>
              </div>

              {/* Additional Services Section */}
              <div className="lg:col-span-3 border-t border-gray-100 pt-6 mt-2">
                <h4 className="text-xs font-black text-black uppercase tracking-widest mb-4 flex items-center gap-2">
                  <ShieldCheck className="w-4 h-4 text-black" />
                  Додаткові послуги
                </h4>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                  <label className="flex items-center gap-3 p-4 bg-gray-50 rounded-xl border border-gray-100 cursor-pointer hover:bg-gray-100 transition-all">
                    <input 
                      type="checkbox" 
                      checked={purchaseForm.isFabric}
                      onChange={(e) => {
                        const val = e.target.checked;
                        const newForm = { ...purchaseForm, isFabric: val };
                        if (val) newForm.shippingCost = (newForm.shippingCost || 0) + 5;
                        else newForm.shippingCost = (newForm.shippingCost || 0) - 5;
                        setPurchaseForm(newForm);
                      }}
                      className="w-5 h-5 accent-black"
                    />
                    <span className="text-xs font-bold text-black">Тканина (+5$/кг)</span>
                  </label>

                  <label className="flex items-center gap-3 p-4 bg-gray-50 rounded-xl border border-gray-100 cursor-pointer hover:bg-gray-100 transition-all">
                    <input 
                      type="checkbox" 
                      checked={purchaseForm.isPressed}
                      onChange={(e) => setPurchaseForm({...purchaseForm, isPressed: e.target.checked})}
                      className="w-5 h-5 accent-black"
                    />
                    <span className="text-xs font-bold text-black">Пресування</span>
                  </label>

                  <label className="flex items-center gap-3 p-4 bg-gray-50 rounded-xl border border-gray-100 cursor-pointer hover:bg-gray-100 transition-all">
                    <input 
                      type="checkbox" 
                      checked={purchaseForm.isInsured}
                      onChange={(e) => {
                        const val = e.target.checked;
                        const newForm = { ...purchaseForm, isInsured: val };
                        if (val) newForm.shippingCost = (newForm.shippingCost || 0) + (newForm.declaredValue || 0) * 0.02;
                        else newForm.shippingCost = (newForm.shippingCost || 0) - (newForm.declaredValue || 0) * 0.02;
                        setPurchaseForm(newForm);
                      }}
                      className="w-5 h-5 accent-black"
                    />
                    <span className="text-xs font-bold text-black">Страхування (2%)</span>
                  </label>

                  {purchaseForm.isInsured && (
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Оголошена вартість ($)</label>
                      <input 
                        type="number" 
                        value={purchaseForm.declaredValue || ''}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value) || 0;
                          const oldVal = purchaseForm.declaredValue || 0;
                          const newForm = { ...purchaseForm, declaredValue: val };
                          // Adjust shipping cost based on difference in insurance
                          newForm.shippingCost = (newForm.shippingCost || 0) - (oldVal * 0.02) + (val * 0.02);
                          setPurchaseForm(newForm);
                        }}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                    </div>
                  )}
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Статус</label>
                <select 
                  value={purchaseForm.status}
                  onChange={(e) => setPurchaseForm({...purchaseForm, status: e.target.value as any})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                >
                  {Object.entries(statusLabels).map(([val, label]) => (
                    <option key={val} value={val}>{label}</option>
                  ))}
                </select>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Трек-номер</label>
                <input 
                  type="text" 
                  value={purchaseForm.trackNumber}
                  onChange={(e) => handleTrackNumberChange(e.target.value)}
                  placeholder="TB123456789" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black transition-all"
                />
              </div>

              <div className="space-y-2 lg:col-span-3">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Фото товару</label>
                <div className="flex flex-col sm:flex-row items-center sm:items-start gap-4 sm:gap-6">
                  <div className="w-24 h-24 sm:w-32 sm:h-32 bg-gray-50 rounded-2xl border-2 border-dashed border-gray-200 flex items-center justify-center overflow-hidden shrink-0">
                    {purchaseForm.photo ? (
                      <img src={purchaseForm.photo} alt="Preview" className="w-full h-full object-cover" />
                    ) : (
                      <Package className="w-8 h-8 text-gray-300" />
                    )}
                  </div>
                  <div className="flex-1 space-y-3">
                    <div className="flex gap-2">
                      <label className="flex-1 bg-white border border-gray-200 text-black py-3 px-4 rounded-xl font-bold text-xs flex items-center justify-center gap-2 hover:bg-gray-50 transition-all cursor-pointer">
                        <Upload className="w-4 h-4" />
                        Завантажити фото
                        <input 
                          type="file" 
                          accept="image/*" 
                          className="hidden" 
                          onChange={(e) => {
                            const file = e.target.files?.[0];
                            if (file) {
                              const reader = new FileReader();
                              reader.onloadend = () => {
                                setPurchaseForm({ ...purchaseForm, photo: reader.result as string });
                              };
                              reader.readAsDataURL(file);
                            }
                          }}
                        />
                      </label>
                      {purchaseForm.photo && (
                        <button 
                          onClick={() => setPurchaseForm({ ...purchaseForm, photo: '' })}
                          className="p-3 bg-red-50 text-red-500 rounded-xl hover:bg-red-100 transition-all"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                    <div className="relative">
                      <input 
                        type="text" 
                        value={purchaseForm.photo}
                        onChange={(e) => setPurchaseForm({ ...purchaseForm, photo: e.target.value })}
                        placeholder="Або вставте посилання на фото..." 
                        className="w-full p-3 bg-gray-50 rounded-xl border border-gray-100 font-bold text-xs focus:outline-none focus:ring-2 focus:ring-black" 
                      />
                      <Globe className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-300" />
                    </div>
                    <p className="text-[10px] text-gray-400 font-medium italic">Порада: Ви можете просто вставити фото з буфера обміну (Ctrl+V)</p>
                  </div>
                </div>
              </div>

              <div className="space-y-2 lg:col-span-3">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Коментар</label>
                <textarea 
                  value={purchaseForm.comment}
                  onChange={(e) => setPurchaseForm({...purchaseForm, comment: e.target.value})}
                  rows={3}
                  placeholder="Додаткова інформація..." 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
            </div>

            <div className="flex flex-col sm:flex-row gap-4 mt-10">
              <button 
                onClick={() => handleSavePurchase(false)}
                className="flex-1 py-5 bg-black text-white rounded-xl font-black text-lg hover:bg-gray-900 transition-all shadow-lg shadow-gray-100 flex items-center justify-center gap-3"
              >
                💾 Зберегти
              </button>
              <button 
                onClick={() => handleSavePurchase(true)}
                className="flex-1 py-5 bg-black text-white rounded-xl font-black text-lg hover:bg-gray-900 transition-all shadow-lg shadow-gray-100 flex items-center justify-center gap-3"
              >
                💾 Зберегти і додати ще
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {showAssignTrackModal && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowAssignTrackModal(false)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => setShowAssignTrackModal(false)} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <Layers className="w-6 h-6 md:w-8 md:h-8 text-black" />
              Призначити накладну
            </h3>
            <div className="space-y-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Трек-номер (Накладна)</label>
                <div className="relative">
                  <input 
                    type="text" 
                    list="existing-tracks-modal"
                    value={assignTrackNumber}
                    onChange={(e) => setAssignTrackNumber(e.target.value)}
                    placeholder="Введіть або вставте трек-номер..."
                    className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                  />
                  <datalist id="existing-tracks-modal">
                    {existingTrackNumbers.map(track => (
                      <option key={track} value={track} />
                    ))}
                  </datalist>
                  {existingTrackNumbers.length > 0 && (
                    <div className="mt-2 flex flex-wrap gap-2">
                      {existingTrackNumbers.slice(0, 5).map(track => (
                        <button
                          key={track}
                          onClick={() => setAssignTrackNumber(track)}
                          className="text-[10px] font-black bg-gray-100 text-gray-600 px-3 py-1.5 rounded-lg hover:bg-gray-200 transition-colors"
                        >
                          {track}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              
              <div className="bg-gray-50 p-6 rounded-2xl border border-gray-100">
                <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-4">Вибрані товари ({selectedPurchaseIds.length})</p>
                <div className="space-y-2 max-h-[200px] overflow-y-auto pr-2">
                  {purchases.filter(p => selectedPurchaseIds.includes(p.id)).map(p => (
                    <div key={p.id} className="flex items-center gap-3 bg-white p-3 rounded-xl border border-gray-100">
                      <div className="w-8 h-8 rounded-lg bg-gray-50 overflow-hidden flex-shrink-0">
                        {p.photo ? <img src={p.photo} alt={p.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" /> : <Package className="w-4 h-4 m-2 text-gray-300" />}
                      </div>
                      <p className="text-xs font-bold text-black truncate flex-1">{p.name}</p>
                      <p className="text-[10px] font-black text-gray-400">{p.quantity} шт</p>
                    </div>
                  ))}
                </div>
              </div>
            </div>
            <button 
              onClick={() => handleAssignTrack(selectedPurchaseIds, assignTrackNumber)}
              className="w-full mt-10 py-5 bg-black text-white rounded-xl font-black text-xl hover:bg-gray-900 transition-all shadow-lg shadow-gray-100"
            >
              Призначити та перенести на склад
            </button>
          </motion.div>
        </div>
      )}

      {showImportTracksModal && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowImportTracksModal(false)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => setShowImportTracksModal(false)} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <Upload className="w-6 h-6 md:w-8 md:h-8 text-black" />
              Імпорт трек-номерів
            </h3>
            <div className="space-y-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Вставте текст з трек-номерами</label>
                <textarea 
                  rows={8}
                  value={bulkImportText}
                  onChange={(e) => setBulkImportText(e.target.value)}
                  placeholder="Вставте скопійований текст тут..."
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-mono text-sm focus:outline-none focus:ring-2 focus:ring-black"
                />
              </div>
              <div className="p-4 bg-blue-50 rounded-xl border border-blue-100 flex gap-4">
                <Info className="w-6 h-6 text-blue-500 shrink-0" />
                <p className="text-xs text-blue-700 font-medium leading-relaxed">
                  Система автоматично розпізнає трек-номери, дати, вагу, об'єм та щільність з тексту та додасть їх до бази даних зі статусом "Куплено".
                </p>
              </div>
            </div>
            <button 
              onClick={handleBulkImport}
              className="w-full mt-10 py-5 bg-black text-white rounded-xl font-black text-xl hover:bg-gray-900 transition-all shadow-lg shadow-gray-100"
            >
              Імпортувати треки
            </button>
          </motion.div>
        </div>
      )}

      {showCreateBatchModal && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowCreateBatchModal(false)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-2xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => {
              setShowCreateBatchModal(false);
              setEditingBatchId(null);
              setBatchForm({
                name: '',
                shipmentDate: new Date().toISOString().split('T')[0],
                warehouse: 'Guangzhou',
                deliveryType: 'sea'
              });
            }} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <Truck className="w-6 h-6 md:w-8 md:h-8 text-black" />
              {editingBatchId ? 'Редагувати партію' : 'Створити партію доставки'}
            </h3>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва партії</label>
                <input 
                  type="text" 
                  value={batchForm.name}
                  onChange={(e) => setBatchForm({...batchForm, name: e.target.value})}
                  placeholder="Назва партії" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Дата відправки</label>
                <input 
                  type="date" 
                  value={batchForm.shipmentDate}
                  onChange={(e) => setBatchForm({...batchForm, shipmentDate: e.target.value})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Склад</label>
                <select 
                  value={batchForm.warehouse}
                  onChange={(e) => setBatchForm({...batchForm, warehouse: e.target.value})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                >
                  <option value="Guangzhou">Guangzhou</option>
                  <option value="Shenzhen">Shenzhen</option>
                  <option value="Yiwu">Yiwu</option>
                </select>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Тип доставки</label>
                <select 
                  value={batchForm.deliveryType}
                  onChange={(e) => setBatchForm({...batchForm, deliveryType: e.target.value as 'sea' | 'air'})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                >
                  <option value="air">✈️ Авіа</option>
                  <option value="sea">🚢 Море</option>
                </select>
              </div>
            </div>

            <div className="bg-gray-50 rounded-2xl p-6 border border-gray-100 mb-8">
              <h4 className="text-xs font-black text-black uppercase tracking-widest mb-4 flex items-center justify-between">
                <span>Товари в дорозі (без партії)</span>
                <span className="bg-black text-white px-2 py-0.5 rounded text-[10px]">
                  {purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId).length}
                </span>
              </h4>
              <div className="space-y-2 max-h-[300px] overflow-y-auto pr-2">
                {purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId).map(p => (
                  <div key={p.id} className="flex items-center gap-4 bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
                    <div className="w-10 h-10 bg-gray-50 rounded-lg flex items-center justify-center text-[10px] font-mono font-bold text-blue-500">
                      {p.trackNumber.slice(-4)}
                    </div>
                    <div className="flex-1">
                      <p className="text-sm font-bold text-black">{p.name}</p>
                      <p className="text-[10px] text-gray-400 font-black uppercase tracking-widest">{p.trackNumber}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-black text-black">{p.weight || 0} {p.weightUnit || 'кг'}</p>
                      <p className="text-[10px] text-gray-400 uppercase font-black tracking-widest">вага</p>
                    </div>
                  </div>
                ))}
                {purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId).length === 0 && (
                  <div className="text-center py-8">
                    <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">Немає товарів зі статусом "Відправлено з Китаю" без партії</p>
                  </div>
                )}
              </div>
            </div>

            <button 
              onClick={handleCreateBatch}
              disabled={!editingBatchId && purchases.filter(p => p.status === 'shipped_to_ua' && !p.batchId).length === 0}
              className="w-full py-5 bg-black text-white rounded-xl font-black text-xl hover:bg-gray-900 transition-all shadow-lg shadow-gray-100 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {editingBatchId ? 'Зберегти зміни' : 'Додати до партії'}
            </button>
          </motion.div>
        </div>
      )}

      {showCreateLocalDeliveryModal && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowCreateLocalDeliveryModal(false)} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-2xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => {
              setShowCreateLocalDeliveryModal(false);
              setEditingLocalDeliveryId(null);
              setLocalDeliveryForm({
                name: '',
                type: 'novaposhta',
                trackingNumber: '',
                cost: 0
              });
            }} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <Truck className="w-6 h-6 md:w-8 md:h-8 text-black" />
              {editingLocalDeliveryId ? 'Редагувати відправку' : 'Створити відправку'}
            </h3>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва відправки</label>
                <input 
                  type="text" 
                  value={localDeliveryForm.name}
                  onChange={(e) => setLocalDeliveryForm({...localDeliveryForm, name: e.target.value})}
                  placeholder="Назва" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Тип доставки</label>
                <select 
                  value={localDeliveryForm.type}
                  onChange={(e) => setLocalDeliveryForm({...localDeliveryForm, type: e.target.value as 'novaposhta' | 'pickup'})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black"
                >
                  <option value="novaposhta">📦 Нова Пошта</option>
                  <option value="pickup">🏪 Самовивіз</option>
                </select>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">ТТН (необов'язково)</label>
                <input 
                  type="text" 
                  value={localDeliveryForm.trackingNumber || ''}
                  onChange={(e) => setLocalDeliveryForm({...localDeliveryForm, trackingNumber: e.target.value})}
                  placeholder="Номер накладної" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Вартість доставки / затрати (грн)</label>
                <input 
                  type="number" 
                  value={localDeliveryForm.cost || ''}
                  onChange={(e) => setLocalDeliveryForm({...localDeliveryForm, cost: parseFloat(e.target.value) || 0})}
                  placeholder="0" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-black" 
                />
              </div>
            </div>

            <div className="bg-gray-50 rounded-2xl p-6 border border-gray-100 mb-8">
              <h4 className="text-xs font-black text-black uppercase tracking-widest mb-4 flex items-center justify-between">
                <span>Товари на складі (без відправки)</span>
                <span className="bg-black text-white px-2 py-0.5 rounded text-[10px]">
                  {purchases.filter(p => p.status === 'arrived_ua' && !p.localDeliveryId).length}
                </span>
              </h4>
              <div className="space-y-2 max-h-[300px] overflow-y-auto pr-2">
                {purchases.filter(p => p.status === 'arrived_ua' && !p.localDeliveryId).map(p => (
                  <div key={p.id} className="flex items-center gap-4 bg-white p-4 rounded-xl border border-gray-100 shadow-sm">
                    <div className="w-10 h-10 bg-gray-50 rounded-lg flex items-center justify-center text-[10px] font-mono font-bold text-blue-500">
                      {p.trackNumber.slice(-4)}
                    </div>
                    <div className="flex-1">
                      <p className="text-sm font-bold text-black">{p.name}</p>
                      <p className="text-[10px] text-gray-400 font-black uppercase tracking-widest">{p.trackNumber}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-black text-black">{p.weight || 0} {p.weightUnit || 'кг'}</p>
                      <p className="text-[10px] text-gray-400 uppercase font-black tracking-widest">вага</p>
                    </div>
                  </div>
                ))}
                {purchases.filter(p => p.status === 'arrived_ua' && !p.localDeliveryId).length === 0 && (
                  <div className="text-center py-8">
                    <p className="text-xs text-gray-400 font-bold uppercase tracking-widest">Немає товарів зі статусом "На складі Україна Київ" без відправки</p>
                  </div>
                )}
              </div>
            </div>

            <button 
              onClick={handleCreateLocalDelivery}
              disabled={!editingLocalDeliveryId && purchases.filter(p => p.status === 'arrived_ua' && !p.localDeliveryId).length === 0}
              className="w-full py-5 bg-black text-white rounded-xl font-black text-xl hover:bg-gray-900 transition-all shadow-lg shadow-gray-100 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {editingLocalDeliveryId ? 'Зберегти зміни' : 'Створити відправку'}
            </button>
          </motion.div>
        </div>
      )}

      {trackingModal.show && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setTrackingModal({ show: false, data: null, loading: false, error: '' })} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-2xl w-full relative z-10 shadow-2xl border-b-8 border-black max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => setTrackingModal({ show: false, data: null, loading: false, error: '' })} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <Truck className="w-6 h-6 md:w-8 md:h-8 text-black" />
              Статус відправки
            </h3>
            
            {trackingModal.loading ? (
              <div className="text-center py-10">
                <p className="text-gray-400 font-bold">Завантаження...</p>
              </div>
            ) : trackingModal.error ? (
              <div className="text-center py-10">
                <p className="text-red-500 font-bold">{trackingModal.error}</p>
              </div>
            ) : trackingModal.data ? (
              <div className="space-y-6">
                <div className="bg-gray-50 p-6 rounded-2xl">
                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Поточний статус</p>
                  <p className="text-lg font-black text-black">{trackingModal.data.Status}</p>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-gray-50 p-4 rounded-2xl">
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Дата відправки</p>
                    <p className="text-sm font-bold text-black">{trackingModal.data.DateCreated}</p>
                  </div>
                  <div className="bg-gray-50 p-4 rounded-2xl">
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Очікувана дата</p>
                    <p className="text-sm font-bold text-black">{trackingModal.data.ScheduledDeliveryDate}</p>
                  </div>
                </div>
              </div>
            ) : null}
          </motion.div>
        </div>
      )}

      {showCostModal.show && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-4 sm:px-6">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowCostModal({ show: false, batchId: null })} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-6 md:p-10 max-w-md w-full relative z-10 shadow-2xl border-b-8 border-amber-400 max-h-[90vh] overflow-y-auto"
          >
            <button onClick={() => setShowCostModal({ show: false, batchId: null })} className="absolute top-4 right-4 md:top-6 md:right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-5 h-5 md:w-6 md:h-6" />
            </button>
            <h3 className="text-xl md:text-3xl font-black text-black uppercase tracking-tight mb-6 md:mb-8 flex items-center gap-2 md:gap-3">
              <DollarSign className="w-6 h-6 md:w-8 md:h-8 text-amber-400" />
              Вартість доставки
            </h3>
            
            <div className="space-y-6 max-h-[60vh] overflow-y-auto pr-2">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Тариф</label>
                <select 
                  value={costForm.tariffId}
                  onChange={(e) => setCostForm({...costForm, tariffId: e.target.value})}
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-amber-400"
                >
                  {tariffs.map(t => (
                    <option key={t.id} value={t.id}>{t.name}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Вага партії (кг)</label>
                  <input 
                    type="number" 
                    value={costForm.totalWeight || ''}
                    onChange={(e) => setCostForm({...costForm, totalWeight: parseFloat(e.target.value) || 0})}
                    placeholder="0.00" 
                    className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-amber-400" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Об'єм партії (м³)</label>
                  <input 
                    type="number" 
                    value={costForm.volume || ''}
                    onChange={(e) => setCostForm({...costForm, volume: parseFloat(e.target.value) || 0})}
                    placeholder="0.000" 
                    step="0.001"
                    className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-amber-400" 
                  />
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Оголошена вартість ($)</label>
                <input 
                  type="number" 
                  value={costForm.declaredValue || ''}
                  onChange={(e) => setCostForm({...costForm, declaredValue: parseFloat(e.target.value) || 0})}
                  placeholder="0.00" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-amber-400" 
                />
              </div>

              <div className="grid grid-cols-1 gap-3">
                <button 
                  onClick={() => setCostForm(p => ({ ...p, isInsured: !p.isInsured }))}
                  className={cn(
                    "flex items-center justify-between p-4 rounded-xl border-2 transition-all",
                    costForm.isInsured ? "border-amber-400 bg-amber-50" : "border-gray-50 hover:border-gray-100"
                  )}
                >
                  <span className="text-[10px] font-black uppercase tracking-widest">Страхування (2%)</span>
                  <div className={cn("w-4 h-4 rounded-full border-2 flex items-center justify-center", costForm.isInsured ? "border-amber-400 bg-amber-400" : "border-gray-200")}>
                    {costForm.isInsured && <div className="w-1 h-1 bg-white rounded-full" />}
                  </div>
                </button>
                <button 
                  onClick={() => setCostForm(p => ({ ...p, isFabric: !p.isFabric }))}
                  className={cn(
                    "flex items-center justify-between p-4 rounded-xl border-2 transition-all",
                    costForm.isFabric ? "border-amber-400 bg-amber-50" : "border-gray-50 hover:border-gray-100"
                  )}
                >
                  <span className="text-[10px] font-black uppercase tracking-widest">Тканина (+0.2$/кг)</span>
                  <div className={cn("w-4 h-4 rounded-full border-2 flex items-center justify-center", costForm.isFabric ? "border-amber-400 bg-amber-400" : "border-gray-200")}>
                    {costForm.isFabric && <div className="w-1 h-1 bg-white rounded-full" />}
                  </div>
                </button>
                <button 
                  onClick={() => setCostForm(p => ({ ...p, isPressed: !p.isPressed }))}
                  className={cn(
                    "flex items-center justify-between p-4 rounded-xl border-2 transition-all",
                    costForm.isPressed ? "border-amber-400 bg-amber-50" : "border-gray-50 hover:border-gray-100"
                  )}
                >
                  <span className="text-[10px] font-black uppercase tracking-widest">Пресування (+5$)</span>
                  <div className={cn("w-4 h-4 rounded-full border-2 flex items-center justify-center", costForm.isPressed ? "border-amber-400 bg-amber-400" : "border-gray-200")}>
                    {costForm.isPressed && <div className="w-1 h-1 bg-white rounded-full" />}
                  </div>
                </button>
              </div>

              <div className="pt-4 border-t border-gray-100">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Або введіть вартість вручну ($)</label>
                <input 
                  type="number" 
                  value={costForm.deliveryCost || ''}
                  onChange={(e) => setCostForm({...costForm, deliveryCost: parseFloat(e.target.value) || 0})}
                  placeholder="0.00" 
                  className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-amber-400" 
                />
              </div>

              <div className="p-6 bg-amber-50 rounded-2xl border border-amber-100">
                <div className="flex justify-between items-center mb-2">
                  <span className="text-[10px] font-black text-amber-800 uppercase tracking-widest">Розрахункова вартість</span>
                  <span className="text-2xl font-black text-amber-600">
                    {(() => {
                      const selectedTariff = tariffs.find(t => t.id === costForm.tariffId) || tariffs[0];
                      let shippingCost = 0;
                      if (selectedTariff.pricePerKg) {
                        const volumetricWeight = (costForm.volume * 1000000) / (selectedTariff.volumetricFactor || 5000);
                        shippingCost = selectedTariff.pricePerKg * Math.max(costForm.totalWeight, volumetricWeight);
                      } else if (selectedTariff.densityTiers) {
                        const density = costForm.volume > 0 ? costForm.totalWeight / costForm.volume : 0;
                        const tier = selectedTariff.densityTiers.find(t => density >= t.min && (t.max === null || density < t.max));
                        if (tier) {
                          shippingCost = tier.unit === 'm3' ? tier.price * costForm.volume : tier.price * costForm.totalWeight;
                        }
                      }
                      const insurance = (costForm.isInsured && costForm.declaredValue) ? Math.max(1, costForm.declaredValue * 0.02) : 0;
                      const fabricSurcharge = costForm.isFabric ? costForm.totalWeight * 0.2 : 0;
                      const pressingCost = costForm.isPressed ? 5 : 0;
                      const localDelivery = (selectedTariff.localDeliveryPrice || 0) * costForm.totalWeight;
                      const total = shippingCost + insurance + localDelivery + fabricSurcharge + pressingCost;
                      return `$${total.toFixed(2)}`;
                    })()}
                  </span>
                </div>
                <p className="text-[10px] text-amber-700 font-medium leading-relaxed">
                  Вартість буде розподілена між товарами пропорційно їх вазі.
                </p>
              </div>
            </div>

            <button 
              onClick={handleSaveCosts}
              className="w-full mt-10 py-5 bg-amber-400 text-white rounded-xl font-black text-xl hover:bg-amber-500 transition-all shadow-lg shadow-amber-100"
            >
              Зберегти вартість
            </button>
          </motion.div>
        </div>
      )}
      {showSaleModal.show && (
        <div className="fixed inset-0 z-[300] flex items-center justify-center px-6">
          <div className="absolute inset-0 bg-black/95 backdrop-blur-md" onClick={() => setShowSaleModal({ show: false, purchaseId: null })} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-[40px] p-10 max-w-2xl w-full relative z-10 shadow-2xl border-b-8 border-emerald-400 overflow-hidden"
          >
            <button onClick={() => setShowSaleModal({ show: false, purchaseId: null })} className="absolute top-8 right-8 text-gray-400 hover:text-black transition-colors z-20">
              <X className="w-8 h-8" />
            </button>

            <div className="flex flex-col md:flex-row gap-10">
              <div className="w-full md:w-1/2 space-y-6">
                <div className="aspect-square bg-gray-50 rounded-[32px] overflow-hidden border border-gray-100 shadow-inner group relative">
                  {(() => {
                    const purchase = purchases.find(p => p.id === showSaleModal.purchaseId);
                    return purchase?.photo ? (
                      <img src={purchase.photo} alt={purchase.name} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-700" referrerPolicy="no-referrer" />
                    ) : (
                      <div className="w-full h-full flex flex-col items-center justify-center text-gray-200">
                        <Package className="w-20 h-20 mb-4" />
                        <span className="text-[10px] font-black uppercase tracking-widest">Фото відсутнє</span>
                      </div>
                    );
                  })()}
                  <div className="absolute bottom-4 left-4 right-4 bg-white/90 backdrop-blur-sm p-4 rounded-2xl shadow-lg border border-white/50">
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Товар</p>
                    <p className="text-sm font-black text-black line-clamp-1">
                      {purchases.find(p => p.id === showSaleModal.purchaseId)?.name}
                    </p>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Вага</p>
                    <p className="text-lg font-black text-black">
                      {purchases.find(p => p.id === showSaleModal.purchaseId)?.weight || 0} {purchases.find(p => p.id === showSaleModal.purchaseId)?.weightUnit || 'кг'}
                    </p>
                  </div>
                  <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">Кількість</p>
                    <p className="text-lg font-black text-black">
                      {purchases.find(p => p.id === showSaleModal.purchaseId)?.quantity || 0} шт
                    </p>
                  </div>
                </div>
              </div>

              <div className="w-full md:w-1/2 space-y-6">
                <h3 className="text-3xl font-black text-black uppercase tracking-tight flex items-center gap-3">
                  <Store className="w-8 h-8 text-emerald-400" />
                  Видача товару
                </h3>

                <div className="space-y-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Куди видано</label>
                    <select
                      value={saleForm.saleDestination || 'physical_store'}
                      onChange={(e) => setSaleForm({...saleForm, saleDestination: e.target.value as any})}
                      className="w-full p-4 bg-gray-50 rounded-2xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-emerald-400 transition-all"
                    >
                      <option value="physical_store">Фізичний магазин</option>
                      <option value="online_store">Інтернет-магазин</option>
                      <option value="personal_use">Власні потреби</option>
                    </select>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна продажу (грн)</label>
                    <input 
                      type="number" 
                      value={saleForm.sellingPrice || ''}
                      onChange={(e) => setSaleForm({...saleForm, sellingPrice: parseFloat(e.target.value) || 0})}
                      placeholder="0.00" 
                      className="w-full p-4 bg-gray-50 rounded-2xl border border-gray-100 font-black text-xl text-black focus:outline-none focus:ring-2 focus:ring-emerald-400 transition-all" 
                      disabled={saleForm.saleDestination === 'personal_use'}
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Доставка Україна (грн)</label>
                    <input 
                      type="number" 
                      value={saleForm.ukraineDeliveryCost || ''}
                      onChange={(e) => setSaleForm({...saleForm, ukraineDeliveryCost: parseFloat(e.target.value) || 0})}
                      placeholder="0.00" 
                      className="w-full p-4 bg-gray-50 rounded-2xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-emerald-400 transition-all" 
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Доставка НП (грн)</label>
                    <input 
                      type="number" 
                      value={saleForm.novaPoshtaCost || ''}
                      onChange={(e) => setSaleForm({...saleForm, novaPoshtaCost: parseFloat(e.target.value) || 0})}
                      placeholder="0.00" 
                      className="w-full p-4 bg-gray-50 rounded-2xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-emerald-400 transition-all" 
                    />
                  </div>

                  <div className="flex items-center gap-3 p-4 bg-gray-50 rounded-2xl border border-gray-100 cursor-pointer hover:bg-gray-100 transition-all" onClick={() => setSaleForm({...saleForm, markup: !saleForm.markup})}>
                    <div className={cn(
                      "w-6 h-6 rounded-lg border-2 flex items-center justify-center transition-all",
                      saleForm.markup ? "bg-emerald-400 border-emerald-400" : "bg-white border-gray-200"
                    )}>
                      {saleForm.markup && <CheckCircle2 className="w-4 h-4 text-white" />}
                    </div>
                    <div>
                      <p className="text-xs font-black text-black uppercase tracking-widest">Націнка</p>
                      <p className="text-[10px] text-gray-400 font-medium">Не враховувати в основний розрахунок</p>
                    </div>
                  </div>

                  {(() => {
                    const purchase = purchases.find(p => p.id === showSaleModal.purchaseId);
                    if (!purchase) return null;
                    const costPriceUah = (purchase.priceYuan * purchase.quantity) * purchase.exchangeRate;
                    const chinaDeliveryUah = (purchase.deliveryCostPerItem || 0) * usdToUah;
                    const intDeliveryUah = (purchase.shippingCost || 0) * usdToUah;
                    const totalCost = costPriceUah + chinaDeliveryUah + intDeliveryUah + saleForm.ukraineDeliveryCost + (saleForm.novaPoshtaCost || 0) + getLocalDeliveryCost(purchase);
                    const profit = saleForm.sellingPrice - totalCost;
                    
                    return (
                      <div className={cn(
                        "p-6 rounded-3xl border-2 transition-all",
                        profit >= 0 ? "bg-emerald-50 border-emerald-100" : "bg-red-50 border-red-100"
                      )}>
                        <div className="flex justify-between items-center">
                          <div>
                            <p className="text-[10px] font-black uppercase tracking-widest text-gray-400 mb-1">Очікуваний прибуток</p>
                            <p className={cn("text-2xl font-black", profit >= 0 ? "text-emerald-600" : "text-red-600")}>
                              {profit > 0 ? '+' : ''}{profit.toFixed(0)} грн
                            </p>
                          </div>
                          <div className={cn(
                            "w-12 h-12 rounded-2xl flex items-center justify-center",
                            profit >= 0 ? "bg-emerald-100 text-emerald-600" : "bg-red-100 text-red-600"
                          )}>
                            <BarChart3 className="w-6 h-6" />
                          </div>
                        </div>
                      </div>
                    );
                  })()}
                </div>

                <button 
                  onClick={handleSale}
                  className="w-full py-5 bg-black text-white rounded-[24px] font-black text-xl hover:bg-gray-900 transition-all shadow-xl shadow-gray-100 flex items-center justify-center gap-3"
                >
                  <CheckCircle2 className="w-6 h-6" />
                  Підтвердити
                </button>
              </div>
            </div>
          </motion.div>
        </div>
      )}

      {showTariffModal.show && (
        <div className="fixed inset-0 z-[400] flex items-center justify-center px-4">
          <div className="absolute inset-0 bg-black/90 backdrop-blur-sm" onClick={() => setShowTariffModal({ show: false, tariffId: null })} />
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-3xl p-8 max-w-2xl w-full relative z-10 shadow-2xl border-b-8 border-blue-400 flex flex-col max-h-[90vh]"
          >
            <button onClick={() => setShowTariffModal({ show: false, tariffId: null })} className="absolute top-6 right-6 text-gray-400 hover:text-black transition-colors">
              <X className="w-6 h-6" />
            </button>
            <h3 className="text-2xl font-black text-black uppercase tracking-tight mb-6">
              {showTariffModal.tariffId ? 'Редагувати тариф' : 'Додати тариф'}
            </h3>
            
            <div className="overflow-y-auto pr-2 space-y-6 flex-1">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Назва тарифу</label>
                    <input 
                      type="text" 
                      value={tariffForm.name}
                      onChange={(e) => setTariffForm({...tariffForm, name: e.target.value})}
                      placeholder="Наприклад: Авіа Експрес" 
                      className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Іконка</label>
                      <select 
                        value={tariffForm.iconName}
                        onChange={(e) => setTariffForm({...tariffForm, iconName: e.target.value as any})}
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400"
                      >
                        <option value="Plane">Літак</option>
                        <option value="Ship">Корабель</option>
                        <option value="Truck">Вантажівка</option>
                        <option value="Train">Потяг</option>
                        <option value="Zap">Експрес</option>
                      </select>
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Термін доставки</label>
                      <input 
                        type="text" 
                        value={tariffForm.deliveryDays}
                        onChange={(e) => setTariffForm({...tariffForm, deliveryDays: e.target.value})}
                        placeholder="7-10 днів" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Опис</label>
                    <textarea 
                      value={tariffForm.description}
                      onChange={(e) => setTariffForm({...tariffForm, description: e.target.value})}
                      placeholder="Короткий опис тарифу" 
                      rows={3}
                      className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400 resize-none" 
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Ціна за кг ($)</label>
                      <input 
                        type="number" 
                        value={tariffForm.pricePerKg || ''}
                        onChange={(e) => setTariffForm({...tariffForm, pricePerKg: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Коеф. об'єму</label>
                      <input 
                        type="number" 
                        value={tariffForm.volumetricFactor || ''}
                        onChange={(e) => setTariffForm({...tariffForm, volumetricFactor: parseFloat(e.target.value) || 0})}
                        placeholder="5000" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>
                </div>

                <div className="space-y-6">
                  <div className="grid grid-cols-3 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Мін. вага (кг)</label>
                      <input 
                        type="number" 
                        value={tariffForm.minWeight || ''}
                        onChange={(e) => setTariffForm({...tariffForm, minWeight: parseFloat(e.target.value) || 0})}
                        placeholder="0" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Мін. об'єм (м³)</label>
                      <input 
                        type="number" 
                        value={tariffForm.minVolume || ''}
                        onChange={(e) => setTariffForm({...tariffForm, minVolume: parseFloat(e.target.value) || 0})}
                        placeholder="0" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Мін. вартість ($)</label>
                      <input 
                        type="number" 
                        value={tariffForm.minCost || ''}
                        onChange={(e) => setTariffForm({...tariffForm, minCost: parseFloat(e.target.value) || 0})}
                        placeholder="0" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-3 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Страховка (%)</label>
                      <input 
                        type="number" 
                        value={tariffForm.insuranceRate || ''}
                        onChange={(e) => setTariffForm({...tariffForm, insuranceRate: parseFloat(e.target.value) || 0})}
                        placeholder="2" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Пак. ($/кг)</label>
                      <input 
                        type="number" 
                        value={tariffForm.packagingCost || ''}
                        onChange={(e) => setTariffForm({...tariffForm, packagingCost: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Пак. ($/м³)</label>
                      <input 
                        type="number" 
                        value={tariffForm.packagingCostPerM3 || ''}
                        onChange={(e) => setTariffForm({...tariffForm, packagingCostPerM3: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Митниця ($)</label>
                      <input 
                        type="number" 
                        value={tariffForm.customsFee || ''}
                        onChange={(e) => setTariffForm({...tariffForm, customsFee: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Обробка ($)</label>
                      <input 
                        type="number" 
                        value={tariffForm.handlingFee || ''}
                        onChange={(e) => setTariffForm({...tariffForm, handlingFee: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Паливна надб. (%)</label>
                      <input 
                        type="number" 
                        value={tariffForm.fuelSurcharge || ''}
                        onChange={(e) => setTariffForm({...tariffForm, fuelSurcharge: parseFloat(e.target.value) || 0})}
                        placeholder="0" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                    <div className="space-y-2">
                      <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Доставка по UA ($/кг)</label>
                      <input 
                        type="number" 
                        value={tariffForm.localDeliveryPrice || ''}
                        onChange={(e) => setTariffForm({...tariffForm, localDeliveryPrice: parseFloat(e.target.value) || 0})}
                        placeholder="0.00" 
                        className="w-full p-4 bg-gray-50 rounded-xl border border-gray-100 font-bold focus:outline-none focus:ring-2 focus:ring-blue-400" 
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* Density Tiers Section */}
              <div className="space-y-4 pt-4 border-t">
                <div className="flex justify-between items-center">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Тарифікація за щільністю</label>
                  <button 
                    onClick={() => {
                      const newTiers = [...(tariffForm.densityTiers || []), { min: 0, max: 0, price: 0, unit: 'kg' as const }];
                      setTariffForm({...tariffForm, densityTiers: newTiers});
                    }}
                    className="text-[10px] font-black text-blue-500 uppercase tracking-widest hover:text-blue-600"
                  >
                    + Додати поріг
                  </button>
                </div>
                
                <div className="space-y-3">
                  {tariffForm.densityTiers?.map((tier, index) => (
                    <div key={index} className="grid grid-cols-5 gap-2 items-end bg-gray-50 p-3 rounded-xl border border-gray-100">
                      <div className="space-y-1">
                        <label className="text-[8px] font-black text-gray-400 uppercase">Мін</label>
                        <input 
                          type="number" 
                          value={tier.min}
                          onChange={(e) => {
                            const newTiers = [...(tariffForm.densityTiers || [])];
                            newTiers[index].min = parseFloat(e.target.value) || 0;
                            setTariffForm({...tariffForm, densityTiers: newTiers});
                          }}
                          className="w-full p-2 bg-white rounded-lg border border-gray-100 text-xs font-bold focus:outline-none"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[8px] font-black text-gray-400 uppercase">Макс</label>
                        <input 
                          type="number" 
                          value={tier.max || ''}
                          onChange={(e) => {
                            const newTiers = [...(tariffForm.densityTiers || [])];
                            newTiers[index].max = e.target.value ? parseFloat(e.target.value) : null;
                            setTariffForm({...tariffForm, densityTiers: newTiers});
                          }}
                          className="w-full p-2 bg-white rounded-lg border border-gray-100 text-xs font-bold focus:outline-none"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[8px] font-black text-gray-400 uppercase">Ціна</label>
                        <input 
                          type="number" 
                          value={tier.price}
                          onChange={(e) => {
                            const newTiers = [...(tariffForm.densityTiers || [])];
                            newTiers[index].price = parseFloat(e.target.value) || 0;
                            setTariffForm({...tariffForm, densityTiers: newTiers});
                          }}
                          className="w-full p-2 bg-white rounded-lg border border-gray-100 text-xs font-bold focus:outline-none"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[8px] font-black text-gray-400 uppercase">Од.</label>
                        <select 
                          value={tier.unit}
                          onChange={(e) => {
                            const newTiers = [...(tariffForm.densityTiers || [])];
                            newTiers[index].unit = e.target.value as 'kg' | 'm3';
                            setTariffForm({...tariffForm, densityTiers: newTiers});
                          }}
                          className="w-full p-2 bg-white rounded-lg border border-gray-100 text-xs font-bold focus:outline-none"
                        >
                          <option value="kg">кг</option>
                          <option value="m3">м³</option>
                        </select>
                      </div>
                      <button 
                        onClick={() => {
                          const newTiers = tariffForm.densityTiers?.filter((_, i) => i !== index);
                          setTariffForm({...tariffForm, densityTiers: newTiers});
                        }}
                        className="p-2 text-red-400 hover:text-red-500 transition-colors"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  ))}
                  {(!tariffForm.densityTiers || tariffForm.densityTiers.length === 0) && (
                    <p className="text-center text-[10px] text-gray-400 py-2 italic">Пороги щільності не задані</p>
                  )}
                </div>
              </div>
            </div>

            <div className="pt-6 mt-6 border-t">
              <button 
                onClick={() => {
                  if (!tariffForm.name) {
                    addNotification('Введіть назву тарифу', 'error');
                    return;
                  }
                  if (showTariffModal.tariffId) {
                    setTariffs(prev => prev.map(t => t.id === showTariffModal.tariffId ? { ...t, ...tariffForm } as Tariff : t));
                    addNotification('Тариф оновлено', 'success');
                  } else {
                    const newTariff: Tariff = {
                      ...tariffForm as Tariff,
                      id: Math.random().toString(36).substr(2, 9)
                    };
                    setTariffs(prev => [...prev, newTariff]);
                    addNotification('Тариф додано', 'success');
                  }
                  setShowTariffModal({ show: false, tariffId: null });
                }}
                className="w-full py-5 bg-black text-white rounded-xl font-black text-lg hover:bg-gray-900 transition-all shadow-lg shadow-gray-100"
              >
                {showTariffModal.tariffId ? 'Зберегти зміни' : 'Додати тариф'}
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {confirmModal && (
        <DeleteConfirmModal 
          show={confirmModal.show}
          title={confirmModal.title}
          message={confirmModal.message}
          onConfirm={() => {
            confirmModal.onConfirm();
            setConfirmModal(null);
          }}
          onCancel={() => setConfirmModal(null)}
        />
      )}

      {showImportWaybillModal && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4 overflow-y-auto">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            className="bg-white rounded-[40px] w-full max-w-2xl overflow-hidden shadow-2xl"
          >
            <div className="p-10 border-b border-gray-50 flex justify-between items-center bg-indigo-600 text-white">
              <div>
                <h2 className="text-3xl font-black uppercase tracking-tight">Імпорт накладних</h2>
                <p className="text-indigo-100 text-sm font-bold mt-1 uppercase tracking-widest">Вставте текст з даними для автоматичного створення</p>
              </div>
              <button 
                onClick={() => setShowImportWaybillModal(false)}
                className="w-12 h-12 rounded-2xl bg-white/10 flex items-center justify-center hover:bg-white/20 transition-all"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="p-10 space-y-8">
              <div className="space-y-4">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest block ml-1">Текст для імпорту</label>
                <textarea 
                  value={waybillImportText}
                  onChange={(e) => setWaybillImportText(e.target.value)}
                  placeholder="Вставте дані тут..."
                  className="w-full h-64 bg-gray-50 border border-gray-100 rounded-3xl p-6 text-sm font-bold outline-none focus:ring-4 focus:ring-indigo-500/10 transition-all resize-none"
                />
                <div className="bg-amber-50 border border-amber-100 rounded-2xl p-4 flex gap-3">
                  <Info className="w-5 h-5 text-amber-600 flex-shrink-0" />
                  <p className="text-xs font-bold text-amber-800 leading-relaxed">
                    Система автоматично розпізнає трек-номер, дату, вагу, об'єм та щільність. 
                    Вартість доставки буде розрахована згідно з обраним тарифом.
                  </p>
                </div>
              </div>

              <button 
                onClick={handleImportWaybills}
                disabled={!waybillImportText.trim()}
                className="w-full py-6 bg-indigo-600 text-white rounded-2xl font-black text-xl uppercase tracking-widest hover:bg-indigo-700 transition-all shadow-xl shadow-indigo-100 disabled:opacity-50 disabled:shadow-none flex items-center justify-center gap-3"
              >
                <PlusCircle className="w-6 h-6" />
                Створити накладні
              </button>
            </div>
          </motion.div>
        </div>
      )}

      {selectedPurchaseIds.length > 0 && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-black text-white px-6 py-4 rounded-2xl shadow-2xl flex items-center gap-6 z-50 animate-in slide-in-from-bottom-10 fade-in duration-300">
          <div className="flex items-center gap-2">
            <span className="font-black text-lg">{selectedPurchaseIds.length}</span>
            <span className="text-xs uppercase tracking-widest text-gray-400">обрано</span>
          </div>
          <div className="h-8 w-px bg-white/20"></div>
          <div className="flex items-center gap-4">
            <select
              onChange={(e) => {
                if (e.target.value) {
                  handleBulkStatusChange(e.target.value as Purchase['status']);
                  e.target.value = '';
                }
              }}
              className="bg-white/10 text-white border-none rounded-xl px-4 py-2 text-xs font-bold outline-none cursor-pointer"
              defaultValue=""
            >
              <option value="" disabled>Змінити статус...</option>
              {Object.entries(statusLabels).map(([value, label]) => (
                <option key={value} value={value} className="text-black">{label}</option>
              ))}
            </select>
            <button 
              onClick={() => setSelectedPurchaseIds([])}
              className="text-xs font-bold text-gray-400 hover:text-white transition-colors"
            >
              Скасувати
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
