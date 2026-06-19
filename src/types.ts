export interface Product {
  ten: string;
  searchStr?: string;
}

export interface WeightMap {
  [key: string]: number;
}

export interface InventoryItem {
  id: string; // unique item id
  ke: string;
  ten: string;
  bao: number;
  slBao: number;
  le: number;
  tongCon: number;
  khoiLuong: number;
  ghiChu: string;
}

export interface HistoryItem {
  id: string;
  ngay: string;
  loai: 'NHẬP' | 'XUẤT';
  ke: string;
  tenHang: string;
  bao: number;
  slBao: number;
  le: number;
  tong: number;
  kg: number;
  ghiChu: string;
}

export interface ContItem {
  stt: number;
  id: string; // unique row tracking ID
  grade: string;
  standard: string;
  size: string;
  finish: string;
  no: string;
  contName: string;
  bao: number;
  slBao: number;
  le: number;
  ghiChu: string;
}
