export interface MappingItem {
  Child: string;
  Parent: string;
}

export interface Warehouse {
  Id: number;
  Code: string;
  Name: string;
}

export interface ExportItem {
  Sku: string;
  Name: string;
  SoLuong: number;
  ThungPercent: number;
  LePercent: number;
  SlThung: number;
  SlLe: number;
  SlGoiLe: number;
  QuyCach: number;
  FinalSku: string;
  Is3N: boolean;
  HasMapping: boolean;
  ErrorMessage: string;
}

export interface XeData {
  XeName: string;
  Items: ExportItem[];
  TotalSlThung: number;
  TotalSlGoiLe: number;
}

export interface ProcessResult {
  Success: boolean;
  AllItems: ExportItem[];
  XeList: XeData[];
  Errors: string[];
  TotalRows: number;
  SuccessRows: number;
  ErrorRows: number;
  TotalSlThung: number;
  TotalSlGoiLe: number;
  SelectedWarehouse: Warehouse | null;
}
