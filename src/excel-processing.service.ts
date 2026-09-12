import { Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import { MappingItem, ExportItem, XeData, ProcessResult, Warehouse } from './index';

@Injectable()
export class ExcelProcessingService {
  private readonly mappingPath = path.join(process.cwd(), 'data', 'mapping.json');
  private readonly warehousePath = path.join(process.cwd(), 'data', 'warehouses.json');

  constructor() {
    const dataDir = path.join(process.cwd(), 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }
  }

  loadMapping(): MappingItem[] {
    if (!fs.existsSync(this.mappingPath)) {
      return [];
    }
    const data = fs.readFileSync(this.mappingPath, 'utf-8');
    return JSON.parse(data);
  }

  saveMapping(mapping: MappingItem[]): void {
    fs.writeFileSync(this.mappingPath, JSON.stringify(mapping, null, 2));
  }

  loadWarehouses(): Warehouse[] {
    if (!fs.existsSync(this.warehousePath)) {
      return [];
    }
    const data = fs.readFileSync(this.warehousePath, 'utf-8');
    return JSON.parse(data);
  }

  loadMappingToDict(): Map<string, string> {
    const mapping = this.loadMapping();
    const dict = new Map<string, string>();
    mapping.forEach(item => {
      if (item.Child && item.Parent) {
        dict.set(item.Child.toUpperCase(), item.Parent.toUpperCase());
      }
    });
    return dict;
  }

  processBangTheoDoi(filePath: string, mapping: Map<string, string>): ProcessResult {
    const result: ProcessResult = {
      Success: true,
      AllItems: [],
      XeList: [],
      Errors: [],
      TotalRows: 0,
      SuccessRows: 0,
      ErrorRows: 0,
      TotalSlThung: 0,
      TotalSlGoiLe: 0,
      SelectedWarehouse: null,
    };

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames.find(name => 
      name.includes('QUY RA SỐ KHỐI') || name.includes('QUY RA')
    );

    if (!sheetName) {
      result.Success = false;
      result.Errors.push("Không tìm thấy sheet 'QUY RA SỐ KHỐI'");
      return result;
    }

    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

    // Find all xe/cont rows
    const xeRows: number[] = [];
    data.forEach((row, index) => {
      const col2 = row[1]?.toString().trim();
      if (col2 && (col2.toLowerCase().startsWith('xe') || col2.toLowerCase().startsWith('cont'))) {
        xeRows.push(index);
      }
    });

    // If no xe/cont rows found, treat all data as one group
    if (xeRows.length === 0) {
      xeRows.push(-1);
    }

    // Process each xe/cont
    for (let i = 0; i < xeRows.length; i++) {
      const xeRow = xeRows[i];
      const xeName = xeRow === -1 ? 'All' : (data[xeRow][1]?.toString().trim() || `Xe ${i + 1}`);
      const startRow = xeRow + 1;
      const endRow = (i + 1 < xeRows.length) ? xeRows[i + 1] - 1 : data.length - 1;

      const xeData: XeData = {
        XeName: xeName,
        Items: [],
        TotalSlThung: 0,
        TotalSlGoiLe: 0,
      };

      for (let r = startRow; r <= endRow; r++) {
        const row = data[r];
        if (!row) continue;

        const sku = row[2]?.toString().trim();
        const name = row[3]?.toString().trim();
        const quyCachText = row[6]?.toString().trim();
        const slText = row[7]?.toString().trim().replace(/,/g, '');
        const thungPercentText = row[8]?.toString().trim().replace(/%/g, '');
        const lePercentText = row[9]?.toString().trim().replace(/%/g, '');
        const slThungText = row[10]?.toString().trim().replace(/,/g, '');
        const slLeText = row[11]?.toString().trim().replace(/,/g, '');

        if (!sku) continue;

        const item: ExportItem = {
          Sku: sku,
          Name: name || '',
          Is3N: sku.toLowerCase().startsWith('3n'),
          QuyCach: parseFloat(quyCachText) || 0,
          SoLuong: parseFloat(slText) || 0,
          ThungPercent: parseFloat(thungPercentText) || 0,
          LePercent: parseFloat(lePercentText) || 0,
          SlThung: parseFloat(slThungText) || 0,
          SlLe: parseFloat(slLeText) || 0,
          SlGoiLe: 0,
          FinalSku: '',
          HasMapping: false,
          ErrorMessage: '',
        };

        item.SlGoiLe = item.SlLe * item.QuyCach;

        // Process mapping - Tat ca ma (1N va 3N) deu check mapping
        if (mapping.has(sku.toUpperCase())) {
          item.FinalSku = mapping.get(sku.toUpperCase())!;
          item.HasMapping = true;
        } else {
          item.FinalSku = sku;
          item.HasMapping = false;
        }

        xeData.Items.push(item);
        result.AllItems.push(item);
        result.TotalRows++;

        if (item.ErrorMessage) {
          result.Errors.push(item.ErrorMessage);
          result.ErrorRows++;
        } else {
          result.SuccessRows++;
          result.TotalSlThung += Math.round(item.SlThung);
          result.TotalSlGoiLe += Math.round(item.SlGoiLe);
          xeData.TotalSlThung += Math.round(item.SlThung);
          xeData.TotalSlGoiLe += Math.round(item.SlGoiLe);
        }
      }

      if (xeData.Items.length > 0) {
        result.XeList.push(xeData);
      }
    }

    result.Success = result.ErrorRows === 0;
    return result;
  }

  generateHaravanFile(processResult: ProcessResult, xeName?: string): Buffer {
    const items = xeName 
      ? processResult.XeList.find(x => x.XeName === xeName)?.Items || []
      : processResult.AllItems;

    const haravanData: any[] = [];

    items.filter(item => !item.ErrorMessage).forEach(item => {
      // Thung
      if (item.SlThung > 0) {
        haravanData.push({
          'Mã sản phẩm *': item.FinalSku,
          'Số lượng *': Math.round(item.SlThung),
        });
      }

      // Le (SL Goi Le)
      if (item.SlGoiLe > 0) {
        haravanData.push({
          'Mã sản phẩm *': `${item.FinalSku}_LE01`,
          'Số lượng *': Math.round(item.SlGoiLe),
        });
      }
    });

    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(haravanData);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, xeName || 'Haravan Export');

    return XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
  }
}
