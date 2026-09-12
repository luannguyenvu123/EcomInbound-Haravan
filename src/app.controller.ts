import { Controller, Get, Post, Body, Res, UploadedFile, UseInterceptors, Query } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import type { Response } from 'express';
import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';
import { ExcelProcessingService } from './excel-processing.service';
import type { ProcessResult, Warehouse, MappingItem } from './index';

@Controller()
export class AppController {
  private lastResult: ProcessResult | null = null;

  constructor(private readonly excelService: ExcelProcessingService) {}

  @Get()
  getIndex(@Res() res: Response) {
    const warehouses = this.excelService.loadWarehouses();
    res.render('index', { warehouses });
  }

  @Get('mapping')
  getMapping(@Res() res: Response) {
    const mappings = this.excelService.loadMapping();
    res.render('mapping', { mappings });
  }

  @Post('mapping/add')
  addMapping(@Body() body: { child: string; parent: string }, @Res() res: Response) {
    if (!body.child || !body.parent) {
      res.redirect('/mapping');
      return;
    }

    const mappings = this.excelService.loadMapping();
    
    if (mappings.some((m: MappingItem) => m.Child.toUpperCase() === body.child.toUpperCase())) {
      res.redirect('/mapping');
      return;
    }

    mappings.push({ Child: body.child.toUpperCase(), Parent: body.parent.toUpperCase() });
    this.excelService.saveMapping(mappings);
    res.redirect('/mapping');
  }

  @Post('mapping/delete')
  deleteMapping(@Body() body: { child: string }, @Res() res: Response) {
    const mappings = this.excelService.loadMapping();
    const filtered = mappings.filter((m: MappingItem) => m.Child.toUpperCase() !== body.child.toUpperCase());
    this.excelService.saveMapping(filtered);
    res.redirect('/mapping');
  }

  @Post('mapping/update')
  updateMapping(@Body() body: { oldChild: string; newChild: string; newParent: string }, @Res() res: Response) {
    if (!body.oldChild || !body.newChild || !body.newParent) {
      res.redirect('/mapping');
      return;
    }

    const mappings = this.excelService.loadMapping();
    const index = mappings.findIndex((m: MappingItem) => m.Child.toUpperCase() === body.oldChild.toUpperCase());
    
    if (index !== -1) {
      mappings[index].Child = body.newChild.toUpperCase();
      mappings[index].Parent = body.newParent.toUpperCase();
      this.excelService.saveMapping(mappings);
    }
    
    res.redirect('/mapping');
  }

  @Post('mapping/import')
  @UseInterceptors(FileInterceptor('mappingFile'))
  async importMapping(@UploadedFile() file: any, @Res() res: Response) {
    console.log('File received:', file);
    
    if (!file) {
      console.log('No file uploaded');
      res.redirect('/mapping');
      return;
    }

    try {
      console.log('Reading file:', file.path);
      const workbook = XLSX.readFile(file.path);
      const sheetName = workbook.SheetNames[0];
      console.log('Sheet name:', sheetName);
      
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
      console.log('Total rows:', data.length);

      const existingMappings = this.excelService.loadMapping();
      const existingDict = new Map(existingMappings.map((m: MappingItem) => [m.Child.toUpperCase(), m.Parent.toUpperCase()]));

      let lastParentCode = '';
      let addedCount = 0;
      
      data.slice(1).forEach((row: any[], index: number) => {
        let parentCode = row[1]?.toString().trim() || '';
        const childCode = row[2]?.toString().trim() || '';
        const parentFromLastCol = row[8]?.toString().trim() || '';

        if (!parentCode && parentFromLastCol) {
          parentCode = parentFromLastCol;
        }

        if (!parentCode) {
          parentCode = lastParentCode;
        }

        if (parentCode) {
          lastParentCode = parentCode;
        }

        if (parentCode && childCode) {
          existingDict.set(childCode.toUpperCase(), parentCode.toUpperCase());
          addedCount++;
        }
      });

      console.log('Added mappings:', addedCount);
      const updatedMappings = Array.from(existingDict.entries()).map(([Child, Parent]) => ({ Child, Parent }));
      this.excelService.saveMapping(updatedMappings);

      // Delete uploaded file
      if (fs.existsSync(file.path)) {
        fs.unlinkSync(file.path);
      }
      
      res.redirect(`/mapping?import=success&count=${addedCount}`);
    } catch (error) {
      console.error('Import error:', error);
      // Delete uploaded file on error
      if (file.path && fs.existsSync(file.path)) {
        fs.unlinkSync(file.path);
      }
      res.redirect('/mapping?import=error');
    }
  }

  @Get('mapping/template')
  downloadTemplate(@Res() res: Response) {
    // Tao file mau giong dinh dang file TH vat tu
    const ws = XLSX.utils.aoa_to_sheet([
      ['STT', 'Mã cha', 'Mã con', 'Tên', 'ĐVT', 'Mã kho', 'Quy cách', 'Số lượng'],
      [1, '1NBDGVFDC31N', '3NBDGHDDC001N', 'Bánh đa cua 60grx30 NĐ (MCPP)', 'Thùng', 'HD', 30, 100],
      [2, '', '3NBDGHPDC01N', 'Bánh đa cua 60grx30 NĐ (HD)', 'Thùng', 'HD', 30, 50],
      [3, '', '3NBDGHPDC004N', 'Bánh đa cua 60grx30 NĐ (ECOM)', 'Thùng', 'HD', 30, 30],
      [4, '1NMIGVFTTGC02N', '3NMIGHPTTGC01N', 'Mì thịt thật gà cay 85grx24 (8lốcx3) NÐ (HP)', 'Thùng', 'HD', 24, 100],
      [5, '', '3NMIGHPTTGC02N', 'Mì thịt thật gà cay 85grx24 NÐ (HP)', 'Thùng', 'HD', 24, 80],
      [6, '', '3NMIGHPTTGC03N', 'Mì thịt thật gà cay 85grx24 NÐ (ECOM) (HP)', 'Thùng', 'HD', 24, 60],
    ]);

    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, ws, 'QUY RA SỐ KHỐI THÁNG 9.2026');

    const buffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="mapping_template.xlsx"',
    });

    res.send(buffer);
  }

  @Post('detect-warehouse')
  detectWarehouse(@Body() body: { filename: string }, @Res() res: Response) {
    const filename = body.filename?.toLowerCase() || '';
    const warehouses = this.excelService.loadWarehouses();
    
    // Detect warehouse from filename
    let detectedWarehouse: Warehouse | null = null;
    
    if (filename.includes('swb ct') || filename.includes('can tho') || filename.includes('cần thơ')) {
      detectedWarehouse = warehouses.find((w: Warehouse) => w.Code === 'CT') || null;
    } else if (filename.includes('swb hn') || filename.includes('ha noi') || filename.includes('hà nội')) {
      detectedWarehouse = warehouses.find((w: Warehouse) => w.Code === 'HN') || null;
    } else if (filename.includes('swb hcm') || filename.includes('tp.hcm') || filename.includes('sg')) {
      detectedWarehouse = warehouses.find((w: Warehouse) => w.Code === 'HCM') || null;
    } else if (filename.includes('tiki') || filename.includes('mb') || filename.includes('tiger')) {
      detectedWarehouse = warehouses.find((w: Warehouse) => w.Code === 'TNSL_TIGER2') || null;
    } else if (filename.includes('baspro') || filename.includes('bas')) {
      detectedWarehouse = warehouses.find((w: Warehouse) => w.Code === 'BAS') || null;
    }

    res.json({ 
      detected: detectedWarehouse !== null,
      warehouse: detectedWarehouse,
      filename: body.filename
    });
  }

  @Post('process')
  @UseInterceptors(FileInterceptor('bangFile'))
  async processFiles(
    @UploadedFile() file: any,
    @Body() body: { warehouseId: string },
    @Res() res: Response
  ) {
    if (!file) {
      res.redirect('/');
      return;
    }

    const mappingDict = this.excelService.loadMappingToDict();
    if (mappingDict.size === 0) {
      res.redirect('/');
      return;
    }

    try {
      const warehouses = this.excelService.loadWarehouses();
      const selectedWarehouse = warehouses.find((w: Warehouse) => w.Id === parseInt(body.warehouseId));

      this.lastResult = this.excelService.processBangTheoDoi(file.path, mappingDict);
      this.lastResult.SelectedWarehouse = selectedWarehouse || null;

      // Delete uploaded file
      if (fs.existsSync(file.path)) {
        fs.unlinkSync(file.path);
      }
      
      res.render('result', { result: this.lastResult });
    } catch (error) {
      console.error('Process error:', error);
      // Delete uploaded file on error
      if (file.path && fs.existsSync(file.path)) {
        fs.unlinkSync(file.path);
      }
      res.redirect('/');
    }
  }

  @Get('download')
  downloadFile(@Query('xe') xeName: string, @Res() res: Response) {
    if (!this.lastResult) {
      res.redirect('/');
      return;
    }

    const buffer = this.excelService.generateHaravanFile(this.lastResult, xeName);
    const warehouseCode = this.lastResult.SelectedWarehouse?.Code || 'All';
    const fileName = xeName
      ? `Haravan_${warehouseCode}_${xeName}_${Date.now()}.xlsx`
      : `Haravan_${warehouseCode}_All_${Date.now()}.xlsx`;

    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${fileName}"`,
    });

    res.send(buffer);
  }
}
