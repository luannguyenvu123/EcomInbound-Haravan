# Export Haravan

Hệ thống xử lý và xuất file đơn hàng Haravan.

## Setup

```bash
npm install
```

## Chạy development

```bash
npm run start:dev
```

## Deploy lên Railway

1. Push code lên GitHub
2. Vào https://railway.app → Login with GitHub
3. Click "New Project" → "Deploy from GitHub repo"
4. Chọn repo này
5. Railway tự build và deploy

## Tính năng

- Upload file Excel "Bảng Theo Dõi"
- Auto-detect kho từ tên file
- Map mã 3N → 1N
- Tính số lượng thùng, lẻ, gói lẻ
- Export file Haravan theo từng xe
- Quản lý mapping (CRUD)
