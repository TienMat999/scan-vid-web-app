# 📋 Scan VID Web App

Web application hỗ trợ scan và quản lý VID (Visual Inspection Data) cho quy trình sản xuất.

## 🌐 Environments

| Environment | Branch | URL | Version |
|---|---|---|---|
| Production | `main` | [scan-vid-web-app.vercel.app](https://scan-vid-web-app.vercel.app) | v0.0.1 |
| Preview | `preview` | Auto-generated bởi Vercel | v0.0.2-dev |

## 🔄 Workflow

```
preview branch ──push──▶ Vercel Preview URL ──test OK──▶ Pull Request ──merge──▶ main ──auto-deploy──▶ Production
```

1. Checkout nhánh `preview`
2. Code & test locally
3. Push lên `preview` → Vercel auto-deploy preview URL
4. Test trên preview URL
5. Tạo Pull Request từ `preview` → `main`
6. Review & merge → Auto-deploy production

## 🛠️ Development

```bash
# Clone repo
git clone git@github.com:TienMat999/scan-vid-web-app.git

# Chuyển sang nhánh preview để phát triển
git checkout preview

# Mở file index.html trực tiếp hoặc dùng live server
```

## 🧪 Testing

```bash
# Chạy bộ test tự động (cần Node.js + npm install)
node test.js
```

## 📦 Tech Stack

- **Frontend**: Vanilla HTML5, CSS3, JavaScript (ES6+)
- **Library**: SheetJS (xlsx) — parsing/export Excel
- **Hosting**: Vercel (static deployment)
- **Testing**: Puppeteer (14 test groups)

## 👤 Owner

TANG, HUYNH MINH TIEN