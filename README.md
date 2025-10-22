## DOCX to PDF Converter (Microsoft Graph API)

Ứng dụng Node.js đơn giản cho phép người dùng đăng nhập bằng Microsoft Personal Account, upload file .docx và tải về dưới dạng PDF.

## Tính năng

- Đăng nhập bằng Microsoft Personal Account (MSA).
- Upload file DOCX từ máy tính.
- Convert file sang PDF bằng Microsoft Graph API.
- Xóa file tạm trên server và OneDrive sau khi convert.

## Yêu cầu

- Node.js >= 18

- Tài khoản Microsoft cá nhân (MSA, ví dụ Outlook.com)

- App đăng ký trên Azure với quyền Files.ReadWrite

## Cài đặt

Clone repository hoặc tải source code.

```bash
git clone <repo-url>
cd <project-folder>
```

Cài dependencies:

```bash
npm install
```

Tạo file .env:

```bash
PORT=5000
CLIENT_ID=YOUR_APP_CLIENT_ID
CLIENT_SECRET=YOUR_APP_CLIENT_SECRET
REDIRECT_URI=http://localhost:5000/auth/callback
SCOPES=Files.ReadWrite
```

- CLIENT_ID, CLIENT_SECRET: từ app Azure đã đăng ký.

- REDIRECT_URI: phải trùng với URI đăng ký trên Azure.

- SCOPES: Files.ReadWrite là đủ để upload, convert và xóa file.

Chạy ứng dụng

```bash
node server.js
```

Mở trình duyệt và truy cập: http://localhost:5000

Click Login with Microsoft và đồng ý quyền truy cập OneDrive.

Upload file .docx → nhận file PDF để tải về.

## Cấu trúc thư mục

```bash
project/
├─ uploads/ # file tạm upload
├─ public/
│ └─ index.html # giao diện web
├─ server.js # server Node.js
├─ .env # biến môi trường
└─ README.md
```

## Lưu ý

- File trên OneDrive chỉ tồn tại tạm thời trong quá trình convert.

- Nếu gặp lỗi về token hoặc quyền, kiểm tra app registration và scope.
