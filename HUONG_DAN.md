# Kiểm kê Freshshop — PWA chạy offline

## Cấu trúc
| File | Vai trò |
|---|---|
| `index.html`, `styles.css`, `app.js` | Giao diện và luồng kiểm |
| `parsers.js` | Đọc ASSET.xlsx, STOCK.csv (tự chuyển font TCVN3 → Unicode, tính CF), MASTERR.xlsx |
| `db.js` | Lưu offline bằng IndexedDB |
| `sw.js` | Service worker: cache app để chạy offline, nhận file chia sẻ (Android) |
| `manifest.webmanifest`, `icons/` | Cài lên màn hình chính, Web Share Target |
| `vendor/xlsx.full.min.js` | SheetJS 0.20.3 (đóng gói sẵn, không cần mạng) |

Danh sách cửa hàng nằm sẵn trong app: `data/MASTERR.xlsx` (trang GitHub Pages là công khai, ai có link đều tải được file này). File tài sản và tồn kho do người dùng chọn trên máy, chỉ lưu trong điện thoại.

## Khi master thay đổi
1. Sửa file `Input_Master\MASTERR.xlsx` như bình thường (giữ nguyên tên cột).
2. Chạy `cap_nhat_master.bat` (chép file vào `app\data\`).
3. Trên GitHub: vào thư mục `data` → **Add file → Upload files** → kéo `MASTERR.xlsx` vào → **Commit changes**.
4. Xong. Điện thoại mở app khi có mạng sẽ tự tải danh sách mới (dòng "Danh sách cửa hàng" ở cuối màn hình chính hiện số cửa hàng và giờ tải). Không cần sửa `VERSION` trong `sw.js`.

## Đưa lên GitHub Pages (làm lần đầu)
Lưu ý: GitHub Pages miễn phí là trang **công khai**. Ai có link đều mở được app và tải được `data/MASTERR.xlsx`. File STOCK, DVT, tài sản, ảnh và biên bản chỉ nằm trên điện thoại, không lên GitHub.

1. Vào https://github.com → **Sign up** (nếu chưa có tài khoản) → đăng nhập.
2. Góc phải trên bấm **+** → **New repository**.
   - Repository name: `kiemke-freshshop`
   - Chọn **Public**, không tích "Add a README" → **Create repository**.
3. Ở trang repo trống, bấm link **uploading an existing file**.
4. Mở File Explorer tới `D:\Claude\KiemkhoFrehshoppp`, bấm **Ctrl + A** (chọn hết file và các thư mục data, fonts, icons, report, vendor) rồi **kéo thả** vào khung upload trên trình duyệt (Chrome / Edge).
   - Kéo **nội dung bên trong** thư mục `app`, không kéo cả thư mục `app`.
   - Chờ hiện đủ khoảng 25 file (thư mục được giữ nguyên cấu trúc).
5. Kéo xuống dưới, bấm **Commit changes**.
6. Vào **Settings** → menu trái **Pages** → mục *Build and deployment*:
   - Source: **Deploy from a branch**
   - Branch: **main** / **(root)** → **Save**.
7. Chờ 1–3 phút (tab **Actions** hiện dấu ✓ xanh). Quay lại **Settings → Pages**, link app nằm ở dòng "Your site is live at":
   `https://<tên-github>.github.io/kiemke-freshshop/`
8. Mở link, kéo xuống cuối màn hình chính thấy "Phiên bản …" và số cửa hàng là thành công.

## Cập nhật app về sau
- **Sửa code app:** tăng `VERSION` trong `sw.js` (vd. `kk-v1.7.2`), sau đó vào repo → **Add file → Upload files** → kéo các file đã sửa vào (đúng thư mục) → **Commit changes**. Trên điện thoại: mở app khi có mạng, đóng hẳn rồi mở lại là nhận bản mới.
- **Sửa danh sách cửa hàng:** xem mục "Khi master thay đổi" (không cần đổi VERSION).

## Cài lên điện thoại
- **Android (Chrome):** mở link → menu ⋮ → *Thêm vào màn hình chính / Cài đặt ứng dụng*. Sau khi cài, trong app OneDrive bấm **Chia sẻ** file → chọn **Kiểm kê**.
- **iPhone (Safari):** mở link → nút Chia sẻ → *Thêm vào MH chính*. Trong app, bấm **Chọn file** → *Duyệt* → OneDrive.

## Cách dùng
**1. Chuyến kiểm mới** (màn hình đầu tiên)
- Dữ liệu: chọn file **Tồn kho** (STOCK.csv), **ĐVT tồn kho** (DVT.csv của cùng cửa hàng, bắt buộc khi kiểm tồn kho) và **Tài sản** (ASSET.xlsx tất cả cửa hàng, app tự lọc theo cửa hàng).
- Cửa hàng: app tự chọn theo OPER trong file STOCK / DVT; hoặc gõ tìm theo tên / mã. STOCK, DVT và cửa hàng đang chọn phải cùng OPER, sai thì không cho bắt đầu.
- Nhập tên người kiểm, chụp ảnh check-in → **Bắt đầu kiểm**.
- Mã tồn kho có CF, số lượng và trọng lượng đều bằng 0 được bỏ khỏi danh sách kiểm.

**2. Ba mục kiểm**
- **Kiểm hàng tồn kho:** kiểm theo ĐVT trong file DVT (Hộp, Khay, Bao, Kg…). Mỗi dòng bấm **Đúng** hoặc **Sai** (nhập số thực tế theo ĐVT, app quy đổi ra Quả / Kg; chọn lý do, ghi chú).
- **Kiểm tài sản:** bấm OK / Sai như trước.
- **Biên bản kiểm tra:** ảnh check-in, 10 tiêu chí, ghi chú Cửa hàng / Stock (nội dung + hướng xử lý) (tiêu chí Tiền mặt có bảng tính: doanh thu, tiền chuyển khoản/nộp, số tờ từng mệnh giá 500.000 → 500đ; app tự so và chọn Đạt / Không đạt, ghi chú do bạn tự ghi), hình ảnh kiểm tra kèm lý do, tên quản lý, ký tên → **Xác nhận & tạo PDF** → Tải / Gửi.

**3. Kết thúc chuyến kiểm** ở màn hình chính để xoá dữ liệu cửa hàng này và sang cửa hàng khác (nhớ gửi PDF trước).

## Biên bản PDF (theme Fresh Shop, khổ 792 × 612)
Bìa (logo CP Fresh Shop, ảnh check-in, người kiểm tra) · Nội dung chính · 1. Kiểm tra cửa hàng và tồn kho (ghi chú Cửa hàng / Stock bên trái, biên bản kiểm tra bên phải) · Biên bản kiểm kê hàng tồn kho · 2. Tài sản cố định · 3. Hình ảnh (3 ảnh một trang, chú thích dưới ảnh) · Thank You.
Ghi chú: mỗi mục gồm Nội dung + Hướng xử lý (HXL); mục để trống không đưa vào PDF. Viết "Trưng bày: ..." thì phần trước dấu ":" in đậm (mục Stock tô vàng).
Phần nào không có dữ liệu thì bỏ qua và mục lục tự đánh số lại. Font Tinos, Montserrat (OFL) trong `fonts/`.

## Chạy thử trên máy tính
```bash
python -m http.server 8080
```
Chạy trong thư mục `app`, mở http://localhost:8080
