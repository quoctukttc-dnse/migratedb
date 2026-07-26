# Dashboard Tiến độ Migration SO — SCAX → SCAF

Dashboard theo dõi tiến độ migration Sales Order từ SCAX sang SCAF, chi tiết theo từng RCM (người phụ trách), phục vụ chuẩn bị Golive SAP-ERP (Phase 2).

**File duy nhất `index.html`** — tự chứa toàn bộ (HTML/CSS/JS, kể cả thư viện Chart.js), không cần build, không phụ thuộc mạng ngoài. Mở trực tiếp file này bằng trình duyệt cũng xem được, hoặc publish qua GitHub Pages để có link chia sẻ.

Dữ liệu thô (JSON) đi kèm trong `data/dashboard_data.json` để tham khảo hoặc dùng lại cho báo cáo khác.

## Cách đẩy lên GitHub và bật GitHub Pages

Repo Git đã được khởi tạo sẵn trong file zip (đã có `.git/` và 1 commit) — chỉ cần trỏ về GitHub và push.

1. Tạo repo mới trên GitHub (ví dụ: `so-migration-dashboard`), để trống, **không** thêm README/license/.gitignore lúc tạo.
2. Giải nén file zip này ra một thư mục, mở terminal tại thư mục vừa giải nén (thư mục có sẵn `index.html`, `README.md`, `.git`) và chạy:

   ```bash
   git branch -M main
   git remote add origin https://github.com/<tên-tài-khoản>/<tên-repo>.git
   git push -u origin main
   ```

3. Vào repo trên GitHub → **Settings → Pages**.
4. Ở mục **Build and deployment**, chọn **Source: Deploy from a branch**, **Branch: main**, thư mục **/ (root)** → **Save**.
5. Sau khoảng 1–2 phút, GitHub sẽ cấp link dạng:
   `https://<tên-tài-khoản>.github.io/<tên-repo>/`
   — đây là link dashboard có thể chia sẻ cho cả team.

## Cập nhật dữ liệu sau này

Dashboard đọc dữ liệu tĩnh embedded sẵn trong `index.html` (biến `DATA` ở cuối file), lấy từ file `Danh Sách SO Cần Migration.xlsx` tại thời điểm tạo báo cáo (xem ngày trong header dashboard). Để cập nhật số liệu mới, tạo lại `index.html` từ dữ liệu Excel mới nhất và push lên GitHub — Pages sẽ tự deploy bản mới sau khi push.

## Nội dung dashboard

- **KPI tổng quan**: tổng số dòng SO cần migration, % đã chuyển sang SCAF, % đã upload SAP, tổng số SO, % SCAF Detail có PRI/BOM, số SO loại trừ (không migration).
- **Điểm nghẽn chính**: biểu đồ so sánh 3 mốc của luồng migration (tổng → đã sang SCAF → đã upload SAP) — cho thấy SCAF đã gần xong nhưng Upload SAP mới chỉ bắt đầu.
- **Chi tiết theo RCM — tồn đọng**: xếp hạng RCM theo số dòng SO chưa chuyển sang SCAF, ai cần ưu tiên xử lý trước.
- **Chi tiết theo RCM — % Upload SAP**: thước đo rủi ro Golive theo từng RCM (đỏ/vàng/xanh).
- **Bảng chi tiết đầy đủ**: sortable, filter theo tên, đủ các chỉ số cho từng RCM.

## Nguồn dữ liệu

`10. Transaction Data/3. Sales Order/Danh Sách SO Cần Migration.xlsx` — các sheet: `1. SO Monitor Migration`, `3. SO Khong Migration`, `2. SCAF SO Detail`.
