# SCP Dashboard – đăng nhập Microsoft SSO, mỗi team lead chỉ thấy team của mình

Một trang `index.html` (không chứa dữ liệu). Sau khi đăng nhập, API kiểm tra email trong `Dashboard Access.xlsx` và chỉ trả dữ liệu đúng team của người đó. Master (sheet `Master access`) thấy mọi team qua ô chọn team. Giống Goal Dashboard.

## Cập nhật hằng kỳ
Thả file mới vào folder `SCP Dashboard` trên SharePoint (People Insights Hub), không tạo thêm folder:
- `OneOnOnes_YYYY_MM_DD_HHMMSS.xlsx` (sheet Manager Levels) – workflow lấy file có ngày mới nhất
- `User Initiated 1-on-1s*.xlsx` (sheet Data) – thả thẳng file vừa tải về, **không cần đổi tên**. Mỗi export chứa toàn bộ lịch sử nên thường chỉ cần giữ bản mới nhất (ghi đè cùng tên), nhưng thả thêm file cũ/mới vẫn được: dashboard đọc tất cả, sắp theo ngày "Exported On" trong sheet Overview, dòng trùng (cùng nhân viên + ngày launch) lấy bản export mới nhất. Ngày hiển thị "SCP data as of" cũng lấy từ "Exported On"
- `Dashboard Access.xlsx` – ai xem team nào (sheet Views: tên team lead + email, cách nhau bằng `;`; sheet Master access)
- Danh sách nhân viên dùng chung nằm ở thư mục gốc People Insights Hub (`Danh sach nhan vien*.xlsx`)

Trong vòng 1 giờ dashboard tự cập nhật (hoặc chạy tay workflow, tick Force). Người xem tải lại trang.

## Quy tắc tính (xem đầu file `build_data.py`)
- Team = mọi người dưới team lead ở mọi cấp (Manager Levels). Supervisor = cấp quản lý trực tiếp.
- New hire dưới 2 tháng (tính đến ngày dữ liệu): chỉ tính vào Total Headcount, hiện tên kèm nhãn "New hire", không vào số liệu khác.
- Mọi số theo kỳ (Period bar): tháng tính theo ngày nhân viên share check-in. Weeks submitted = số tuần (Thứ Hai–Chủ Nhật, tuần đã kết thúc) có ít nhất một check-in được share.
- Topics ở tab Themes: tự đếm cụm từ / tên khách hàng trong comment (không dùng danh sách chủ đề cố định); `themes.py` không còn dùng để hiển thị.

## File trong folder Code (tất cả đưa lên GitHub, trừ legacy-static-html)
`index.html`, `build.py`, `build_data.py`, `themes.py`, `build_kit.py`, `fetch_kit.py`, `requirements.txt`, `.gitignore`, `README.md`, `api_src/`, `.github/workflows/refresh.yml`.
Không đưa lên GitHub: `legacy-static-html/` (chứa dữ liệu thật), mọi file Excel, `SCP_preview_*.html` (bản xem thử có dữ liệu thật).

## Cấu hình GitHub / Azure
- Variables: `SITE_URL` = `https://tcapc.sharepoint.com/sites/performance`; `FOLDER` = `SCP Dashboard`; `SHAREPOINT_ORIGIN` = `https://tcapc.sharepoint.com`;
  `FILES` = `{"access":"Dashboard Access\\.xlsx","oneonones":"OneOnOnes_(\\d{4}_\\d{2}_\\d{2})_\\d+\\.xlsx","scp[all]":"User Initiated 1-on-1s.*\\.xlsx"}`
  `SHARED_FILES` = `{"employees":"Danh sach nhan vien.*\\.xlsx"}`
- Secrets: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` (giống Goal Dashboard), `AZURE_STATIC_WEB_APPS_API_TOKEN` (của Static Web App **mới** cho SCP).
- App VN Office Page: thêm Redirect URI loại Single-page application = URL của Static Web App SCP.
