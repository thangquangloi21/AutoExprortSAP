# MB52.vbs Login Popup Change Plan

## Muc tieu
- Gia co buoc dang nhap SAP trong `Script/MB52.vbs`.
- Neu xuat hien popup "License Information for Multiple Logons", tu dong chon dong thu 2:
  "Continue with this logon, without ending any other logons in the system".
- Khong thay doi luong chay MB52 va export file `D:\4.DEV\Python\XuatDuLieuSAP\Data\MB52.XLSX`.

## Suy nghi truoc khi sua
- Script da co xu ly popup Multiple Logons bang `wnd[1]/usr/radMULTI_LOGON_OPT2`.
- Do popup co the xuat hien cham hoac nam ngoai nhanh login hien tai, nen can tach thanh helper rieng co vong cho ngan.
- Helper se bo qua an toan neu popup khong xuat hien, tranh lam loi cac lan dang nhap binh thuong.

## Thay doi du kien
- Them helper `HandleMultipleLogonPopup(ByRef sapSession, ByRef sapConnection)`.
- Helper refresh session, cho toi da 5 giay, chon `radMULTI_LOGON_OPT2`, roi bam OK.
- Goi helper sau buoc dang nhap/SSO.
- Goi helper them mot lan truoc khi chay transaction MB52.

## Da thuc hien
- Da them helper `HandleMultipleLogonPopup` vao `Script/MB52.vbs`.
- Da thay doan xu ly popup Multiple Logon inline bang loi goi helper.
- Da goi helper them mot lan truoc khi chay transaction `MB52`.
- Da them `Script/sap_login_config.json` de luu `environment`, `client`, `user`.
- Da cap nhat `Script/MB52.vbs` doc `environment`, `client`, `user` tu file JSON, co gia tri mac dinh neu file bi thieu.
- Da them nhom `paths` vao `Script/sap_login_config.json`: `sap_logon_exe`, `log_file`, `export_folder`, `export_file`.
- Da cap nhat `Script/MB52.vbs` de dung cac duong dan va ten file tu JSON khi mo SAP Logon, ghi log, export MB52 va dong workbook Excel.
- Da them `vbs_script` va `csv_file` vao `Script/sap_login_config.json`.
- Da cap nhat `main.py` de doc file VBS input, file Excel input va file CSV output tu JSON; giu nguyen connect SQL va tham so BCP SQL.
- Da cap nhat `main.py` de ho tro chay sau khi build PyInstaller one-file: uu tien config ngoai canh exe, fallback config nhung trong exe.
- Da build `dist/XuatDuLieuSAP.exe` bang PyInstaller `--onefile --noconsole` va dat them ban config/script ngoai trong `dist/Script` de co the sua sau build.
- Da them `paths.log_root` vao `Script/sap_login_config.json`.
- Da cap nhat `main.py` de ghi log tien trinh vao cau truc `log_root/YYYY/MM/DD/XuatDuLieuSAP_YYYYMMDD_HHMMSS.log`.
- Da rebuild `dist/XuatDuLieuSAP.exe` sau khi them logging.
- Da doi cau truc log thanh `log_root/YYYY/MM/XuatDuLieuSAP_YYYYMMDD.log`, moi ngay mot file va cac lan chay trong ngay ghi tiep vao cung file.
- Da rebuild `dist/XuatDuLieuSAP.exe` sau khi doi cau truc log.

## Kiem tra
- Truong hop khong co popup: script tiep tuc vao MB52.
- Truong hop co popup Multiple Logons: script chon dong thu 2 va bam OK.
- Khong doi connection, client, user, transaction, thu muc export, ten file export.
