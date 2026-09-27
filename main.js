const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn, exec } = require('child_process'); // 🔴 Thêm module để gọi Excel ngầm

let mainWindow;

// 28.09.2026 - WebApp mới ưu tiên VPS, CodeSandbox chỉ làm dự phòng.
const PRIMARY_WEB_URL = 'https://103.20.102.220/webapp-test/';
const FALLBACK_WEB_URL = 'https://gz7242.csb.app';
const PRIMARY_LOAD_TIMEOUT_MS = 15000;
let webLoadTarget = 'primary';
let primaryLoadTimer = null;

function clearPrimaryLoadTimer() {
    if (primaryLoadTimer) {
        clearTimeout(primaryLoadTimer);
        primaryLoadTimer = null;
    }
}

function offlineHtml(reason) {
    const safeReason = String(reason || 'Không kết nối được Web App').replace(/[<>&"']/g, '');
    return '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>Nhiêu Lộc - Mất kết nối</title><style>body{font-family:Segoe UI,Arial,sans-serif;background:#f7f7f7;color:#222;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}' +
        '.box{background:#fff;border:1px solid #ddd;border-radius:12px;padding:28px;max-width:620px;box-shadow:0 8px 28px rgba(0,0,0,.08)}h2{margin-top:0}' +
        'a{display:inline-block;margin:8px 8px 0 0;padding:10px 14px;border-radius:8px;text-decoration:none;background:#1266d4;color:white}.secondary{background:#555}.hint{color:#666;font-size:13px}</style>' +
        '</head><body><div class="box"><h2>Web App Nhiêu Lộc chưa kết nối được</h2><p>' + safeReason + '</p>' +
        '<p>Ứng dụng đã thử VPS và CodeSandbox nhưng cả hai đều chưa mở được.</p>' +
        '<a href="' + PRIMARY_WEB_URL + '">Thử lại VPS</a><a class="secondary" href="' + FALLBACK_WEB_URL + '">Mở CodeSandbox</a>' +
        '<p class="hint">Dữ liệu Google/XUAT không bị xóa. Có thể đóng và mở lại ứng dụng để thử lại.</p></div></body></html>';
}

function showOfflinePage(reason) {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    clearPrimaryLoadTimer();
    webLoadTarget = 'offline';
    const url = 'data:text/html;charset=utf-8,' + encodeURIComponent(offlineHtml(reason));
    mainWindow.loadURL(url).catch(() => {});
}

function loadFallback(reason) {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    if (webLoadTarget === 'fallback') return;
    clearPrimaryLoadTimer();
    webLoadTarget = 'fallback';
    console.log('VPS WebApp lỗi, chuyển CodeSandbox:', reason || 'unknown');
    mainWindow.loadURL(FALLBACK_WEB_URL).catch((err) => {
        showOfflinePage('CodeSandbox cũng không mở được: ' + (err && err.message ? err.message : err));
    });
}

function loadPrimaryWebApp() {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    clearPrimaryLoadTimer();
    webLoadTarget = 'primary';
    console.log('Mở WebApp VPS:', PRIMARY_WEB_URL);

    mainWindow.loadURL(PRIMARY_WEB_URL).catch((err) => {
        if (webLoadTarget === 'primary') {
            loadFallback(err && err.message ? err.message : err);
        }
    });

    primaryLoadTimer = setTimeout(() => {
        if (!mainWindow || mainWindow.isDestroyed() || webLoadTarget !== 'primary') return;
        if (mainWindow.webContents.isLoadingMainFrame()) {
            console.log('VPS quá 15 giây, dùng CodeSandbox dự phòng');
            webLoadTarget = 'switching';
            try { mainWindow.webContents.stop(); } catch (_) {}
            loadFallback('VPS timeout quá 15 giây');
        }
    }, PRIMARY_LOAD_TIMEOUT_MS);
}

// Đường dẫn lưu file config 2 máy in
const configPath = path.join(app.getPath('userData'), 'printerConfig.json');

// 🔴 Biến chứa tiến trình Zombie Excel
let excelZombie;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1280, 
        height: 800,
        autoHideMenuBar: true,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,  
            nodeIntegration: false   
        }
    });

    // Load WebApp: VPS trước, CodeSandbox dự phòng.
    mainWindow.webContents.on('did-finish-load', () => {
        clearPrimaryLoadTimer();
        console.log('WebApp đã tải:', webLoadTarget, mainWindow.webContents.getURL());
    });

    mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription, validatedURL, isMainFrame) => {
        if (!isMainFrame || errorCode === -3) return; // -3 = ERR_ABORTED khi chủ động chuyển URL
        if (webLoadTarget === 'primary') {
            loadFallback(`${errorCode} ${errorDescription}`);
        } else if (webLoadTarget === 'fallback') {
            showOfflinePage(`CodeSandbox lỗi ${errorCode}: ${errorDescription}`);
        }
    });

    // Hai nút trên trang mất kết nối dùng chính cơ chế fallback của app.
    mainWindow.webContents.on('will-navigate', (event, url) => {
        if (webLoadTarget !== 'offline') return;
        if (url === PRIMARY_WEB_URL) {
            event.preventDefault();
            loadPrimaryWebApp();
        } else if (url === FALLBACK_WEB_URL || url === FALLBACK_WEB_URL + '/') {
            event.preventDefault();
            webLoadTarget = 'fallback';
            mainWindow.loadURL(FALLBACK_WEB_URL).catch((err) => showOfflinePage(String(err)));
        }
    });

    loadPrimaryWebApp();
}

app.whenReady().then(() => {
    createWindow();

    // 🔴 NUÔI ZOMBIE EXCEL: Mở ngầm 1 process Excel tàng hình khi bật App
    excelZombie = spawn('powershell.exe', [
        '-Command', 
        '$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false; while($true){ Start-Sleep -Seconds 60 }'
    ]);
    console.log("🧟 Đã nuôi sẵn Zombie Excel chạy nền!");
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// 🔴 DỌN DẸP ZOMBIE: Tắt App thì phải giết Zombie để giải phóng RAM cho máy tính
app.on('will-quit', () => {
    if (excelZombie) excelZombie.kill();
    exec('taskkill /f /im excel.exe'); 
});

/* =========================================================================
   BỘ XỬ LÝ MÁY IN (2 MÁY IN)
   ========================================================================= */

// 1. Quét máy in
ipcMain.handle('get-printers', async (event) => {
    try {
        const webContents = event.sender;
        if (typeof webContents.getPrintersAsync === 'function') {
            return await webContents.getPrintersAsync();
        } else if (typeof webContents.getPrinters === 'function') {
            return webContents.getPrinters();
        }
        return [];
    } catch (error) {
        console.log("Lỗi quét máy in:", error);
        return [];
    }
});

// 2. Lưu cấu hình 2 máy in
ipcMain.on('save-printers-config', (event, config) => {
    try {
        fs.writeFileSync(configPath, JSON.stringify(config));
        console.log("Đã lưu cấu hình máy in:", config);
    } catch (error) {
        console.log("Lỗi lưu cấu hình:", error);
    }
});

// 3. Đọc cấu hình 2 máy in
ipcMain.handle('get-saved-printers', () => {
    if (fs.existsSync(configPath)) {
        try {
            return JSON.parse(fs.readFileSync(configPath));
        } catch (error) {
            return { epson: "", a4: "" };
        }
    }
    return { epson: "", a4: "" };
});



// =========================================================================
// IN HTML NGẦM CHO PHIẾU CHUYỂN KHO / CÁC MẪU HTML A5
// Dùng máy in được chọn trong app, không bật hộp thoại in.
// =========================================================================

// =========================================================================
// IN PHIẾU CHUYỂN KHO BẰNG FILE MẪU EXCEL
// File mẫu: ChuyenKho_Template.xlsx
// Sheet CKAB / CKNB, in A5 ngang, fit 1 page.
// =========================================================================
ipcMain.handle('in-chuyen-kho-excel-nong', async (event, data) => {
    const tempFile = path.join(app.getPath('userData'), 'temp_chuyen_kho.json');
    const templatePath = path.join(__dirname, 'ChuyenKho_Template.xlsx');

    if (!fs.existsSync(templatePath)) {
        console.log("❌ KHÔNG TÌM THẤY FILE MẪU: ChuyenKho_Template.xlsx");
        return { ok: false, error: "Không tìm thấy ChuyenKho_Template.xlsx" };
    }

    fs.writeFileSync(tempFile, JSON.stringify(data), 'utf8');

    const psScript = `
        $ErrorActionPreference = 'Stop'
        $data = Get-Content '${tempFile}' -Encoding UTF8 -Raw | ConvertFrom-Json

        function ToNum($v) {
            if ($null -eq $v) { return "" }
            $s = "$v".Trim()
            if ($s -eq "") { return "" }
            $s = $s.Replace(",", ".")
            $n = 0
            if ([double]::TryParse($s, [ref]$n)) { return $n }
            return $v
        }

        function PlFromThung($thung, $qcach) {
            $t = ToNum $thung
            $q = ToNum $qcach
            if ($q -eq "" -or $q -eq 0) { return "" }
            return [math]::Floor([double]$t / [double]$q)
        }

        function SetupPage($ws, $printArea) {
            $ws.PageSetup.Orientation = 2 # xlLandscape
            $ws.PageSetup.PaperSize = 11 # xlPaperA5
            $ws.PageSetup.Zoom = $false
            $ws.PageSetup.FitToPagesWide = 1
            $ws.PageSetup.FitToPagesTall = 1
            $ws.PageSetup.LeftMargin = $ws.Application.CentimetersToPoints(0.2)
            $ws.PageSetup.RightMargin = $ws.Application.CentimetersToPoints(0.2)
            $ws.PageSetup.TopMargin = $ws.Application.CentimetersToPoints(0.2)
            $ws.PageSetup.BottomMargin = $ws.Application.CentimetersToPoints(0.2)
            $ws.PageSetup.HeaderMargin = $ws.Application.CentimetersToPoints(0)
            $ws.PageSetup.FooterMargin = $ws.Application.CentimetersToPoints(0)
            $ws.PageSetup.PrintArea = $printArea
        }

        function ClearCKAB($ws) {
            $ws.Range("A6:N12").ClearContents()
            $ws.Range("A6:N12").Font.Bold = $false
            $ws.Range("A6:N12").Font.ColorIndex = 1
            $ws.Range("N6:N12").ClearContents()
            $ws.Range("G13:H13").ClearContents()
        }

        function ClearCKNB($ws) {
            $ws.Range("A6:J12").ClearContents()
            $ws.Range("A6:J12").Font.Bold = $false
            $ws.Range("A6:J12").Font.ColorIndex = 1
            $ws.Range("G13:H13").ClearContents()
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        try {
            $wb = $excel.Workbooks.Open('${templatePath}', [Type]::Missing, $true)

            foreach ($page in $data.pages) {
                $loai = "$($page.loai)"
                if ([string]::IsNullOrWhiteSpace($loai)) { $loai = "CKNB" }
                $loai = $loai.ToUpper()

                if ($loai -eq "CKAB") {
                    $ws = $wb.Sheets.Item("CKAB")
                    ClearCKAB $ws
                    SetupPage $ws "A1:N18"

                    $ws.Cells.Item(1,14).Value = "$($page.timeStr)"
                    $ws.Cells.Item(2,14).Value = "$($page.ngay)"
                    $ws.Cells.Item(3,14).Value = "$($page.soPhieu)"
                    $ws.Cells.Item(14,2).Value = "$($page.ngay)"

                    $row = 6
                    foreach ($item in $page.items) {
                        if ($row -gt 12) { break }
                        $qc = ToNum $item.qcach
                        $tonP = PlFromThung $item.tonTruocDi $item.qcach

                        $ws.Cells.Item($row,1).Value = "$($item.stt)"
                        $ws.Cells.Item($row,2).Value = "$($item.maHang)"
                        $ws.Cells.Item($row,3).Value = "$($item.hieu)"
                        $ws.Cells.Item($row,4).Value = $qc
                        $ws.Cells.Item($row,5).Value = $tonP
                        $ws.Cells.Item($row,6).Value = ToNum $item.tonTruocDi
                        $ws.Cells.Item($row,7).Value = ToNum $item.soPL
                        $ws.Cells.Item($row,8).Value = ToNum $item.tongThung
                        $ws.Cells.Item($row,9).Value = "$($item.fromViTri)"
                        $ws.Cells.Item($row,10).Value = ToNum $item.tonTruocDen
                        $ws.Cells.Item($row,11).Value = ToNum $item.soPL
                        $ws.Cells.Item($row,12).Value = ToNum $item.tongThung
                        $ws.Cells.Item($row,13).Value = "$($item.toViTri)"
                        $ws.Cells.Item($row,14).Value = "$($item.ghiChu)"
                        $row++
                    }

                    $ws.Cells.Item(13,7).Value = ToNum $page.totalPL
                    $ws.Cells.Item(13,8).Value = ToNum $page.totalThung
                    $ws.PrintOut([Type]::Missing, [Type]::Missing, 1, $false, $data.mayIn)
                } else {
                    $ws = $wb.Sheets.Item("CKNB")
                    ClearCKNB $ws
                    SetupPage $ws "A1:J18"

                    $ws.Cells.Item(1,10).Value = "$($page.timeStr)"
                    $ws.Cells.Item(2,10).Value = "$($page.ngay)"
                    $ws.Cells.Item(3,10).Value = "$($page.soPhieu)"
                    $ws.Cells.Item(14,2).Value = "$($page.ngay)"

                    $row = 6
                    foreach ($item in $page.items) {
                        if ($row -gt 12) { break }
                        $qc = ToNum $item.qcach
                        $tonP = PlFromThung $item.tonTruocDi $item.qcach

                        $ws.Cells.Item($row,1).Value = "$($item.stt)"
                        $ws.Cells.Item($row,2).Value = "$($item.maHang)"
                        $ws.Cells.Item($row,3).Value = "$($item.hieu)"
                        $ws.Cells.Item($row,4).Value = $qc
                        $ws.Cells.Item($row,5).Value = $tonP
                        $ws.Cells.Item($row,6).Value = ToNum $item.tonTruocDi
                        $ws.Cells.Item($row,7).Value = ToNum $item.soPL
                        $ws.Cells.Item($row,8).Value = ToNum $item.tongThung
                        $ws.Cells.Item($row,9).Value = "$($item.fromViTri)"
                        $ws.Cells.Item($row,10).Value = "$($item.toViTri)"
                        $row++
                    }

                    $ws.Cells.Item(13,7).Value = ToNum $page.totalPL
                    $ws.Cells.Item(13,8).Value = ToNum $page.totalThung
                    $ws.PrintOut([Type]::Missing, [Type]::Missing, 2, $false, $data.mayIn)
                }
            }

        } finally {
            if ($wb) { $wb.Close($false) }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    `;

    const psFile = path.join(app.getPath('userData'), 'print_job_chuyen_kho.ps1');
    fs.writeFileSync(psFile, '\ufeff' + psScript, 'utf8');

    return new Promise((resolve) => {
        exec(`powershell.exe -ExecutionPolicy Bypass -File "${psFile}"`, (error, stdout, stderr) => {
            if (error) {
                console.log("❌ Lỗi in Chuyển Kho Excel:", error);
                resolve({ ok: false, error: error.toString() });
            } else {
                console.log("✅ Đã in Chuyển Kho bằng Excel thành công!");
                resolve({ ok: true });
            }
        });
    });
});


ipcMain.handle('in-html-ngam', async (event, data) => {
    return new Promise((resolve) => {
        try {
            const html = data && data.html ? data.html : "";
            const deviceName = data && data.deviceName ? data.deviceName : "";
            const landscape = data && data.landscape !== undefined ? !!data.landscape : true;
            const pageSize = data && data.pageSize ? data.pageSize : { width: 210000, height: 148000 };

            if (!html) {
                console.log("❌ in-html-ngam: Thiếu HTML");
                resolve({ ok: false, error: "Thiếu HTML" });
                return;
            }

            if (!deviceName) {
                console.log("❌ in-html-ngam: Thiếu tên máy in");
                resolve({ ok: false, error: "Thiếu tên máy in" });
                return;
            }

            const printWin = new BrowserWindow({
                show: false,
                width: 1200,
                height: 900,
                webPreferences: {
                    nodeIntegration: false,
                    contextIsolation: true,
                    sandbox: false
                }
            });

            const safeHtml = html.includes("<html")
                ? html
                : `<!doctype html><html><head><meta charset="utf-8"></head><body>${html}</body></html>`;

            printWin.loadURL("data:text/html;charset=utf-8," + encodeURIComponent(safeHtml));

            printWin.webContents.on("did-finish-load", () => {
                setTimeout(() => {
                    printWin.webContents.print({
                        silent: true,
                        printBackground: true,
                        deviceName: deviceName,
                        landscape: landscape,
                        margins: { marginType: "none" },
                        pageSize: pageSize
                    }, (success, failureReason) => {
                        if (!success) {
                            console.log("❌ Lỗi in HTML ngầm:", failureReason);
                            resolve({ ok: false, error: failureReason || "print failed" });
                        } else {
                            console.log("✅ Đã in HTML ngầm tới:", deviceName);
                            resolve({ ok: true });
                        }

                        try {
                            printWin.close();
                        } catch (e) {}
                    });
                }, 450);
            });

            printWin.webContents.on("did-fail-load", (event, errorCode, errorDescription) => {
                console.log("❌ Không load được HTML để in:", errorDescription);
                resolve({ ok: false, error: errorDescription });
                try {
                    printWin.close();
                } catch (e) {}
            });
        } catch (err) {
            console.log("❌ Exception in-html-ngam:", err);
            resolve({ ok: false, error: err.toString() });
        }
    });
});

/// =========================================================================
// 4. HÀM IN EXCEL TỐC ĐỘ CAO (QUY ĐỔI HẾT VỀ SỐ NGUYÊN ĐỂ SUM CHUẨN)
// =========================================================================
ipcMain.handle('in-excel-nong', async (event, data) => {
    const tempFile = path.join(app.getPath('userData'), 'temp_lenhkho.json');
    const templatePath = path.join(__dirname, 'LenhKho_Template.xlsx'); 

    if (!fs.existsSync(templatePath)) {
        console.log("❌ KHÔNG TÌM THẤY FILE MẪU: LenhKho_Template.xlsx");
        return;
    }

    fs.writeFileSync(tempFile, JSON.stringify(data), 'utf8');

    const psScript = `
        $ErrorActionPreference = 'Stop'
        $data = Get-Content '${tempFile}' -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # ======================================================
        # 🔥 ĐÃ NÂNG CẤP: TẠO EXCEL "TÀNG HÌNH" ĐỘC LẬP
        # (Không đụng chạm tới các file Excel sếp đang mở)
        # ======================================================
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false        # Giấu hoàn toàn cửa sổ
        $excel.DisplayAlerts = $false  # Tắt mọi cảnh báo gây phiền
        
        try {
            $wb = $excel.Workbooks.Open('${templatePath}', [Type]::Missing, $true)
            $ws = $wb.Sheets.Item(1)

            # ======================================================
            # ✅ ĐIỀN TÊN KẾ TOÁN VÀO EXCEL TEMPLATE
            # Trong file mẫu chỉ cần đặt chữ: {{KE_TOAN}}
            # Sỉ = Thúy Thuong | Lẻ = Bạch Kim
            # ======================================================
            $keToan = "$($data.keToan)"
            if ([string]::IsNullOrWhiteSpace($keToan)) {
                $flowText = ("$($data.flow) $($data.appFlow) $($data.luongXuat)").ToUpper()
                if ($flowText.Contains("LE") -or $flowText.Contains("LẺ")) {
                    $keToan = "Bạch Kim"
                } else {
                    $keToan = "Thúy Thuong"
                }
            }
            foreach ($sh in @($wb.Worksheets)) {
                try {
                    $usedRange = $sh.UsedRange
                    $usedRange.Replace("{{KE_TOAN}}", $keToan, 2, 1, $false, $false, $false, $false) | Out-Null
                } catch {}
            }
            
            # ĐIỀN HEADER
            $ws.Cells.Item(1, 9).Value = "$($data.ngayStr)"
            $ws.Cells.Item(2, 9).Value = "$($data.timeStr)"
            $ws.Cells.Item(3, 5).Value = "$($data.orderCode)"
            $ws.Cells.Item(4, 3).Value = "$($data.maKH)"
            $ws.Cells.Item(5, 3).Value = "$($data.tenKH)"
            $ws.Cells.Item(4, 9).Value = "$($data.soXe)"
            
            # ======================================================
            foreach ($page in $data.pages) {
                
                # 1. TẨY TRẮNG CẢ DỮ LIỆU LẪN MÀU SẮC (Để không bị dính màu đỏ từ tờ trước)
                $wsRange = $ws.Range("A7:K16")
                $wsRange.ClearContents()
                $wsRange.Font.Bold = $false
                $wsRange.Font.ColorIndex = 1 # Trả về màu đen mặc định
                
                $row = 7
                foreach ($item in $page.items) {
                    
                    if ($item.isNote) {
                        # NẾU LÀ GHI CHÚ: Chỉ điền vào cột Tên Hàng và tô Đỏ Đậm
                        $ws.Cells.Item($row, 3).Value = "$($item.tenHang)"
                        $ws.Cells.Item($row, 3).Font.Bold = $true
                        $ws.Cells.Item($row, 3).Font.Color = 255 # Màu Đỏ
                    } else {
                        # NẾU LÀ HÀNG THẬT: Điền đầy đủ các cột
                        $ws.Cells.Item($row, 1).Value = "$($item.stt)"
                        $ws.Cells.Item($row, 2).Value = "$($item.maHang)"
                        $ws.Cells.Item($row, 3).Value = "$($item.tenHang)"
                        $ws.Cells.Item($row, 4).Value = "$($item.hieu)"
                        $ws.Cells.Item($row, 6).Value = "$($item.plHienThi)"
                        $ws.Cells.Item($row, 10).Value = "$($item.viTri)"
                        $ws.Cells.Item($row, 11).Value = "$($item.tonKho)"
                        
                        if ("$($item.qcach)" -ne "") { $ws.Cells.Item($row, 5).Value2 = [double]$item.qcach }
                        if ("$($item.soThung)" -ne "") { $ws.Cells.Item($row, 7).Value2 = [double]$item.soThung }
                        if ("$($item.qcKgT)" -ne "") { $ws.Cells.Item($row, 8).Value2 = [double]$item.qcKgT }
                        if ("$($item.trongLuong)" -ne "") { $ws.Cells.Item($row, 9).Value2 = [double]$item.trongLuong }
                    }
                    $row++
                }
                
                # IN 3 TỜ CHO MỖI LỆNH TÁCH
                $ws.PrintOut([Type]::Missing, [Type]::Missing, 3, $false, $data.mayIn)
            }
            
        } finally {
            # ======================================================
            # 🔥 TỰ ĐỘNG DỌN DẸP XÁC EXCEL SAU KHI IN XONG
            # ======================================================
            if ($wb) { $wb.Close($false) }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    `;

    const psFile = path.join(app.getPath('userData'), 'print_job_lenh.ps1');
    fs.writeFileSync(psFile, '\ufeff' + psScript, 'utf8'); 
    
    exec(`powershell.exe -ExecutionPolicy Bypass -File "${psFile}"`, (error, stdout, stderr) => {
        if (error) {
            console.log("❌ Lỗi in Excel:", error);
        } else {
            console.log("✅ Đã in Lệnh xuất kho thành công!");
        }
    });
});
// =========================================================================
// 5. HÀM IN PHIẾU XUẤT TỐC ĐỘ CAO BẰNG EXCEL (XE ỦI ĐẤT CHỐNG SẬP)
// =========================================================================
ipcMain.handle('in-px-excel-nong', async (event, data) => {
    const tempFile = path.join(app.getPath('userData'), 'temp_px.json');
    const templatePath = path.join(__dirname, 'PhieuXuat_Template.xlsx');

    if (!fs.existsSync(templatePath)) {
        console.log("❌ KHÔNG TÌM THẤY FILE MẪU: PhieuXuat_Template.xlsx");
        return;
    }

    data.items.forEach(item => {
        const epThanhSoNguyen = (val) => {
            if (val === "" || val == null) return "";
            let n = Number(val.toString().replace(/,/g, '.')); 
            return isNaN(n) ? "" : Math.round(n);
        };
        item.soThung = epThanhSoNguyen(item.soThung);
        item.trongLuong = epThanhSoNguyen(item.trongLuong);
    });

    fs.writeFileSync(tempFile, JSON.stringify(data), 'utf8');

    const psScript = `
        # 🔴 LỆNH XE ỦI: Bất chấp lỗi gộp ô ở tiêu đề, cứ lẳng lặng điền tiếp, không được sập!
        $ErrorActionPreference = 'SilentlyContinue' 
        
        $data = Get-Content '${tempFile}' -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # ======================================================
        # 🔥 ĐÃ NÂNG CẤP: TẠO EXCEL "TÀNG HÌNH" ĐỘC LẬP CHO PHIẾU XUẤT
        # ======================================================
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        try {
            $wb = $excel.Workbooks.Open('${templatePath}', [Type]::Missing, $true)
            $ws = $wb.Sheets.Item(1)

            # ======================================================
            # ✅ ĐIỀN TÊN KẾ TOÁN VÀO EXCEL TEMPLATE
            # Trong file mẫu chỉ cần đặt chữ: {{KE_TOAN}}
            # Sỉ = Thúy Thuong | Lẻ = Bạch Kim
            # ======================================================
            $keToan = "$($data.keToan)"
            if ([string]::IsNullOrWhiteSpace($keToan)) {
                $flowText = ("$($data.flow) $($data.appFlow) $($data.luongXuat)").ToUpper()
                if ($flowText.Contains("LE") -or $flowText.Contains("LẺ")) {
                    $keToan = "Bạch Kim"
                } else {
                    $keToan = "Thúy Thuong"
                }
            }
            foreach ($sh in @($wb.Worksheets)) {
                try {
                    $usedRange = $sh.UsedRange
                    $usedRange.Replace("{{KE_TOAN}}", $keToan, 2, 1, $false, $false, $false, $false) | Out-Null
                } catch {}
            }
            
            # Khúc này lỡ tọa độ (Dòng, Cột) không khớp ô gộp nó cũng tự bỏ qua, không chết code
            $ws.Cells.Item(1, 5).Value = "$($data.ngayStr)"
            $ws.Cells.Item(2, 5).Value = "$($data.timeStr)"
            $ws.Cells.Item(3, 5).Value = "$($data.orderCode)"
            $ws.Cells.Item(4, 3).Value = "$($data.maKH)"
            $ws.Cells.Item(5, 3).Value = "$($data.tenKH)"
            $ws.Cells.Item(4, 5).Value = "$($data.soXe)"
            
            $ws.Range("A7:E16").ClearContents()
            
            $row = 7
            foreach ($item in $data.items) {
                # Khúc này điền bảng chắc chắn sống vì không có gộp ô
                $ws.Cells.Item($row, 1).Value = "$($item.stt)"
                $ws.Cells.Item($row, 2).Value = "$($item.maHang)"
                $ws.Cells.Item($row, 3).Value = "$($item.tenHang)"
                
                if ("$($item.soThung)" -ne "") { $ws.Cells.Item($row, 4).Value2 = [int]$item.soThung }
                if ("$($item.trongLuong)" -ne "") { $ws.Cells.Item($row, 5).Value2 = [int]$item.trongLuong }
                
                $row++
            }
            
            $ws.PrintOut([Type]::Missing, [Type]::Missing, 2, $false, $data.mayIn)
        } finally {
            # ======================================================
            # 🔥 TỰ ĐỘNG DỌN DẸP XÁC EXCEL SAU KHI IN
            # ======================================================
            if ($wb) { $wb.Close($false) }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    `;

    const psFile = path.join(app.getPath('userData'), 'print_job_px.ps1');
    fs.writeFileSync(psFile, '\ufeff' + psScript, 'utf8'); 
    
    exec(`powershell.exe -ExecutionPolicy Bypass -File "${psFile}"`, (error) => {
        if (error) console.log("❌ Lỗi in Excel PX:", error);
        else console.log("✅ Đã in Phiếu Xuất bằng Excel thành công!");
    });
});
// =========================================================================
// HÀM IN BILL CHO MÁY IN KIM (GIẤY LIÊN TỤC)
// =========================================================================
ipcMain.handle('in-bill-excel-nong', async (event, data) => {
    const tempFile = path.join(app.getPath('userData'), 'temp_bill.json');
    const templatePath = path.join(__dirname, 'Bill_Template.xlsm');

    if (!fs.existsSync(templatePath)) {
        console.log("❌ KHÔNG TÌM THẤY FILE MẪU: Bill_Template.xlsm");
        return;
    }

    fs.writeFileSync(tempFile, JSON.stringify(data), 'utf8');

    const psScript = `
        $ErrorActionPreference = 'SilentlyContinue'
        $data = Get-Content '${tempFile}' -Encoding UTF8 -Raw | ConvertFrom-Json
        
        # ======================================================
        # 🔥 ĐÃ NÂNG CẤP: TẠO EXCEL "TÀNG HÌNH" ĐỘC LẬP CHO HÓA ĐƠN
        # ======================================================
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        try {
            $wb = $excel.Workbooks.Open('${templatePath}', [Type]::Missing, $true)
            $ws = $wb.Sheets.Item("INHOADON")

            # ======================================================
            # ✅ ĐIỀN TÊN KẾ TOÁN VÀO EXCEL TEMPLATE
            # Trong file mẫu chỉ cần đặt chữ: {{KE_TOAN}}
            # Sỉ = Thúy Thuong | Lẻ = Bạch Kim
            # ======================================================
            $keToan = "$($data.keToan)"
            if ([string]::IsNullOrWhiteSpace($keToan)) {
                $flowText = ("$($data.flow) $($data.appFlow) $($data.luongXuat)").ToUpper()
                if ($flowText.Contains("LE") -or $flowText.Contains("LẺ")) {
                    $keToan = "Bạch Kim"
                } else {
                    $keToan = "Thúy Thuong"
                }
            }
            foreach ($sh in @($wb.Worksheets)) {
                try {
                    $usedRange = $sh.UsedRange
                    $usedRange.Replace("{{KE_TOAN}}", $keToan, 2, 1, $false, $false, $false, $false) | Out-Null
                } catch {}
            }
            
            # 1. ĐIỀN HEADER (Chỉ cần điền 1 lần vì tờ nào header cũng giống nhau)
            $ws.Cells.Item(2, 5).Value = "'$($data.ngay)"
            $ws.Cells.Item(2, 6).Value = "'$($data.thang)"
            $ws.Cells.Item(2, 7).Value = "'$($data.nam)"
            if ("$($data.timeStr)" -ne "") { $ws.Cells.Item(3, 5).Value = "'$($data.timeStr)" }
            
            $ws.Cells.Item(5, 3).Value = "$($data.tenKH)"
            $ws.Cells.Item(5, 6).Value = "$($data.maKH)"
            if ("$($data.mst)" -ne "") { $ws.Cells.Item(7, 7).Value = "$($data.mst)" }
            if ("$($data.soXe)" -ne "") { $ws.Cells.Item(6, 6).Value = "$($data.soXe)" }
            
            # ========================================================
            # 2. VÒNG LẶP IN TỪNG TRANG (TÁCH BILL / PHÂN TRANG)
            # ========================================================
            foreach ($page in $data.pages) {
    
                # CHỈ XÓA ĐÚNG 5 DÒNG HÀNG HÓA, KHÔNG ĐỤNG CHẠM DÒNG 15
                $ws.Range("B10:G14").ClearContents()
                
                $row = 10
                foreach ($item in $page.items) {
                    $ws.Cells.Item($row, 2).Value = "$($item.maHang)"
                    $ws.Cells.Item($row, 3).Value = "$($item.tenHang)"
                    
                    if ($item.soThung -ne $null -and $item.soThung -ne "") { $ws.Cells.Item($row, 4).Value2 = [double]$item.soThung }
                    if ($item.trongLuong -ne $null -and $item.trongLuong -ne "") { $ws.Cells.Item($row, 5).Value2 = [double]$item.trongLuong }
                    
                    # Nếu có Ẩn giá, core.js sẽ gửi "donGia" là rỗng ("")
                    if ($item.donGia -ne $null -and $item.donGia -ne "") { $ws.Cells.Item($row, 6).Value2 = [double]$item.donGia }
                    if ($item.thanhTien -ne $null -and $item.thanhTien -ne "") { $ws.Cells.Item($row, 7).Value2 = [double]$item.thanhTien }
                    
                    $row++
                }
                
                # IN PHỤ THU
                if ($page.phuThu -ne $null) {
                    $ws.Cells.Item($row, 3).Value = "$($page.phuThu.ten)"
                    if ($page.phuThu.gia -ne $null -and $page.phuThu.gia -ne "") {
                        $ws.Cells.Item($row, 7).Value2 = [double]$page.phuThu.gia
                    }
                }
                
                # ========================================================
                # 🔥 TẨY TRẮNG TỔNG TIỀN NẾU CHỌN "IN ẨN GIÁ"
                # ========================================================
                if ($data.isAnGia -eq $true) {
                    $ws.Cells.Item(15, 7).Value = "" # Tẩy ô Tổng Tiền Bằng Số
                    $ws.Cells.Item(16, 3).Value = "" # Tẩy ô Tổng Tiền Bằng Chữ
                }
                
                # RA LỆNH IN 1 TỜ
                $ws.PrintOut([Type]::Missing, [Type]::Missing, 1, $false, $data.mayIn)
            }
            # ========================================================
            
        } finally {
            # ======================================================
            # 🔥 TỰ ĐỘNG DỌN DẸP XÁC EXCEL SAU KHI IN
            # ======================================================
            if ($wb) { $wb.Close($false) }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    `;

    const psFile = path.join(app.getPath('userData'), 'print_job_bill.ps1');
    fs.writeFileSync(psFile, '\ufeff' + psScript, 'utf8'); 
    
    exec(`powershell.exe -ExecutionPolicy Bypass -File "${psFile}"`, (error) => {
        if (error) console.log("❌ Lỗi in Bill:", error);
        else console.log("✅ Đã in Bill bằng Excel thành công!");
    });
});