import { Product, WeightMap, InventoryItem, HistoryItem, ContItem } from "../types";
import { getAccessTokenSync } from "./firebaseAuth";

export interface GoogleSheetsConfig {
  spreadsheetId: string;
  connectionType: "local" | "appscript" | "oauth";
  webAppUrl: string;
  accessToken: string;
  clientId: string;
}

const DEFAULT_SPREADSHEET_ID = "1QKueqrlHhA0xhW1ZnkE4QrpKLUVD6eDTRNYxUMBtbhc";

export const getSavedConfig = (): GoogleSheetsConfig => {
  const defaultUrl = "https://script.google.com/macros/s/AKfycbxuHHa2RNny5HCPyswWwRAZDFs8F6KiL0nWIE2oDs9X-dEoOLoWgiI7fCopvD7n3gtCYw/exec";
  const localConfig = localStorage.getItem("cteg_gs_config");
  if (localConfig) {
    try {
      const parsed = JSON.parse(localConfig);
      return {
        spreadsheetId: DEFAULT_SPREADSHEET_ID,
        accessToken: "",
        clientId: "",
        ...parsed,
        connectionType: parsed.connectionType === "local" ? "appscript" : (parsed.connectionType || "appscript"),
        webAppUrl: parsed.webAppUrl || defaultUrl,
      };
    } catch (e) {
      // fallback
    }
  }
  return {
    spreadsheetId: DEFAULT_SPREADSHEET_ID,
    connectionType: "appscript",
    webAppUrl: defaultUrl,
    accessToken: "",
    clientId: "",
  };
};

export const saveConfig = (config: GoogleSheetsConfig) => {
  localStorage.setItem("cteg_gs_config", JSON.stringify(config));
};

// ===================================
// CLIENT-SIDE LOCAL FALLBACK ENGINE
// ===================================
// When google sheet is not linked, this mirrors the exact Apps Script DB V2.6 logic offline!
// 100% compliant with locking, separate lines with: Ke + Ten + CB + GhiChu.

export const localProcessMultiAdjust = (ke: string, items: any[]) => {
  const localInventory = localStorage.getItem("cteg_inventory");
  const localHistory = localStorage.getItem("cteg_history");
  const localWeightMap = localStorage.getItem("cteg_weight_map");

  const inventory: InventoryItem[] = localInventory ? JSON.parse(localInventory) : [];
  const history: HistoryItem[] = localHistory ? JSON.parse(localHistory) : [];
  const weightMap: WeightMap = localWeightMap ? JSON.parse(localWeightMap) : {};

  // CHECK CHẶN TỒN KHO ÂM (VÒNG LẬP 1)
  for (const item of items) {
    const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
    if (!isNhap) {
      const ten = item.tenHang.trim();
      const b = parseFloat(item.bao) || 0;
      const cb = parseFloat(item.slBao) || 0;
      const l = parseFloat(item.le) || 0;
      const note = (item.ghiChu || "").trim();
      const xuatTotalCon = b * cb + l;

      // Find matching item by Ke + Ten + CB + GhiChu (Trim & UpperCase)
      const matched = inventory.find(
        (i) =>
          i.ke.toUpperCase() === ke.toUpperCase() &&
          i.ten.toUpperCase() === ten.toUpperCase() &&
          i.slBao === cb &&
          (i.ghiChu || "").trim().toUpperCase() === note.toUpperCase()
      );

      if (!matched) {
        return {
          success: false,
          msg: `LỖI: Không tìm thấy hàng "${ten}" quy cách ${cb} con/bao ${
            note ? `(Ghi chú: ${note})` : ""
          } tại kệ ${ke.toUpperCase()} để xuất!`,
        };
      }

      if (matched.tongCon < xuatTotalCon) {
        return {
          success: false,
          msg:
            `LỖI: Kệ ${ke.toUpperCase()} không đủ hàng cho mã "${ten}" quy cách ${cb}! \n` +
            `Yêu cầu xuất: ${xuatTotalCon} con (Tương đương ${b} bao x ${cb} + ${l} lẻ). \n` +
            `Tồn kho thực tế hiện tại: ${matched.tongCon} con (Tương đương ${matched.bao} bao x ${matched.slBao} + ${matched.le} lẻ).`,
        };
      }
    }
  }

  // EXECUTE ADJUSTMENTS (VÒNG LẶP 2)
  const now = new Date();
  const timeStr = now.toLocaleString("vi-VN", { hour12: true });

  items.forEach((item) => {
    const ten = item.tenHang.trim();
    const b = parseFloat(item.bao) || 0;
    const cb = parseFloat(item.slBao) || 0;
    const l = parseFloat(item.le) || 0;
    const note = (item.ghiChu || "").trim();
    const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";

    const totalCon = b * cb + l;
    const tyTrong = weightMap[ten.toUpperCase()] || 0;
    const weight = Math.round(((totalCon * tyTrong) / 1000) * 100) / 100;

    // 1. History Append
    history.unshift({
      id: `hist-${now.getTime()}-${Math.random()}`,
      ngay: timeStr,
      loai: isNhap ? "NHẬP" : "XUẤT",
      ke: ke.toUpperCase(),
      tenHang: ten,
      bao: b,
      slBao: cb,
      le: l,
      tong: totalCon,
      kg: weight,
      ghiChu: item.ghiChu || "",
    });

    // 2. Update Inventory
    let found = false;
    for (let i = 0; i < inventory.length; i++) {
      const inv = inventory[i];
      if (
        inv.ke.toUpperCase() === ke.toUpperCase() &&
        inv.ten.toUpperCase() === ten.toUpperCase() &&
        inv.slBao === cb &&
        (inv.ghiChu || "").trim().toUpperCase() === note.toUpperCase()
      ) {
        const newTotal = isNhap ? inv.tongCon + totalCon : inv.tongCon - totalCon;

        inventory[i] = {
          ...inv,
          bao: Math.floor(newTotal / cb),
          le: newTotal % cb,
          tongCon: newTotal,
          khoiLuong: Math.round(((newTotal * tyTrong) / 1000) * 100) / 100,
        };
        found = true;
        break;
      }
    }

    if (!found && isNhap) {
      inventory.push({
        id: `inv-${now.getTime()}-${Math.random()}`,
        ke: ke.toUpperCase(),
        ten: ten,
        bao: b,
        slBao: cb,
        le: l,
        tongCon: totalCon,
        khoiLuong: weight,
        ghiChu: item.ghiChu || "",
      });
    }
  });

  // Filter out empty rows (tongCon <= 0)
  const finalInventory = inventory.filter((r) => r.tongCon > 0);

  localStorage.setItem("cteg_inventory", JSON.stringify(finalInventory));
  localStorage.setItem("cteg_history", JSON.stringify(history));

  return { success: true };
};

export const localDeleteSelectedHistory = (ids: string[], password?: string) => {
  const matKhauHopLe = "Tanlh";
  if (password !== matKhauHopLe) return "Sai mật khẩu!";

  const localInventory = localStorage.getItem("cteg_inventory");
  const localHistory = localStorage.getItem("cteg_history");
  const localWeightMap = localStorage.getItem("cteg_weight_map");

  const inventory: InventoryItem[] = localInventory ? JSON.parse(localInventory) : [];
  const history: HistoryItem[] = localHistory ? JSON.parse(localHistory) : [];
  const weightMap: WeightMap = localWeightMap ? JSON.parse(localWeightMap) : {};

  const historyToDelete = history.filter((h) => ids.includes(h.id));
  const remainingHistory = history.filter((h) => !ids.includes(h.id));

  historyToDelete.forEach((deletedItem) => {
    const isNhap = deletedItem.loai.toUpperCase() === "NHẬP" || deletedItem.loai.toUpperCase() === "NHAP";
    const ke = deletedItem.ke.toUpperCase();
    const tenHang = deletedItem.tenHang.toUpperCase();
    const slXoa = deletedItem.tong;
    const cb = deletedItem.slBao || 1;
    const note = (deletedItem.ghiChu || "").trim().toUpperCase();

    // Revert inventory
    let found = false;
    for (let j = 0; j < inventory.length; j++) {
      const inv = inventory[j];
      if (
        inv.ke.toUpperCase() === ke &&
        inv.ten.toUpperCase() === tenHang &&
        inv.slBao === cb &&
        (inv.ghiChu || "").trim().toUpperCase() === note
      ) {
        let res = isNhap ? inv.tongCon - slXoa : inv.tongCon + slXoa;
        res = res < 0 ? 0 : res;

        const tyTrong = weightMap[tenHang] || 0;
        inventory[j] = {
          ...inv,
          bao: Math.floor(res / cb),
          le: res % cb,
          tongCon: res,
          khoiLuong: Math.round(((res * tyTrong) / 1000) * 100) / 100,
        };
        found = true;
        break;
      }
    }

    // If item was deleted completely but we are putting it back from an cancelled export
    if (!found && !isNhap) {
      const tyTrong = weightMap[tenHang] || 0;
      inventory.push({
        id: `inv-${Date.now()}-${Math.random()}`,
        ke: ke,
        ten: deletedItem.tenHang,
        bao: Math.floor(slXoa / cb),
        slBao: cb,
        le: slXoa % cb,
        tongCon: slXoa,
        khoiLuong: Math.round(((slXoa * tyTrong) / 1000) * 100) / 100,
        ghiChu: deletedItem.ghiChu,
      });
    }
  });

  const finalInventory = inventory.filter((r) => r.tongCon > 0);

  localStorage.setItem("cteg_inventory", JSON.stringify(finalInventory));
  localStorage.setItem("cteg_history", JSON.stringify(remainingHistory));

  return "Đã xóa lịch sử, cập nhật tồn kho và xóa dòng trống!";
};


// ===================================
// GOOGLE SHEETS live SERVICE API
// ===================================

export class GoogleSheetsService_Old {
  private config: GoogleSheetsConfig;

  constructor() {
    this.config = getSavedConfig();
  }

  isConfigured(): boolean {
    if (this.config.connectionType === "appscript" && this.config.webAppUrl) return true;
    if (this.config.connectionType === "oauth" && this.config.accessToken) return true;
    return false;
  }

  getConnectionTypeLabel(): string {
    if (this.config.connectionType === "appscript") return "Apps Script Web App";
    if (this.config.connectionType === "oauth") return "Google Sheets OAuth";
    return "Lưu trữ Cục bộ Offline";
  }

  async testConnection(config: GoogleSheetsConfig): Promise<{ success: boolean; message: string }> {
    const backup = this.config;
    try {
      this.config = config;
      if (config.connectionType === "local") {
        return { success: true, message: "Chế độ Local Offline hoạt động tốt!" };
      }

      if (config.connectionType === "appscript") {
        if (!config.webAppUrl) {
          return { success: false, message: "Vui lòng nhập đường dẫn Web App trong cấu hình!" };
        }
        try {
          const url = new URL(config.webAppUrl);
          url.searchParams.append("action", "test");
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 8000); // 8s timeout
          
          const res = await fetch(url.toString(), { 
            method: "GET", 
            mode: "cors",
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          
          if (!res.ok) {
            return { success: false, message: `Lỗi phản hồi HTTP: ${res.status}` };
          }
          return { success: true, message: "Kết nối Apps Script Web App thành công!" };
        } catch (e: any) {
          if (e.name === "AbortError") {
            return { success: false, message: "Yêu cầu kết nối quá hạn (Timeout - 8 giây). Vui lòng thử lại!" };
          }
          return { success: false, message: `Không thể kết nối đến Web App: ${e.message || e}` };
        }
      }

      if (config.connectionType === "oauth") {
        if (!config.spreadsheetId) {
          return { success: false, message: "Vui lòng nhập Google Sheets Spreadsheet ID!" };
        }
        if (!config.accessToken) {
          return { success: false, message: "Vui lòng nhập Access Token!" };
        }
        try {
          const res = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${config.spreadsheetId}`,
            {
              headers: {
                Authorization: `Bearer ${config.accessToken}`,
                "Content-Type": "application/json",
              }
            }
          );
          if (!res.ok) {
            const errJson = await res.json().catch(() => ({}));
            return { 
              success: false, 
              message: `Lỗi Google API (${res.status}): ${errJson.error?.message || "Token hoặc Spreadsheet ID chưa đúng/hết hạn!"}` 
            };
          }
          return { success: true, message: "Xác thực Google Sheets API thành công!" };
        } catch (e: any) {
          return { success: false, message: `Lỗi mạng kết nối Google Server: ${e.message || e}` };
        }
      }

      return { success: false, message: "Không tìm thấy phương thức phù hợp." };
    } finally {
      this.config = backup;
    }
  }

  // Helper for requests
  private async request(action: string, method: "GET" | "POST", data?: any): Promise<any> {
    const spreadsheetId = this.config.spreadsheetId || DEFAULT_SPREADSHEET_ID;

    // APPS SCRIPT WEB APP GATEWAY
    if (this.config.connectionType === "appscript" && this.config.webAppUrl) {
      const url = new URL(this.config.webAppUrl);
      let res: Response;
      try {
        if (method === "GET") {
          url.searchParams.append("action", action);
          if (data) {
            Object.keys(data).forEach((key) => url.searchParams.append(key, String(data[key])));
          }
          res = await fetch(url.toString(), { method: "GET", mode: "cors" });
        } else {
          // Appscript POST requires CORS handling or sending as text/plain to avoid preflight issues
          res = await fetch(this.config.webAppUrl, {
            method: "POST",
            mode: "cors",
            headers: {
              "Content-Type": "text/plain;charset=utf-8", // bypass standard CORS preflight
            },
            body: JSON.stringify({ action, ...data }),
          });
        }
      } catch (e: any) {
        throw new Error(`Lỗi kết nối đến Google Sheets Apps Script: ${e.message || e}. Xin hãy kiểm tra mạng hoặc đường dẫn Web App.`);
      }

      if (!res.ok) {
        throw new Error(`Google Apps Script API trả về mã lỗi HTTP: ${res.status}`);
      }

      let json: any;
      const resText = await res.text();
      try {
        json = JSON.parse(resText);
      } catch (err) {
        if (resText.includes("<!DOCTYPE html") || resText.includes("<html") || resText.includes("AccessDenied") || resText.includes("login") || resText.includes("Sign in")) {
          throw new Error("Ứng dụng phát hiện trang đăng nhập hoặc trang HTML từ Google. Có thể bạn cấu hình quyền truy cập Apps Script chưa đúng. Hãy kiểm tra và chắc chắn đã chọn 'Anyone' (Bất kỳ ai) có quyền truy cập khi triển khai (Deploy) Web App!");
        }
        throw new Error(`Không thể giải mã dữ liệu trả về từ Apps Script. Định dạng không phải JSON hợp lệ. Chi tiết phản hồi: ${resText.slice(0, 150)}...`);
      }

      if (json && typeof json === "object" && json.success === false) {
        throw new Error(json.msg || json.message || "Apps Script báo lỗi không xác định.");
      }

      return json;
    }

    // DIRECT GOOGLE SHEETS REST API
    if (this.config.connectionType === "oauth" && this.config.accessToken) {
      const headers = {
        Authorization: `Bearer ${this.config.accessToken}`,
        "Content-Type": "application/json",
      };

      const handleFetchWithCheck = async (url: string, init?: RequestInit) => {
        try {
          const r = await fetch(url, init);
          if (!r.ok) {
            const errJson = await r.json().catch(() => ({}));
            const msg = errJson.error?.message || `Lỗi HTTP ${r.status}`;
            throw new Error(`Lỗi Google API (${r.status}): ${msg}`);
          }
          return await r.json();
        } catch (e: any) {
          if (e.message && e.message.includes("Lỗi Google API")) {
            throw e;
          }
          throw new Error(`Lỗi kết nối REST API: ${e.message || e}`);
        }
      };

      if (action === "getProductList" || action === "getTrong_LuongData") {
        const range = "'Mã hàng'!B2:D";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values.map((row: any) => [
          row[0] || "", // Ten
          "",
          row[2] || "0", // Ty trong
        ]);
      }

      if (action === "getRealTimeInventory") {
        const range = "'Tổng hợp'!A2:I";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values.map((row: any) => [
          row[0] || "", // ke
          row[1] || "", // empty
          row[2] || "", // ten
          row[3] || "0", // bao
          row[4] || "0", // slBao
          row[5] || "0", // le
          row[6] || "0", // tongCon
          row[7] || "0", // khoiLuong
          row[8] || "", // ghiChu
        ]);
      }

      if (action === "getHistory") {
        const range = "'Lịch sử'!A2:K";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values;
      }

      if (action === "getContList") {
        const range = "'Cont'!A2:L";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        let filtered = json.values;
        if (data && data.keyword && data.keyword !== "ALL") {
          filtered = json.values.filter((r: any) =>
            String(r[7] || "").toUpperCase().includes(data.keyword.toUpperCase())
          );
        }
        return filtered;
      }

      if (action === "processMultiAdjust") {
        // Direct implementation inside OAuth REST client!
        const productsRaw = await this.request("getTrong_LuongData", "GET");
        const ttMap: Record<string, number> = {};
        productsRaw.forEach((row: any) => {
          const ten = String(row[0]).trim().toUpperCase();
          const tyTrong = parseFloat(String(row[2]).replace(/,/g, ".")) || 0;
          ttMap[ten] = tyTrong;
        });

        const currentThRaw = await this.request("getRealTimeInventory", "GET");
        const thData = currentThRaw.map((row: any) => [
          row[0] || "", // ke
          row[1] || "", // empty
          row[2] || "", // ten
          row[3] || "0", // bao
          row[4] || "0", // slBao
          row[5] || "0", // le
          row[6] || "0", // tongCon
          row[7] || "0", // khoiLuong
          row[8] || "", // ghiChu
        ]);

        const items: any[] = data.items || [];
        const ke: string = data.ke || "";

        // Verification Loop 1
        for (const item of items) {
          const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
          if (!isNhap) {
            const ten = item.tenHang.trim();
            const b = parseFloat(item.bao) || 0;
            const cb = parseFloat(item.slBao) || 0;
            const l = parseFloat(item.le) || 0;
            const note = (item.ghiChu || "").trim();
            const xuatTotalCon = b * cb + l;

            let foundIdx = -1;
            for (let i = 0; i < thData.length; i++) {
              if (
                String(thData[i][0]).toUpperCase() === ke.toUpperCase() &&
                String(thData[i][2]).toUpperCase() === ten.toUpperCase() &&
                parseFloat(thData[i][4]) === cb &&
                String(thData[i][8] || "").trim().toUpperCase() === note.toUpperCase()
              ) {
                foundIdx = i;
                break;
              }
            }

            if (foundIdx === -1) {
              return {
                success: false,
                msg: `LỖI: Không tìm thấy hàng "${ten}" quy cách ${cb} con/bao ${
                  note ? `(Ghi chú: ${note})` : ""
                } tại kệ ${ke.toUpperCase()} để xuất!`,
              };
            }

            const currentB = parseFloat(thData[foundIdx][3]) || 0;
            const currentCB = parseFloat(thData[foundIdx][4]) || cb;
            const currentL = parseFloat(thData[foundIdx][5]) || 0;
            const currentTotalCon = currentB * currentCB + currentL;

            if (currentTotalCon < xuatTotalCon) {
              return {
                success: false,
                msg:
                  `LỖI: Kệ ${ke.toUpperCase()} không đủ hàng cho mã "${ten}" quy cách ${cb}! \n` +
                  `Yêu cầu xuất: ${xuatTotalCon} con (Tương đương ${b} bao x ${cb} + ${l} lẻ). \n` +
                  `Tồn kho thực tế hiện tại: ${currentTotalCon} con (Tương đương ${currentB} bao x ${currentCB} + ${currentL} lẻ).`,
              };
            }
          }
        }

        // Loop 2: Updates
        const now = new Date();
        const historyRowsToAppend: any[][] = [];

        items.forEach((item) => {
          const ten = item.tenHang.trim();
          const b = parseFloat(item.bao) || 0;
          const cb = parseFloat(item.slBao) || 0;
          const l = parseFloat(item.le) || 0;
          const note = (item.ghiChu || "").trim();
          const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";

          const totalCon = b * cb + l;
          const tyTrong = ttMap[ten.toUpperCase()] || 0;
          const weight = Math.round(((totalCon * tyTrong) / 1000) * 100) / 100;

          // Push History row
          historyRowsToAppend.push([
            now.toISOString(),
            item.type,
            ten,
            ke.toUpperCase(),
            b,
            cb,
            l,
            weight + " Kg",
            item.ghiChu || "",
            "ID_" + now.getTime() + "_" + Math.floor(Math.random() * 1000),
            totalCon,
          ]);

          let found = false;
          for (let i = 0; i < thData.length; i++) {
            if (
              String(thData[i][0]).toUpperCase() === ke.toUpperCase() &&
              String(thData[i][2]).toUpperCase() === ten.toUpperCase() &&
              parseFloat(thData[i][4]) === cb &&
              String(thData[i][8] || "").trim().toUpperCase() === note.toUpperCase()
            ) {
              const currentB = parseFloat(thData[i][3]) || 0;
              const currentCB = parseFloat(thData[i][4]) || cb;
              const currentL = parseFloat(thData[i][5]) || 0;

              const currentTotal = currentB * currentCB + currentL;
              const newTotal = isNhap ? currentTotal + totalCon : currentTotal - totalCon;

              thData[i][3] = Math.floor(newTotal / currentCB);
              thData[i][5] = newTotal % currentCB;
              thData[i][6] = newTotal;
              thData[i][7] = Math.round(((newTotal * tyTrong) / 1000) * 100) / 100;
              found = true;
              break;
            }
          }

          if (!found && isNhap) {
            thData.push([ke.toUpperCase(), "", ten, b, cb, l, totalCon, weight, item.ghiChu || ""]);
          }
        });

        // Filter and clear 'Tong hop' and rewrite
        const finalThData = thData.filter((r: any) => {
          const total = parseFloat(r[6]) || 0;
          return total > 0;
        });

        // Clear range 'Tổng hợp'!A2:I1000 then put
        await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Tổng hợp'!A2:I999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (finalThData.length > 0) {
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Tổng hợp'!A2:I" + (finalThData.length + 1)
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: finalThData }),
            }
          );
        }

        // Append to 'Lịch sử'!A2:K
        if (historyRowsToAppend.length > 0) {
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Lịch sử'!A2:K"
            )}:append?valueInputOption=USER_ENTERED`,
            {
              method: "POST",
              headers,
              body: JSON.stringify({ values: historyRowsToAppend }),
            }
          );
        }

        return { success: true };
      }

      if (action === "deleteSelectedHistory") {
        if (data.password !== "Tanlh") {
          return { success: false, msg: "Sai mật khẩu!" };
        }

        // Direct implementation of deleteSelectedHistory in REST client
        const currentThRaw = await this.request("getRealTimeInventory", "GET");
        const thData = currentThRaw.map((row: any) => [
          row[0] || "", row[1] || "", row[2] || "", row[3] || "0", row[4] || "0", row[5] || "0", row[6] || "0", row[7] || "0", row[8] || ""
        ]);

        const historyRaw = await this.request("getHistory", "GET");
        const idsToDelete: string[] = data.ids || [];

        // Match indices of rows to delete from sheet 'Lịch sử'
        // For REST client, we rewrite remaining histories
        const remainingHistoryRows = [
          ["Ngày giờ", "Loại", "Tên sản phẩm", "Kệ", "Số bao", "Quy cách", "Số lẻ", "Khối lượng", "Ghi chú", "ID", "Tổng số con"]
        ];

        for (let i = 0; i < historyRaw.length; i++) {
          const rowItem = historyRaw[i];
          const histId = String(rowItem[9]); // col J is ID

          if (idsToDelete.includes(histId)) {
            // Revert inventory quantities
            const type = String(rowItem[1]).toUpperCase().trim();
            const ke = String(rowItem[3]).trim().toUpperCase();
            const tenHang = String(rowItem[2]).trim().toUpperCase();
            const slXoa = parseFloat(rowItem[10]) || 0;

            for (let j = 0; j < thData.length; j++) {
              if (
                String(thData[j][0]).trim().toUpperCase() === ke &&
                String(thData[j][2]).trim().toUpperCase() === tenHang
              ) {
                const cb = parseFloat(thData[j][4]) || 1;
                const cur = parseFloat(thData[j][6]) || 0;

                let res = (type === "NHẬP" || type === "NHAP") ? (cur - slXoa) : (cur + slXoa);
                res = res < 0 ? 0 : res;

                thData[j][3] = Math.floor(res / cb);
                thData[j][5] = res % cb;
                thData[j][6] = res;
              }
            }
          } else {
            remainingHistoryRows.push(rowItem);
          }
        }

        // Filter and write Tổng hợp
        const filteredThData = thData.filter((r: any) => parseFloat(r[6]) > 0);

        // Delete & write Lịch sử
        await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Lịch sử'!A2:K9999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (remainingHistoryRows.length > 1) {
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Lịch sử'!A1:K" + remainingHistoryRows.length
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: remainingHistoryRows }),
            }
          );
        }

        // Clear & write Tổng hợp
        await fetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Tổng hợp'!A2:I999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (filteredThData.length > 0) {
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Tổng hợp'!A2:I" + (filteredThData.length + 1)
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: filteredThData }),
            }
          );
        }

        return { success: true, msg: "Đã xóa lịch sử, cập nhật tồn kho và xóa dòng trống!" };
      }

      if (action === "saveContData") {
        const rows: any[][] = data.data || [];
        const isReplace: boolean = data.isReplace || false;

        if (isReplace) {
          const finishedContName = rows[0][7]; // ContName
          const currentContRaw = await this.request("getContList", "GET", { keyword: "ALL" });
          const unchangedContRows = currentContRaw.filter(
            (r: any) => String(r[7]).toUpperCase() !== finishedContName.toUpperCase()
          );

          // Full writeback
          const header = ["STT", "ID", "Phân Cấp", "Tiêu chuẩn", "Kích Cỡ", "Bề Mặt", "Số bàn giao (NO)", "Tên Container", "Bao", "CB", "Lẻ", "Ghi chú"];
          const finalRows = [header, ...unchangedContRows, ...rows.map((row: any) => [...row.slice(0, 8), 0, 0, 0, ""])];

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A1:L9999"
            )}:clear`,
            { method: "POST", headers }
          );

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A1:L" + finalRows.length
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: finalRows }),
            }
          );
        } else {
          // simple append
          const formatToAppend = rows.map((row: any) => [...row.slice(0, 8), 0, 0, 0, ""]);
          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A2:L"
            )}:append?valueInputOption=USER_ENTERED`,
            {
              method: "POST",
              headers,
              body: JSON.stringify({ values: formatToAppend }),
            }
          );
        }
        return { success: true };
      }

      if (action === "updateContQuantity") {
        const { grade, standard, size, finish, no, contName, b, cb, l, n, newG, newS, newSz, newF, newNo } = data;
        const currentContRaw = await this.request("getContList", "GET", { keyword: "ALL" });

        let foundRowIdx = -1;
        for (let i = 0; i < currentContRaw.length; i++) {
          const row = currentContRaw[i];
          if (
            String(row[2]) === String(grade) &&
            String(row[3]) === String(standard) &&
            String(row[4]) === String(size) &&
            String(row[5]) === String(finish) &&
            String(row[6]) === String(no) &&
            String(row[7]) === String(contName)
          ) {
            foundRowIdx = i;
            break;
          }
        }

        if (foundRowIdx !== -1) {
          // Column indexes in sheets start from 1. Index i corresponds to row (i + 2)
          // Grade is column 3, Standard 4, Size 5, Finish 6, NO 7, ContName 8, Bao 9, CB 10, Le 11, GhiChu 12
          const targetRowNumber = foundRowIdx + 2;
          const targetRange = `'Cont'!C${targetRowNumber}:L${targetRowNumber}`;

          await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              targetRange
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({
                values: [[newG, newS, newSz, newF, newNo, contName, b, cb, l, n]],
              }),
            }
          );
          return { success: true };
        }
        return { success: false, msg: "Không tìm thấy dữ liệu khớp với thông tin cũ." };
      }
    }

    throw new Error("Không có driver Google Sheet cấu hình phù hợp!");
  }
}

function validateArrayResponse(raw: any, funcName: string): any[] {
  if (!raw) {
    throw new Error(`[Lỗi Google Sheets] Phương thức [${funcName}] nhận về phản hồi rỗng (null/undefined).`);
  }
  if (!Array.isArray(raw)) {
    let detail = "";
    if (typeof raw === "object") {
      detail = raw.msg || raw.message || JSON.stringify(raw);
    } else {
      detail = String(raw);
    }
    if (detail.includes("<!DOCTYPE html") || detail.includes("<html") || detail.includes("login") || detail.includes("Sign in")) {
      detail = "Yêu cầu đăng nhập hoặc cấp quyền truy cập từ Google. Vui lòng kiểm tra lại cấu hình phân quyền (Ai có quyền truy cập: Bất kỳ ai / Anyone) của Web App!";
    }
    throw new Error(`[Lỗi Google Sheets - ${funcName}] Dữ liệu nhận được không phải là một Danh sách Mảng hợp lệ. Chi tiết lỗi từ Google: ${detail}`);
  }
  return raw;
}

export class GoogleSheetsService {
  private config: GoogleSheetsConfig;

  constructor() {
    this.config = getSavedConfig();
  }

  isConfigured(): boolean {
    if (this.config.connectionType === "appscript" && this.config.webAppUrl) return true;
    if (this.config.connectionType === "oauth" && (this.config.accessToken || getAccessTokenSync())) return true;
    return false;
  }

  getConnectionTypeLabel(): string {
    if (this.config.connectionType === "appscript") return "Apps Script Web App";
    if (this.config.connectionType === "oauth") return "Google Sheets OAuth";
    return "Lưu trữ Cục bộ Offline";
  }

  async testConnection(config: GoogleSheetsConfig): Promise<{ success: boolean; message: string }> {
    const backup = this.config;
    try {
      this.config = config;
      if (config.connectionType === "local") {
        return { success: true, message: "Chế độ Local Offline hoạt động tốt!" };
      }

      if (config.connectionType === "appscript") {
        if (!config.webAppUrl) {
          return { success: false, message: "Vui lòng nhập đường dẫn Web App trong cấu hình!" };
        }
        try {
          const url = new URL(config.webAppUrl);
          url.searchParams.append("action", "test");
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 8000); // 8s timeout
          
          const res = await fetch(url.toString(), { 
            method: "GET", 
            mode: "cors",
            signal: controller.signal
          });
          clearTimeout(timeoutId);
          
          if (!res.ok) {
            return { success: false, message: `Lỗi phản hồi HTTP: ${res.status}` };
          }
          
          const resText = await res.text();
          let jsonResult: any;
          try {
            jsonResult = JSON.parse(resText);
          } catch {
            if (resText.includes("<!DOCTYPE html") || resText.includes("<html") || resText.includes("login") || resText.includes("Sign in")) {
              return { success: false, message: "Apps Script trả về giao diện đăng nhập Google (HTML). Bạn cần cấu hình loại quyền truy cập khi triển khai Web App là 'Anyone' (Bất kỳ ai / Mọi người)." };
            }
            return { success: false, message: `Phản hồi không phải JSON: ${resText.slice(0, 100)}...` };
          }
          
          if (jsonResult && jsonResult.success === false) {
            return { success: false, message: jsonResult.msg || jsonResult.message || "Lỗi cấu hình script!" };
          }
          
          return { success: true, message: "Kết nối Apps Script Web App thành công!" };
        } catch (e: any) {
          if (e.name === "AbortError") {
            return { success: false, message: "Yêu cầu kết nối quá hạn (Timeout - 8 giây). Vui lòng thử lại!" };
          }
          return { success: false, message: `Không thể kết nối đến Web App: ${e.message || e}` };
        }
      }

      if (config.connectionType === "oauth") {
        if (!config.spreadsheetId) {
          return { success: false, message: "Vui lòng nhập Google Sheets Spreadsheet ID!" };
        }
        const tokenToTest = config.accessToken || getAccessTokenSync();
        if (!tokenToTest) {
          return { success: false, message: "Vui lòng Đăng nhập với Google ở ngoài hoặc nhập Access Token!" };
        }
        try {
          const res = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${config.spreadsheetId}`,
            {
              headers: {
                Authorization: `Bearer ${tokenToTest}`,
                "Content-Type": "application/json",
              }
            }
          );
          if (!res.ok) {
            const errJson = await res.json().catch(() => ({}));
            return { 
              success: false, 
              message: `Lỗi Google API (${res.status}): ${errJson.error?.message || "Token hoặc Spreadsheet ID chưa đúng/hết hạn!"}` 
            };
          }
          return { success: true, message: "Xác thực Google Sheets API thành công!" };
        } catch (e: any) {
          return { success: false, message: `Lỗi mạng kết nối Google Server: ${e.message || e}` };
        }
      }

      return { success: false, message: "Không tìm thấy phương thức phù hợp." };
    } finally {
      this.config = backup;
    }
  }

  // Helper for requests
  private async request(action: string, method: "GET" | "POST", data?: any): Promise<any> {
    const spreadsheetId = this.config.spreadsheetId || DEFAULT_SPREADSHEET_ID;
    
    // APPS SCRIPT WEB APP GATEWAY
    if (this.config.connectionType === "appscript" && this.config.webAppUrl) {
      const url = new URL(this.config.webAppUrl);
      let res: Response;
      try {
        if (method === "GET") {
          url.searchParams.append("action", action);
          if (data) {
            Object.keys(data).forEach((key) => url.searchParams.append(key, String(data[key])));
          }
          res = await fetch(url.toString(), { method: "GET", mode: "cors" });
        } else {
          // Appscript POST requires CORS handling or sending as text/plain to avoid preflight issues
          res = await fetch(this.config.webAppUrl, {
            method: "POST",
            mode: "cors",
            headers: {
              "Content-Type": "text/plain;charset=utf-8", // bypass standard CORS preflight
            },
            body: JSON.stringify({ action, ...data }),
          });
        }
      } catch (e: any) {
        throw new Error(`Lỗi kết nối đến Google Sheets Apps Script: ${e.message || e}. Xin hãy kiểm tra mạng hoặc đường dẫn Web App.`);
      }

      if (!res.ok) {
        throw new Error(`Google Apps Script API trả về mã lỗi HTTP: ${res.status}`);
      }

      let json: any;
      const resText = await res.text();
      try {
        json = JSON.parse(resText);
      } catch (err) {
        if (resText.includes("<!DOCTYPE html") || resText.includes("<html") || resText.includes("AccessDenied") || resText.includes("login") || resText.includes("Sign in")) {
          throw new Error("Ứng dụng phát hiện trang đăng nhập hoặc trang HTML từ Google. Có thể bạn cấu hình quyền truy cập Apps Script chưa đúng. Hãy kiểm tra và chắc chắn đã chọn 'Anyone' (Bất kỳ ai) có quyền truy cập khi triển khai (Deploy) Web App!");
        }
        throw new Error(`Không thể giải mã dữ liệu trả về từ Apps Script. Định dạng không phải JSON hợp lệ. Chi tiết phản hồi: ${resText.slice(0, 150)}...`);
      }

      if (json && typeof json === "object" && json.success === false) {
        throw new Error(json.msg || json.message || "Apps Script báo lỗi không xác định.");
      }

      return json;
    }

    // DIRECT GOOGLE SHEETS REST API
    const activeToken = this.config.accessToken || getAccessTokenSync();
    if (this.config.connectionType === "oauth" && activeToken) {
      const headers = {
        Authorization: `Bearer ${activeToken}`,
        "Content-Type": "application/json",
      };

      const handleFetchWithCheck = async (url: string, init?: RequestInit) => {
        try {
          const r = await fetch(url, init);
          if (!r.ok) {
            const errJson = await r.json().catch(() => ({}));
            const msg = errJson.error?.message || `Lỗi HTTP ${r.status}`;
            throw new Error(`Lỗi Google API (${r.status}): ${msg}`);
          }
          return await r.json();
        } catch (e: any) {
          if (e.message && e.message.includes("Lỗi Google API")) {
            throw e;
          }
          throw new Error(`Lỗi kết nối REST API: ${e.message || e}`);
        }
      };

      if (action === "getProductList" || action === "getTrong_LuongData") {
        const range = "'Mã hàng'!B2:D";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values.map((row: any) => [
          row[0] || "", // Ten
          "",
          row[2] || "0", // Ty trong
        ]);
      }

      if (action === "getRealTimeInventory") {
        const range = "'Tổng hợp'!A2:I";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values.map((row: any) => [
          row[0] || "", // ke
          row[1] || "", // empty
          row[2] || "", // ten
          row[3] || "0", // bao
          row[4] || "0", // slBao
          row[5] || "0", // le
          row[6] || "0", // tongCon
          row[7] || "0", // khoiLuong
          row[8] || "", // ghiChu
        ]);
      }

      if (action === "getHistory") {
        const range = "'Lịch sử'!A2:K";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        return json.values;
      }

      if (action === "getContList") {
        const range = "'Cont'!A2:L";
        const json = await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
          { headers }
        );
        if (!json.values) return [];
        let filtered = json.values;
        if (data && data.keyword && data.keyword !== "ALL") {
          filtered = json.values.filter((r: any) =>
            String(r[7] || "").toUpperCase().includes(data.keyword.toUpperCase())
          );
        }
        return filtered;
      }

      if (action === "processMultiAdjust") {
        // Direct implementation inside OAuth REST client!
        const productsRaw = await this.request("getTrong_LuongData", "GET");
        const ttMap: Record<string, number> = {};
        productsRaw.forEach((row: any) => {
          const ten = String(row[0]).trim().toUpperCase();
          const tyTrong = parseFloat(String(row[2]).replace(/,/g, ".")) || 0;
          ttMap[ten] = tyTrong;
        });

        const currentThRaw = await this.request("getRealTimeInventory", "GET");
        const thData = currentThRaw.map((row: any) => [
          row[0] || "", // ke
          row[1] || "", // empty
          row[2] || "", // ten
          row[3] || "0", // bao
          row[4] || "0", // slBao
          row[5] || "0", // le
          row[6] || "0", // tongCon
          row[7] || "0", // khoiLuong
          row[8] || "", // ghiChu
        ]);

        const items: any[] = data.items || [];
        const ke: string = data.ke || "";

        // Verification Loop 1
        for (const item of items) {
          const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
          if (!isNhap) {
            const ten = item.tenHang.trim();
            const b = parseFloat(item.bao) || 0;
            const cb = parseFloat(item.slBao) || 0;
            const l = parseFloat(item.le) || 0;
            const note = (item.ghiChu || "").trim();
            const xuatTotalCon = b * cb + l;

            let foundIdx = -1;
            for (let i = 0; i < thData.length; i++) {
              if (
                String(thData[i][0]).toUpperCase() === ke.toUpperCase() &&
                String(thData[i][2]).toUpperCase() === ten.toUpperCase() &&
                parseFloat(thData[i][4]) === cb &&
                String(thData[i][8] || "").trim().toUpperCase() === note.toUpperCase()
              ) {
                foundIdx = i;
                break;
              }
            }

            if (foundIdx === -1) {
              return {
                success: false,
                msg: `LỖI: Không tìm thấy hàng "${ten}" quy cách ${cb} con/bao trong kẹ "${ke}" để xuất hàng!`,
              };
            }

            const currentTotal = parseFloat(thData[foundIdx][6]) || 0;
            if (currentTotal < xuatTotalCon) {
              return {
                success: false,
                msg: `LỖI: Kẹ "${ke}" hiện chỉ còn ${currentTotal} con "${ten}" quy cách ${cb}, không đủ để xuất lượng yêu cầu là ${xuatTotalCon} con!`,
              };
            }
          }
        }

        const historyRowsToAppend: any[][] = [];
        const timestamp = new Date().toISOString();

        for (const item of items) {
          const isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
          const ten = item.tenHang.trim();
          const b = parseFloat(item.bao) || 0;
          const cb = parseFloat(item.slBao) || 0;
          const l = parseFloat(item.le) || 0;
          const note = (item.ghiChu || "").trim();
          const targetTotal = b * cb + l;
          const tyTrong = ttMap[ten.toUpperCase()] || 0;
          const kg = Math.round(((targetTotal * tyTrong) / 1000) * 100) / 100;

          const historyId = "hist-" + Date.now() + "-" + Math.random().toString(36).substr(2, 5);
          historyRowsToAppend.push([
            timestamp,
            isNhap ? "NHẬP" : "XUẤT",
            ten,
            ke,
            b,
            cb,
            l,
            kg.toFixed(2) + " Kg",
            note,
            historyId,
            targetTotal,
          ]);

          let foundIdx = -1;
          for (let i = 0; i < thData.length; i++) {
            if (
              String(thData[i][0]).toUpperCase() === ke.toUpperCase() &&
              String(thData[i][2]).toUpperCase() === ten.toUpperCase() &&
              parseFloat(thData[i][4]) === cb &&
              String(thData[i][8] || "").trim().toUpperCase() === note.toUpperCase()
            ) {
              foundIdx = i;
              break;
            }
          }

          if (foundIdx !== -1) {
            const currentB = parseFloat(thData[foundIdx][3]) || 0;
            const currentL = parseFloat(thData[foundIdx][5]) || 0;
            let finalB = currentB;
            let finalL = currentL;

            if (isNhap) {
              finalB += b;
              finalL += l;
            } else {
              finalB -= b;
              finalL -= l;
              if (finalL < 0) {
                finalB -= 1;
                finalL += cb;
              }
            }

            const finalTotal = finalB * cb + finalL;
            const finalKg = Math.round(((finalTotal * tyTrong) / 1000) * 100) / 100;

            thData[foundIdx][3] = finalB;
            thData[foundIdx][5] = finalL;
            thData[foundIdx][6] = finalTotal;
            thData[foundIdx][7] = finalKg;
          } else {
            if (isNhap) {
              const total = b * cb + l;
              const kgSum = Math.round(((total * tyTrong) / 1000) * 100) / 100;
              thData.push([ke, "", ten, b, cb, l, total, kgSum, note]);
            }
          }
        }

        const finalThData = thData.filter((r: any) => {
          const total = parseFloat(r[6]) || 0;
          return total > 0;
        });

        // Clear range 'Tổng hợp'!A2:I1000 then put
        await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Tổng hợp'!A2:I999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (finalThData.length > 0) {
          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Tổng hợp'!A2:I" + (finalThData.length + 1)
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: finalThData }),
            }
          );
        }

        if (historyRowsToAppend.length > 0) {
          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Lịch sử'!A2:K"
            )}:append?valueInputOption=USER_ENTERED`,
            {
              method: "POST",
              headers,
              body: JSON.stringify({ values: historyRowsToAppend }),
            }
          );
        }

        return { success: true };
      }

      if (action === "deleteSelectedHistory") {
        if (data.password !== "Tanlh") {
          return { success: false, msg: "Sai mật khẩu!" };
        }

        // Direct implementation of deleteSelectedHistory in REST client
        const currentThRaw = await this.request("getRealTimeInventory", "GET");
        const thData = currentThRaw.map((row: any) => [
          row[0] || "", row[1] || "", row[2] || "", row[3] || "0", row[4] || "0", row[5] || "0", row[6] || "0", row[7] || "0", row[8] || ""
        ]);

        const historyRaw = await this.request("getHistory", "GET");
        const idsToDelete: string[] = data.ids || [];

        // Match indices of rows to delete from sheet 'Lịch sử'
        // For REST client, we rewrite remaining histories
        const remainingHistoryRows = [
          ["Ngày giờ", "Loại", "Tên sản phẩm", "Kệ", "Số bao", "Quy cách", "Số lẻ", "Khối lượng", "Ghi chú", "ID", "Tổng số con"]
        ];

        for (let i = 0; i < historyRaw.length; i++) {
          const rowItem = historyRaw[i];
          const histId = String(rowItem[9]); // col J is ID

          if (idsToDelete.includes(histId)) {
            // Revert inventory quantities
            const type = String(rowItem[1]).toUpperCase().trim();
            const ke = String(rowItem[3]).trim().toUpperCase();
            const tenHang = String(rowItem[2]).trim().toUpperCase();
            const slXoa = parseFloat(rowItem[10]) || 0;

            for (let j = 0; j < thData.length; j++) {
              if (
                String(thData[j][0]).trim().toUpperCase() === ke &&
                String(thData[j][2]).trim().toUpperCase() === tenHang
              ) {
                const cb = parseFloat(thData[j][4]) || 1;
                const cur = parseFloat(thData[j][6]) || 0;

                let res = (type === "NHẬP" || type === "NHAP") ? (cur - slXoa) : (cur + slXoa);
                res = res < 0 ? 0 : res;

                thData[j][3] = Math.floor(res / cb);
                thData[j][5] = res % cb;
                thData[j][6] = res;
              }
            }
          } else {
            remainingHistoryRows.push(rowItem);
          }
        }

        // Filter and write Tổng hợp
        const filteredThData = thData.filter((r: any) => parseFloat(r[6]) > 0);

        // Delete & write Lịch sử
        await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Lịch sử'!A2:K9999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (remainingHistoryRows.length > 1) {
          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Lịch sử'!A1:K" + remainingHistoryRows.length
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: remainingHistoryRows }),
            }
          );
        }

        // Clear & write Tổng hợp
        await handleFetchWithCheck(
          `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
            "'Tổng hợp'!A2:I999"
          )}:clear`,
          { method: "POST", headers }
        );

        if (filteredThData.length > 0) {
          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Tổng hợp'!A2:I" + (filteredThData.length + 1)
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: filteredThData }),
            }
          );
        }

        return { success: true, msg: "Đã xóa lịch sử, cập nhật tồn kho và xóa dòng trống!" };
      }

      if (action === "saveContData") {
        const rows: any[][] = data.data || [];
        const isReplace: boolean = data.isReplace || false;

        if (isReplace) {
          const finishedContName = rows[0][7]; // ContName
          const currentContRaw = await this.request("getContList", "GET", { keyword: "ALL" });
          const unchangedContRows = currentContRaw.filter(
            (r: any) => String(r[7]).toUpperCase() !== finishedContName.toUpperCase()
          );

          // Full writeback
          const header = ["STT", "ID", "Phân Cấp", "Tiêu chuẩn", "Kích Cỡ", "Bề Mặt", "Số bàn giao (NO)", "Tên Container", "Bao", "CB", "Lẻ", "Ghi chú"];
          const finalRows = [header, ...unchangedContRows, ...rows.map((row: any) => [...row.slice(0, 8), 0, 0, 0, ""])];

          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A1:L9999"
            )}:clear`,
            { method: "POST", headers }
          );

          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A1:L" + finalRows.length
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({ values: finalRows }),
            }
          );
        } else {
          // simple append
          const formatToAppend = rows.map((row: any) => [...row.slice(0, 8), 0, 0, 0, ""]);
          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              "'Cont'!A2:L"
            )}:append?valueInputOption=USER_ENTERED`,
            {
              method: "POST",
              headers,
              body: JSON.stringify({ values: formatToAppend }),
            }
          );
        }
        return { success: true };
      }

      if (action === "updateContQuantity") {
        const { grade, standard, size, finish, no, contName, b, cb, l, n, newG, newS, newSz, newF, newNo } = data;
        const currentContRaw = await this.request("getContList", "GET", { keyword: "ALL" });

        let foundRowIdx = -1;
        for (let i = 0; i < currentContRaw.length; i++) {
          const row = currentContRaw[i];
          if (
            String(row[2]) === String(grade) &&
            String(row[3]) === String(standard) &&
            String(row[4]) === String(size) &&
            String(row[5]) === String(finish) &&
            String(row[6]) === String(no) &&
            String(row[7]) === String(contName)
          ) {
            foundRowIdx = i;
            break;
          }
        }

        if (foundRowIdx !== -1) {
          // Column indexes in sheets start from 1. Index i corresponds to row (i + 2)
          // Grade is column 3, Standard 4, Size 5, Finish 6, NO 7, ContName 8, Bao 9, CB 10, Le 11, GhiChu 12
          const targetRowNumber = foundRowIdx + 2;
          const targetRange = `'Cont'!C${targetRowNumber}:L${targetRowNumber}`;

          await handleFetchWithCheck(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
              targetRange
            )}?valueInputOption=USER_ENTERED`,
            {
              method: "PUT",
              headers,
              body: JSON.stringify({
                values: [[newG || grade, newS || standard, newSz || size, newF || finish, newNo || no, contName, b, cb, l, n]],
              }),
            }
          );
          return { success: true };
        }
        return { success: false, msg: "Không tìm thấy dữ liệu khớp với thông tin cũ." };
      }

      return { success: false, msg: "Không tìm thấy driver phù hợp!" };
    }

    throw new Error("Không có driver Google Sheet cấu hình phù hợp!");
  }

  // API Call abstractions
  async getTrong_LuongData(): Promise<WeightMap> {
    if (!this.isConfigured()) {
      return JSON.parse(localStorage.getItem("cteg_weight_map") || "{}");
    }
    try {
      const raw = await this.request("getProductList", "GET");
      const validRaw = validateArrayResponse(raw, "Tải Trọng lượng sản phẩm");
      const weightMap: WeightMap = {};
      validRaw.forEach((row: any) => {
        if (row && row[0]) {
          const name = String(row[0]).trim().toUpperCase();
          const tyTrong = parseFloat(String(row[2]).replace(/,/g, ".")) || 0;
          weightMap[name] = tyTrong;
        }
      });
      return weightMap;
    } catch (e) {
      console.error("Sheet read error getTrong_LuongData:", e);
      throw e;
    }
  }

  async getProductList(): Promise<Product[]> {
    if (!this.isConfigured()) {
      return JSON.parse(localStorage.getItem("cteg_products") || "[]");
    }
    try {
      const raw = await this.request("getProductList", "GET");
      const validRaw = validateArrayResponse(raw, "Tải Danh mục sản phẩm");
      return validRaw.filter((r: any) => r && String(r[0]).trim() !== "").map((r: any) => ({
        ten: String(r[0]).trim(),
        searchStr: String(r[0]).toUpperCase(),
      }));
    } catch (e) {
      console.error("Sheet read error:", e);
      throw e;
    }
  }

  async getRealTimeInventory(): Promise<InventoryItem[]> {
    if (!this.isConfigured()) {
      return JSON.parse(localStorage.getItem("cteg_inventory") || "[]");
    }
    try {
      const raw = await this.request("getRealTimeInventory", "GET");
      const validRaw = validateArrayResponse(raw, "Tải Tồn kho thời gian thực");
      return validRaw.map((r: any, idx: number) => ({
        id: `gs-inv-${idx}-${r[0]}`,
        ke: String(r[0]),
        ten: String(r[2]),
        bao: parseFloat(r[3]) || 0,
        slBao: parseFloat(r[4]) || 0,
        le: parseFloat(r[5]) || 0,
        tongCon: parseFloat(r[6]) || 0,
        khoiLuong: parseFloat(r[7]) || 0,
        ghiChu: String(r[8] || ""),
      }));
    } catch (e) {
      console.error("Sheet read error:", e);
      throw e;
    }
  }

  async getHistory(): Promise<HistoryItem[]> {
    if (!this.isConfigured()) {
      return JSON.parse(localStorage.getItem("cteg_history") || "[]");
    }
    try {
      const raw = await this.request("getHistory", "GET");
      const validRaw = validateArrayResponse(raw, "Tải Lịch sử giao dịch");
      const rows = validRaw.filter((r: any) => r && r[0] && String(r[0]).toLowerCase() !== "ngày giờ");
      return rows.reverse().map((r: any, idx: number) => {
        let dateStr = "";
        try {
          dateStr = new Date(r[0]).toLocaleString("vi-VN", { hour12: true });
        } catch (e) {
          dateStr = String(r[0]);
        }
        return {
          id: r[9] ? `gs-hist-${idx}-${r[9]}` : `gs-hist-${idx}`,
          ngay: dateStr,
          loai: String(r[1]).toUpperCase() === "NHẬP" || String(r[1]).toUpperCase() === "NHAP" ? "NHẬP" : "XUẤT",
          tenHang: String(r[2]),
          ke: String(r[3]),
          bao: parseFloat(r[4]) || 0,
          slBao: parseFloat(r[5]) || 0,
          le: parseFloat(r[6]) || 0,
          tong: parseFloat(r[10]) || 0,
          kg: parseFloat(String(r[7]).replace(/[ \t]*Kg/gi, "")) || 0,
          ghiChu: String(r[8] || ""),
        };
      });
    } catch (e) {
      console.error("Sheet read error:", e);
      throw e;
    }
  }

  async processMultiAdjust(ke: string, items: any[]): Promise<{ success: boolean; msg?: string }> {
    if (!this.isConfigured()) {
      const res = localProcessMultiAdjust(ke, items);
      return res;
    }
    try {
      return await this.request("processMultiAdjust", "POST", { ke, items });
    } catch (e: any) {
      return { success: false, msg: e.toString() };
    }
  }

  async deleteSelectedHistory(ids: string[], password?: string): Promise<{ success: boolean; msg?: string }> {
    if (!this.isConfigured()) {
      const localResult = localDeleteSelectedHistory(ids, password);
      if (localResult === "Sai mật khẩu!") {
        return { success: false, msg: "Sai mật khẩu!" };
      }
      return { success: true, msg: localResult };
    }
    try {
      const result = await this.request("deleteSelectedHistory", "POST", { ids, password });
      if (typeof result === "string" && result === "Sai mật khẩu!") {
        return { success: false, msg: result };
      }
      if (result.success !== undefined) return result;
      return { success: true, msg: String(result) };
    } catch (e: any) {
      return { success: false, msg: e.toString() };
    }
  }

  async getContList(keyword: string): Promise<ContItem[]> {
    if (!this.isConfigured()) {
      const cached = localStorage.getItem("cteg_cont");
      const list: ContItem[] = cached ? JSON.parse(cached) : [];
      if (!keyword || keyword === "" || keyword.toUpperCase() === "ALL") return list;
      return list.filter((r) => String(r.contName).toUpperCase().includes(keyword.toUpperCase()));
    }
    try {
      const raw = await this.request("getContList", "GET", { keyword });
      const validRaw = validateArrayResponse(raw, "Tải Danh sách Container");
      const rows = validRaw.filter((r: any) => r && r[0] && String(r[0]).toUpperCase() !== "STT");
      return rows.map((r: any, idx: number) => ({
        stt: parseInt(r[0]) || idx + 1,
        id: `gs-cont-${idx}-${r[1] || ""}`.trim(),
        grade: String(r[2] || ""),
        standard: String(r[3] || ""),
        size: String(r[4] || ""),
        finish: String(r[5] || ""),
        no: String(r[6] || ""),
        contName: String(r[7] || ""),
        bao: parseFloat(r[8]) || 0,
        slBao: parseFloat(r[9]) || 0,
        le: parseFloat(r[10]) || 0,
        ghiChu: String(r[11] || ""),
      }));
    } catch (e) {
      console.error("Sheet read error:", e);
      throw e;
    }
  }

  async saveContData(parsedItems: ContItem[], isReplace: boolean): Promise<{ success: boolean; msg?: string }> {
    if (!this.isConfigured()) {
      const cached = localStorage.getItem("cteg_cont");
      let contItems: ContItem[] = cached ? JSON.parse(cached) : [];

      if (isReplace) {
        const finishedContName = parsedItems[0].contName;
        const unchanged = contItems.filter((item) => item.contName.toUpperCase() !== finishedContName.toUpperCase());
        contItems = [...unchanged, ...parsedItems];
      } else {
        contItems = [...contItems, ...parsedItems];
      }

      localStorage.setItem("cteg_cont", JSON.stringify(contItems));
      return { success: true };
    }

    try {
      // Data in matrix layout matching sheet 'Cont'
      // Row parameters:
      // row[0] = STT, row[1] = ID, row[2] = Grade, row[3] = Standard, row[4] = Size, row[5] = Finish, row[6] = NO, row[7] = Cont Name,
      // row[8] = Số bao, row[9] = Quy cách (con/bao), row[10] = Số lẻ, row[11] = Ghi chú
      const dataMatrix = parsedItems.map((item) => [
        item.stt,
        item.id,
        item.grade,
        item.standard,
        item.size,
        item.finish,
        item.no,
        item.contName,
        item.bao,
        item.slBao,
        item.le,
        item.ghiChu,
      ]);

      return await this.request("saveContData", "POST", { data: dataMatrix, isReplace });
    } catch (e: any) {
      return { success: false, msg: e.toString() };
    }
  }

  async updateContQuantity(
    grade: string,
    standard: string,
    size: string,
    finish: string,
    no: string,
    contName: string,
    b: number,
    cb: number,
    l: number,
    n: string,
    newG: string,
    newS: string,
    newSz: string,
    newF: string,
    newNo: string
  ): Promise<{ success: boolean; msg?: string }> {
    if (!this.isConfigured()) {
      const cached = localStorage.getItem("cteg_cont");
      const list: ContItem[] = cached ? JSON.parse(cached) : [];

      const updated = list.map((item) => {
        if (
          String(item.grade) === String(grade) &&
          String(item.standard) === String(standard) &&
          String(item.size) === String(size) &&
          String(item.finish) === String(finish) &&
          String(item.no) === String(no) &&
          String(item.contName) === String(contName)
        ) {
          return {
            ...item,
            grade: newG,
            standard: newS,
            size: newSz,
            finish: newF,
            no: newNo,
            bao: b,
            slBao: cb,
            le: l,
            ghiChu: n,
          };
        }
        return item;
      });

      localStorage.setItem("cteg_cont", JSON.stringify(updated));
      return { success: true };
    }

    try {
      return await this.request("updateContQuantity", "POST", {
        grade,
        standard,
        size,
        finish,
        no,
        contName,
        b,
        cb,
        l,
        n,
        newG,
        newS,
        newSz,
        newF,
        newNo,
      });
    } catch (e: any) {
      return { success: false, msg: e.toString() };
    }
  }
}

export const googleSheetsService = new GoogleSheetsService();
