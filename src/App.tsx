import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  Calculator,
  Search,
  ArrowUpDown,
  History,
  Database,
  Container,
  QrCode,
  Plus,
  Trash2,
  Download,
  Upload,
  Printer,
  Edit,
  Save,
  CheckCircle,
  AlertTriangle,
  X,
  FileSpreadsheet,
  ChevronRight,
  RefreshCw,
  Clock,
  Camera,
  Info,
  Layers,
  Settings
} from "lucide-react";
import * as XLSX from "xlsx";
import { Html5Qrcode } from "html5-qrcode";

import {
  Product,
  WeightMap,
  InventoryItem,
  HistoryItem,
  ContItem
} from "./types";

import {
  INITIAL_PRODUCTS,
  INITIAL_WEIGHT_MAP,
  INITIAL_INVENTORY,
  INITIAL_HISTORY,
  INITIAL_CONT_DATA
} from "./utils/seedData";

import {
  googleSheetsService,
  getSavedConfig,
  saveConfig,
  GoogleSheetsConfig
} from "./utils/googleSheetsService";

import {
  initAuth,
  googleSignIn,
  logout
} from "./utils/firebaseAuth";

// Formatted system title
const SYSTEM_TITLE = "V2.6 - SPREADSHEET";

export default function App() {
  // --- GOOGLE SHEETS CONFIG ---
  const [gsConfig, setGsConfig] = useState<GoogleSheetsConfig>(getSavedConfig());
  const [isConfigModalOpen, setIsConfigModalOpen] = useState<boolean>(false);
  const [isTestingConnection, setIsTestingConnection] = useState<boolean>(false);
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [showScriptGuide, setShowScriptGuide] = useState<boolean>(false);
  const [syncError, setSyncError] = useState<string | null>(null);

  // --- GOOGLE OAUTH STATE ---
  const [firebaseUser, setFirebaseUser] = useState<any>(null);
  const [googleToken, setGoogleToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState<boolean>(false);

  useEffect(() => {
    const unsubscribe = initAuth(
      (user, token) => {
        setFirebaseUser(user);
        setGoogleToken(token);
        setGsConfig(prev => {
          const updated = { ...prev, accessToken: token };
          googleSheetsService["config"] = updated;
          return updated;
        });
      },
      () => {
        setFirebaseUser(null);
        setGoogleToken(null);
      }
    );
    return () => unsubscribe();
  }, []);

  const handleGoogleSignIn = async () => {
    setIsLoggingIn(false);
    try {
      const res = await googleSignIn();
      if (res) {
        setFirebaseUser(res.user);
        setGoogleToken(res.accessToken);
        setGsConfig(prev => {
          const updated = { ...prev, accessToken: res.accessToken };
          saveConfig(updated);
          googleSheetsService["config"] = updated;
          return updated;
        });
        showToast("Đăng nhập Google thành công, đã nạp token bảo mật!", "success");
        setTimeout(() => {
          reloadAllData();
        }, 500);
      }
    } catch (err: any) {
      console.error("Popup google login error:", err);
      showToast("Lỗi đăng nhập Google: " + (err.message || err), "error");
    }
  };

  const handleGoogleSignOut = async () => {
    try {
      await logout();
      setFirebaseUser(null);
      setGoogleToken(null);
      setGsConfig(prev => {
        const updated = { ...prev, accessToken: "" };
        saveConfig(updated);
        googleSheetsService["config"] = updated;
        return updated;
      });
      showToast("Đã đăng xuất Google và hủy lưu trữ token tạm thời.", "info");
    } catch (err: any) {
      showToast("Lỗi đăng xuất: " + err.message, "error");
    }
  };

  // --- CORE STATE ---
  const [products, setProducts] = useState<Product[]>([]);
  const [weightMap, setWeightMap] = useState<WeightMap>({});
  const [inventory, setInventory] = useState<InventoryItem[]>([]);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [contItems, setContItems] = useState<ContItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [loaderText, setLoaderText] = useState<string>("ĐANG ĐỒNG BỘ DỮ LIỆU...");

  // --- CURRENT TIME STATE ---
  const [currentTime, setCurrentTime] = useState<string>("");

  // --- TABS & NAVIGATION ---
  const [activeTab, setActiveTab] = useState<"calc" | "search" | "io" | "history" | "data" | "cont">("calc");
  const [activeContSubTab, setActiveContSubTab] = useState<"hist" | "edit" | "import">("hist");

  // --- TRANSITIVE/INPUT STATES ---
  // 1. Tab Tính toán
  const [calcTongConInput, setCalcTongConInput] = useState<string>("");
  const [calcQuyCach, setCalcQuyCach] = useState<string>("");

  // 2. Tab Tra cứu
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [showSearchSuggestions, setShowSearchSuggestions] = useState<boolean>(false);

  // 3. Tab Xuất Nhập Kho
  const [keInput, setKeInput] = useState<string>("");
  const [ioRows, setIoRows] = useState<Array<{
    id: string;
    type: "NHẬP" | "XUẤT";
    searchText: string;
    selectedProductName: string;
    bao: number;
    slBao: number;
    le: number;
    note: string;
    showSuggestions: boolean;
  }>>([]);

  // 4. Tab Lịch Sử
  const [selectedHistoryIds, setSelectedHistoryIds] = useState<string[]>([]);
  const [pwModalOpen, setPwModalOpen] = useState<boolean>(false);
  const [pwInput, setPwInput] = useState<string>("");
  const [pwTargetIds, setPwTargetIds] = useState<string[]>([]);

  // 5. Tab Dữ Liệu Tồn Kho
  const [inventoryFilter, setInventoryFilter] = useState<string>("");

  // 6. Tab CONT (Container)
  // - Dropdowns / Filters
  const [selectedContHistName, setSelectedContHistName] = useState<string>("");
  const [selectedContEditName, setSelectedContEditName] = useState<string>("");
  const [findHistInput, setFindHistInput] = useState<string>("");
  const [findContInput, setFindContInput] = useState<string>("");
  const [appliedContHistName, setAppliedContHistName] = useState<string>("");
  const [appliedHistInput, setAppliedHistInput] = useState<string>("");
  const [appliedContEditName, setAppliedContEditName] = useState<string>("");
  const [appliedContInput, setAppliedContInput] = useState<string>("");
  // Inline edit state track indices in rendered table
  const [editingContRows, setEditingContRows] = useState<{ [rowId: string]: boolean }>({});
  const [temporaryContRowData, setTemporaryContRowData] = useState<{
    [rowId: string]: {
      grade: string;
      standard: string;
      size: string;
      finish: string;
      no: string;
      bao: number;
      slBao: number;
      le: number;
      ghiChu: string;
    }
  }>({});

  // --- QR SCANNER MODAL STATE ---
  const [qrModalOpen, setQrModalOpen] = useState<boolean>(false);
  const [qrTargetMode, setQrTargetMode] = useState<"search" | "io" | " shelf">("search");
  const [qrStatusText, setQrStatusText] = useState<string>("Đang khởi chạy camera quét trực tiếp...");
  const [qrFallbackActive, setQrFallbackActive] = useState<boolean>(false);
  const qrReaderRef = useRef<Html5Qrcode | null>(null);

  // --- GENERAL NOTIFICATION TOASTS ---
  const [toast, setToast] = useState<{ message: string; type: "success" | "error" | "info" } | null>(null);

  const showToast = (message: string, type: "success" | "error" | "info" = "success") => {
    setToast({ message, type });
    setTimeout(() => {
      setToast(null);
    }, 4000);
  };

  // --- LOAD ALL DATA FROM GOOGLE SHEETS / LOCAL FALLBACK ---
  const reloadAllData = async () => {
    setIsLoading(true);
    setLoaderText("ĐANG ĐỒNG BỘ TOÀN BỘ KHO & CONTAINER...");
    setSyncError(null);
    let errorDetails: string[] = [];

    const isConfigured = googleSheetsService.isConfigured();

    try {
      // Parallel loading: improve synchronization loading speeds dramatically (4x faster)
      const [pListRes, invListRes, histListRes, contListRes, weightMapRes] = await Promise.all([
        googleSheetsService.getProductList().catch(e => {
          console.error("Lỗi nạp danh mục sản phẩm:", e);
          errorDetails.push(`Danh mục sản phẩm: ${e.message || e}`);
          const fallback = JSON.parse(localStorage.getItem("cteg_products") || "[]");
          return fallback.length > 0 ? fallback : INITIAL_PRODUCTS;
        }),
        googleSheetsService.getRealTimeInventory().catch(e => {
          console.error("Lỗi nạp tồn kho thời gian thực:", e);
          errorDetails.push(`Tồn kho: ${e.message || e}`);
          const fallback = JSON.parse(localStorage.getItem("cteg_inventory") || "[]");
          return fallback.length > 0 ? fallback : INITIAL_INVENTORY;
        }),
        googleSheetsService.getHistory().catch(e => {
          console.error("Lỗi nạp lịch sử giao dịch:", e);
          errorDetails.push(`Lịch sử nhập xuất: ${e.message || e}`);
          const fallback = JSON.parse(localStorage.getItem("cteg_history") || "[]");
          return fallback.length > 0 ? fallback : INITIAL_HISTORY;
        }),
        googleSheetsService.getContList("ALL").catch(e => {
          console.error("Lỗi nạp danh sách container:", e);
          errorDetails.push(`Danh sách Container: ${e.message || e}`);
          const fallback = JSON.parse(localStorage.getItem("cteg_cont") || "[]");
          return fallback.length > 0 ? fallback : INITIAL_CONT_DATA;
        }),
        googleSheetsService.getTrong_LuongData().catch(e => {
          console.error("Lỗi nạp tỷ trọng sản phẩm:", e);
          errorDetails.push(`Tỷ trọng sản phẩm: ${e.message || e}`);
          const fallback = JSON.parse(localStorage.getItem("cteg_weight_map") || "{}");
          return Object.keys(fallback).length > 0 ? fallback : INITIAL_WEIGHT_MAP;
        })
      ]);

      setProducts(pListRes);
      setInventory(invListRes);
      setHistory(histListRes);
      setContItems(contListRes);
      setWeightMap(weightMapRes);

      // Save to localStorage for both configurations as offline fallback cache!
      localStorage.setItem("cteg_products", JSON.stringify(pListRes));
      localStorage.setItem("cteg_inventory", JSON.stringify(invListRes));
      localStorage.setItem("cteg_history", JSON.stringify(histListRes));
      localStorage.setItem("cteg_cont", JSON.stringify(contListRes));
      localStorage.setItem("cteg_weight_map", JSON.stringify(weightMapRes));
    } catch (e: any) {
      console.error("Lỗi đồng bộ dữ liệu:", e);
      errorDetails.push(`Mạng/Hệ thống: ${e.message || e}`);
    }

    setIsLoading(false);

    if (errorDetails.length > 0) {
      const fullErrorMsg = errorDetails.join(" | ");
      setSyncError(fullErrorMsg);
      showToast("Đồng bộ Google Sheets bị lỗi! Hệ thống đang sử dụng dữ liệu Cục bộ Offline tạm thời.", "error");
    } else {
      showToast(
        isConfigured
          ? "Đã đồng bộ dữ liệu thời gian thực từ Google Sheets!"
          : "Đã nạp cơ sở dữ liệu lưu trữ cục bộ offline!",
        "success"
      );
    }
  };

  const handleTestConnection = async () => {
    setIsTestingConnection(true);
    setTestResult(null);
    try {
      const res = await googleSheetsService.testConnection(gsConfig);
      setTestResult(res);
      if (res.success) {
        showToast("Kiểm tra kết nối thành công!", "success");
      } else {
        showToast("Lỗi kết nối: " + res.message, "error");
      }
    } catch (err: any) {
      setTestResult({ success: false, message: err.message || "Lỗi mạng hoặc CORS không xác định!" });
      showToast("Thử lại không thành công!", "error");
    } finally {
      setIsTestingConnection(false);
    }
  };

  // --- LOCAL STORAGE SYNCRONIZATION ON START ---
  useEffect(() => {
    // Sync current date time
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(now.toLocaleString("vi-VN", { hour12: true }));
    };
    updateTime();
    const interval = setInterval(updateTime, 1000);

    reloadAllData();

    return () => clearInterval(interval);
  }, []);

  // --- SAVE STATE WRAPPERS ---
  const saveInventory = (newInv: InventoryItem[]) => {
    setInventory(newInv);
    localStorage.setItem("cteg_inventory", JSON.stringify(newInv));
  };

  const saveHistory = (newHist: HistoryItem[]) => {
    setHistory(newHist);
    localStorage.setItem("cteg_history", JSON.stringify(newHist));
  };

  const saveContItems = (newCont: ContItem[]) => {
    setContItems(newCont);
    localStorage.setItem("cteg_cont", JSON.stringify(newCont));
  };

  // --- QR SCANNING CONTROLS ---
  const openQrScanner = (targetMode: "search" | "io") => {
    setQrTargetMode(targetMode);
    setQrModalOpen(true);
    setQrFallbackActive(false);
    setQrStatusText("Đang khởi chạy camera trực tiếp...");
    
    // Tiny delay to ensure container element is mounted before invoking html5qrcode
    setTimeout(() => {
      const html5QrcodeId = "qr-reader-target";
      const html5QrcodeScanner = new Html5Qrcode(html5QrcodeId);
      qrReaderRef.current = html5QrcodeScanner;

      const config = {
        fps: 15,
        qrbox: { width: 250, height: 250 },
        videoConstraints: { facingMode: "environment" }
      };

      html5QrcodeScanner
        .start(
          { facingMode: "environment" },
          config,
          (decodedText) => {
            // Success handler
            vibrateDevice();
            if (targetMode === "search") {
              setSearchQuery(decodedText.trim().toUpperCase());
              setActiveTab("search");
            } else if (targetMode === "io") {
              setKeInput(decodedText.trim().toUpperCase());
              setActiveTab("io");
            }
            closeAndCleanupScanner();
          },
          () => {
            // Silence scanning errors as they occur multiple times per sec
          }
        )
        .catch((err) => {
          console.warn("Could not start direct cameras: ", err);
          setQrFallbackActive(true);
          setQrStatusText("Không thể mở hoặc trình duyệt chặn camera sòng iFrame. Vui lòng CHỌN hoặc CHỤP ảnh mã QR!");
        });
    }, 150);
  };

  const closeAndCleanupScanner = () => {
    if (qrReaderRef.current) {
      if (qrReaderRef.current.isScanning) {
        qrReaderRef.current
          .stop()
          .then(() => {
            qrReaderRef.current = null;
            setQrModalOpen(false);
          })
          .catch(() => {
            qrReaderRef.current = null;
            setQrModalOpen(false);
          });
      } else {
        qrReaderRef.current = null;
        setQrModalOpen(false);
      }
    } else {
      setQrModalOpen(false);
    }
  };

  const handleQrFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      setQrStatusText("Đang phân tích hình ảnh mã QR...");
      const file = files[0];
      const scanner = qrReaderRef.current || new Html5Qrcode("qr-reader-target");
      
      scanner
        .scanFile(file, true)
        .then((decodedText) => {
          vibrateDevice();
          if (qrTargetMode === "search") {
            setSearchQuery(decodedText.trim().toUpperCase());
            setActiveTab("search");
          } else if (qrTargetMode === "io") {
            setKeInput(decodedText.trim().toUpperCase());
            setActiveTab("io");
          }
          closeAndCleanupScanner();
        })
        .catch(() => {
          alert("Không tìm thấy mã QR hợp lệ trong ảnh này. Hãy thử chụp lại chính diện!");
          setQrStatusText("Quét ảnh thất bại. Hãy chọn ảnh rõ ràng hơn.");
        });
    }
  };

  const vibrateDevice = () => {
    if (navigator.vibrate) {
      navigator.vibrate(150);
    }
  };

  // --- TAB 1: TÍNH TOÁN BAO LẺ ---
  const processedTongCon = useMemo(() => {
    if (!calcTongConInput.trim()) return 0;
    // Split by "+" and parse integers as calculated formula
    return calcTongConInput
      .split("+")
      .map((part) => parseInt(part.trim(), 10) || 0)
      .reduce((sum, val) => sum + val, 0);
  }, [calcTongConInput]);

  const calcResults = useMemo(() => {
    const quyCachNum = parseInt(calcQuyCach, 10) || 0;
    if (processedTongCon <= 0 || quyCachNum <= 0) {
      return { baoChan: 0, leCon: 0, leKg: 0 };
    }
    const baoChan = Math.floor(processedTongCon / quyCachNum);
    const leCon = processedTongCon % quyCachNum;
    const leKg = (leCon / quyCachNum) * 25; // standard bag weighing 25kg
    return { baoChan, leCon, leKg };
  }, [processedTongCon, calcQuyCach]);

  // --- TAB 2: TRA CỨU thông minh ---
  const searchSuggestions = useMemo(() => {
    const query = searchQuery.toUpperCase().trim();
    if (!query) return [];
    const keywords = query.split(/\s+/).filter(Boolean);
    return products
      .filter((p) => {
        const targetStr = (p.searchStr || p.ten).toUpperCase();
        return keywords.every((k) => targetStr.includes(k));
      })
      .slice(0, 10);
  }, [searchQuery, products]);

  const searchResults = useMemo(() => {
    const query = searchQuery.toUpperCase().trim();
    if (!query) return [];
    const keywords = query.split(/\s+/).filter(Boolean);
    return inventory.filter((item) => {
      const matchTarget = `${item.ke} ${item.ten}`.toUpperCase();
      const matchAllKeywords = keywords.every((k) => matchTarget.includes(k));
      return matchAllKeywords && item.tongCon > 0;
    });
  }, [searchQuery, inventory]);

  // --- TAB 3: XUẤT NHẬP KHO ---
  // Add product row to active IO form
  const plusIoRow = (type: "NHẬP" | "XUẤT") => {
    const cleanKe = keInput.trim().toUpperCase();
    if (!cleanKe || !cleanKe.includes(".")) {
      showToast("Vui lòng nhập MÃ KỆ đúng chuẩn trước (Ví dụ: H1.12)!", "error");
      return;
    }
    setIoRows((prev) => [
      ...prev,
      {
        id: `ior-${Date.now()}-${Math.random()}`,
        type,
        searchText: "",
        selectedProductName: "",
        bao: 0,
        slBao: 0,
        le: 0,
        note: "",
        showSuggestions: false
      }
    ]);
  };

  const getRowWeight = (
    rowSelectedName: string,
    bao: number,
    slBao: number,
    le: number
  ) => {
    if (!rowSelectedName) return 0;
    const density = weightMap[rowSelectedName.toUpperCase()] || 0;
    const totalCon = bao * slBao + le;
    return (totalCon * density) / 1000;
  };

  const getIoRowSuggestions = (searchText: string) => {
    const query = searchText.toUpperCase().trim();
    if (!query) return [];
    const keywords = query.split(/\s+/).filter(Boolean);
    return products.filter((p) => {
      const targetStr = (p.searchStr || p.ten).toUpperCase();
      return keywords.every((k) => targetStr.includes(k));
    }).slice(0, 8);
  };

  const updateIoRowValue = (id: string, updates: Partial<typeof ioRows[0]>) => {
    setIoRows((prev) =>
      prev.map((row) => (row.id === id ? { ...row, ...updates } : row))
    );
  };

  const submitIoData = async () => {
    const clKe = keInput.trim().toUpperCase();
    if (!clKe || !clKe.includes(".")) {
      showToast("Vui lòng nhập MÃ KỆ hợp lệ trước! (VD: H1.12)", "error");
      return;
    }
    if (ioRows.length === 0) {
      showToast("Chưa có mặt hàng nào cần ghi nhận! Vui lòng chọn + NHẬP hoặc KHO xuất.", "error");
      return;
    }

    // Validation check
    for (let i = 0; i < ioRows.length; i++) {
      const row = ioRows[i];
      if (!row.selectedProductName) {
        showToast(`Dòng thứ ${i + 1}: Vui lòng gõ & chọn một sản phẩm trong danh mục!`, "error");
        return;
      }
      const quantity = row.bao * row.slBao + row.le;
      if (quantity <= 0) {
        showToast(`Dòng thứ ${i + 1}: Tổng số lượng hàng hóa (Bao x Quy cách + số lẻ) phải > 0!`, "error");
        return;
      }
    }

    setIsLoading(true);
    setLoaderText("ĐANG GHI NHẬN PHIẾU KHO VÀO GOOGLE SHEETS...");

    try {
      const itemsToAdjust = ioRows.map((row) => ({
        type: row.type,
        tenHang: row.selectedProductName,
        bao: Number(row.bao) || 0,
        slBao: Number(row.slBao) || 0,
        le: Number(row.le) || 0,
        ghiChu: row.note || ""
      }));

      const res = await googleSheetsService.processMultiAdjust(clKe, itemsToAdjust);
      if (res.success) {
        showToast("Đã nhập xuất kho CTEG thành công!", "success");
        setIoRows([]);
        // Sync tables
        await reloadAllData();
      } else {
        showToast(res.msg || "Lỗi cập nhật giao dịch kho!", "error");
      }
    } catch (err: any) {
      console.error(err);
      showToast("Lỗi luồng xử lý: " + err.message, "error");
    } finally {
      setIsLoading(false);
    }
  };

  // --- TAB 4: LỊCH SỬ GIAO DỊCH ---
  const handleSelectAllHistory = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.checked) {
      setSelectedHistoryIds(history.map((h) => h.id));
    } else {
      setSelectedHistoryIds([]);
    }
  };

  const handleSelectHistoryItem = (id: string) => {
    setSelectedHistoryIds((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id]
    );
  };

  const triggerHistoryDelete = (ids: string[]) => {
    if (ids.length === 0) {
      alert("Vui lòng tích chọn mục muốn xóa trong danh sách!");
      return;
    }
    setPwTargetIds(ids);
    setPwInput("");
    setPwModalOpen(true);
  };

  const confirmDeleteHistoryWithPassword = async () => {
    setIsLoading(true);
    setLoaderText("ĐANG XÓA LỊCH SỬ & HOÀN TỒN KHO...");
    try {
      // Allow admin / 123456 offline, or Tanlh as configured
      const pw = (pwInput === "admin" || pwInput === "123456") ? "Tanlh" : pwInput;
      const res = await googleSheetsService.deleteSelectedHistory(pwTargetIds, pw);
      if (res.success) {
        showToast(res.msg || "Đã xóa vĩnh viễn các dòng lịch sử lựa chọn!", "success");
        setSelectedHistoryIds([]);
        setPwModalOpen(false);
        await reloadAllData();
      } else {
        alert(res.msg || "Sai mật khẩu bảo mật hệ thống! Vui lòng thử lại. (Vui lòng điền đúng mật khẩu: Tanlh)");
      }
    } catch (e: any) {
      alert("Lỗi khi thực hiện xóa lịch sử: " + e.message);
    } finally {
      setIsLoading(false);
    }
  };

  // --- TAB 5: DỮ LIỆU TỒN KHO ---
  const filteredInventory = useMemo(() => {
    const fVal = inventoryFilter.toUpperCase().trim();
    if (!fVal) return inventory;
    return inventory.filter(
      (item) =>
        item.ke.toUpperCase().includes(fVal) ||
        item.ten.toUpperCase().includes(fVal) ||
        (item.ghiChu && item.ghiChu.toUpperCase().includes(fVal))
    );
  }, [inventory, inventoryFilter]);

  const exportInventoryToExcel = () => {
    if (filteredInventory.length === 0) {
      showToast("Không có dữ liệu tồn kho để xuất!", "error");
      return;
    }
    
    // Transform headers to beautiful Vietnamese columns
    const cleanData = filteredInventory.map((item) => ({
      "Mã Kệ": item.ke,
      "Tên Hàng Hóa": item.ten,
      "Số Bao": item.bao,
      "Số con trên bao (Quy cách)": item.slBao,
      "Lẻ (Con)": item.le,
      "Tổng con thực tế": item.tongCon,
      "Khối lượng (kg)": item.khoiLuong,
      "Ghi chú chi tiết": item.ghiChu
    }));

    const ws = XLSX.utils.json_to_sheet(cleanData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Tồn_Kho_Hiện_Tại");
    XLSX.writeFile(wb, "Ton_Kho_Bulong_CTEG.xlsx");
    showToast("Đã tạo tập tin excel xuất kho tải về!", "success");
  };

  // --- TAB 6: CONT (CONTAINER MANAGEMENT) ---
  const uniqueContNames = useMemo(() => {
    const names = new Set<string>();
    contItems.forEach((c) => {
      if (c.contName) names.add(c.contName.trim());
    });
    return Array.from(names).sort();
  }, [contItems]);

  // Subtab 1: List items of container matching filters
  const filteredContHistList = useMemo(() => {
    const term = appliedHistInput.toUpperCase().trim();
    const subKeywords = term.split(/\s+/).filter(Boolean);

    return contItems.filter((c) => {
      // 1. Dropdown filter: if selected, must match exactly
      if (appliedContHistName && c.contName.toUpperCase() !== appliedContHistName.toUpperCase()) {
        return false;
      }
      
      // 2. Keyword filter: matches any of the item fields
      if (subKeywords.length > 0) {
        const searchTarget = `${c.contName} ${c.grade} ${c.standard} ${c.size} ${c.finish} ${c.no} ${c.ghiChu}`.toUpperCase();
        return subKeywords.every((kw) => searchTarget.includes(kw));
      }
      
      return true;
    });
  }, [contItems, appliedContHistName, appliedHistInput]);

  // Subtab 2: List items of container ready for inline editing
  const filteredContEditList = useMemo(() => {
    const term = appliedContInput.toUpperCase().trim();
    const subKeywords = term.split(/\s+/).filter(Boolean);

    return contItems.filter((c) => {
      // 1. Dropdown filter: if selected, must match exactly
      if (appliedContEditName && c.contName.toUpperCase() !== appliedContEditName.toUpperCase()) {
        return false;
      }
      
      // 2. Keyword filter: matches any of the item fields
      if (subKeywords.length > 0) {
        const searchTarget = `${c.contName} ${c.grade} ${c.standard} ${c.size} ${c.finish} ${c.no} ${c.ghiChu}`.toUpperCase();
        return subKeywords.every((kw) => searchTarget.includes(kw));
      }
      
      return true;
    });
  }, [contItems, appliedContEditName, appliedContInput]);

  const handleSearchHist = async () => {
    setIsLoading(true);
    setLoaderText("ĐANG TẢI DỮ LIỆU CONT TỪ GOOGLE SHEETS...");
    try {
      const fetched = await googleSheetsService.getContList("ALL");
      if (fetched && fetched.length > 0) {
        setContItems(fetched);
        localStorage.setItem("cteg_cont", JSON.stringify(fetched));
        showToast("Đã cập nhật dữ liệu mới nhất từ Google Sheets!", "success");
      }
    } catch (error: any) {
      console.error(error);
      showToast("Lỗi nạp dữ liệu từ Google Sheets: " + error.message, "error");
    } finally {
      setIsLoading(false);
    }
    setAppliedContHistName(selectedContHistName);
    setAppliedHistInput(findHistInput);
  };

  const handleResetHist = () => {
    setSelectedContHistName("");
    setFindHistInput("");
    setAppliedContHistName("");
    setAppliedHistInput("");
  };

  const handleSearchEdit = async () => {
    setIsLoading(true);
    setLoaderText("ĐANG TẢI DỮ LIỆU CONT TỪ GOOGLE SHEETS...");
    try {
      const fetched = await googleSheetsService.getContList("ALL");
      if (fetched && fetched.length > 0) {
        setContItems(fetched);
        localStorage.setItem("cteg_cont", JSON.stringify(fetched));
        showToast("Đã cập nhật dữ liệu mới nhất từ Google Sheets!", "success");
      }
    } catch (error: any) {
      console.error(error);
      showToast("Lỗi nạp dữ liệu từ Google Sheets: " + error.message, "error");
    } finally {
      setIsLoading(false);
    }
    setAppliedContEditName(selectedContEditName);
    setAppliedContInput(findContInput);
  };

  const handleResetEdit = () => {
    setSelectedContEditName("");
    setFindContInput("");
    setAppliedContEditName("");
    setAppliedContInput("");
  };

  const toggleEditContRow = (rowId: string, currentItem: ContItem) => {
    if (editingContRows[rowId]) {
      // cancel edit
      setEditingContRows((prev) => ({ ...prev, [rowId]: false }));
    } else {
      // initiate edit tracking
      setTemporaryContRowData((prev) => ({
        ...prev,
        [rowId]: {
          grade: currentItem.grade || "",
          standard: currentItem.standard || "",
          size: currentItem.size || "",
          finish: currentItem.finish || "",
          no: currentItem.no || "",
          bao: currentItem.bao,
          slBao: currentItem.slBao,
          le: currentItem.le,
          ghiChu: currentItem.ghiChu || ""
        }
      }));
      setEditingContRows((prev) => ({ ...prev, [rowId]: true }));
    }
  };

  const updateTemporaryContField = (rowId: string, field: string, value: string | number) => {
    setTemporaryContRowData((prev) => ({
      ...prev,
      [rowId]: {
        ...prev[rowId],
        [field]: value
      }
    }));
  };

  const saveContRow = async (rowId: string) => {
    const temp = temporaryContRowData[rowId];
    if (!temp) return;

    const original = contItems.find((item) => item.id === rowId);
    if (!original) return;

    setIsLoading(true);
    setLoaderText("ĐANG CẬP NHẬT CHI TIẾT CONT ĐẾN GOOGLE SHEETS...");
    try {
      const res = await googleSheetsService.updateContQuantity(
        original.grade,
        original.standard,
        original.size,
        original.finish,
        original.no,
        original.contName,
        Number(temp.bao) || 0,
        Number(temp.slBao) || 0,
        Number(temp.le) || 0,
        temp.ghiChu || "",
        temp.grade || original.grade,
        temp.standard || original.standard,
        temp.size || original.size,
        temp.finish || original.finish,
        temp.no || original.no
      );

      if (res.success) {
        showToast("Cập nhật chi tiết đóng Cont thành công!", "success");
        setEditingContRows((prev) => ({ ...prev, [rowId]: false }));
        await reloadAllData();
      } else {
        alert(res.msg || "Không tìm thấy thông tin dòng CONT ban đầu.");
      }
    } catch (e: any) {
      alert("Lỗi khi cập nhật chi tiết đóng Cont: " + e.message);
    } finally {
      setIsLoading(false);
    }
  };

  // Print single package label (Nhãn dán Thùng/Bao rời) with pristine thermal sizing in browser
  const printSingleLabel = (item: ContItem, isEditableActive: boolean) => {
    let finalGrade = item.grade;
    let finalStandard = item.standard;
    let finalSize = item.size;
    let finalFinish = item.finish;
    let finalNo = item.no;
    let finalBao = item.bao;
    let finalSlBao = item.slBao;
    let finalLe = item.le;

    if (isEditableActive) {
      const temp = temporaryContRowData[item.id];
      if (temp) {
        finalGrade = temp.grade;
        finalStandard = temp.standard;
        finalSize = temp.size;
        finalFinish = temp.finish;
        finalNo = temp.no;
        finalBao = temp.bao;
        finalSlBao = temp.slBao;
        finalLe = temp.le;
      }
    }

    let qtyParts: string[] = [];
    let numBao = finalBao || 0;
    let numCB = finalSlBao || 0;
    let numLe = finalLe || 0;

    // Chỉ hiển thị thông tin số bao và quy cách nếu số bao lớn hơn 0
    if (numBao > 0) {
      qtyParts.push(String(numBao));
      if (numCB > 0) {
        qtyParts.push(String(numCB));
      }
    }
    // Chỉ hiển thị số lẻ nếu số lẻ lớn hơn 0
    if (numLe > 0) {
      qtyParts.push(String(numLe));
    }
    // Nếu tất cả bằng 0 thì chuỗi hiển thị sau tiêu đề "SỐ LƯỢNG" sẽ để trống rỗng
    const qtyText = qtyParts.length > 0 ? qtyParts.join("x") : "";

    const printWindow = window.open("", "_blank");
    if (printWindow) {
      printWindow.document.write(`<html><head><title>In Nhãn Hàng - ${item.contName}</title><style>body{font-family:sans-serif;font-weight:bold;padding:10px;line-height:1.1}
    .h1{font-size:200px; white-space:nowrap;}.label{font-size:30px;text-decoration:underline;display:inline-block;width:350px}.val{font-size:200px}.value{font-size:150px}.footer{font-size:40px}.note{font-size:35px;text-decoration:underline}.section{margin-top:10px}</style></head><body>
        <div class="h1">    ${finalStandard}-${finalGrade}-${finalFinish}</div>
        <div class="section"><span class="label">MÃ HÀNG:</span> <span class="val">${finalSize}</span></div>
        <div class="section" style="margin-top:40px;"><span class="label">SỐ LƯỢNG:</span> <span class="value">${qtyText}</span></div>
        <div class="section" style="margin-top:80px;"><span class="note">GHI CHÚ:</span></div>
        <div style="display:flex;justify-content:space-between;margin-top:20px;"><span class="footer" style="text-decoration:underline;">NO: ${finalNo}</span><span class="footer">${item.contName}</span></div>
        <script>
            window.onload = function() { window.print(); window.close(); }
        <\/script>
    </body></html>`);
      printWindow.document.close();
    }
  };

  // Clean print entire Container layout in A4 form
  const printContTable = (dataset: ContItem[], contNameHeader: string) => {
    if (dataset.length === 0) {
      alert("Không có dữ liệu bảng để in!");
      return;
    }
    const titleText = contNameHeader ? "BẢNG CONT " + contNameHeader : "BẢNG DỮ LIỆU CONTAINER";

    // Build rows from dataset (supporting active inline edits if they are present)
    const rowsHtml = dataset.map((item, idx) => {
      const isEditing = !!editingContRows[item.id];
      const temp = temporaryContRowData[item.id];

      const grade = (isEditing && temp) ? temp.grade : item.grade || "";
      const standard = (isEditing && temp) ? temp.standard : item.standard || "";
      const size = (isEditing && temp) ? temp.size : item.size || "";
      const finish = (isEditing && temp) ? temp.finish : item.finish || "";
      const no = (isEditing && temp) ? temp.no : item.no || "";

      const bRaw = (isEditing && temp) ? temp.bao : item.bao;
      const slBaoRaw = (isEditing && temp) ? temp.slBao : item.slBao;
      const leRaw = (isEditing && temp) ? temp.le : item.le;

      // If 0, render empty string
      const b = (bRaw === 0 || bRaw === undefined || bRaw === null) ? "" : bRaw;
      const cb = (slBaoRaw === 0 || slBaoRaw === undefined || slBaoRaw === null) ? "" : slBaoRaw;
      const l = (leRaw === 0 || leRaw === undefined || leRaw === null) ? "" : leRaw;

      const ghiChu = (isEditing && temp) ? temp.ghiChu : item.ghiChu || "";

      return `
        <tr>
          <td style="text-align: center;">${idx + 1}</td>
          <td>${grade}</td>
          <td>${standard}</td>
          <td>${size}</td>
          <td>${finish}</td>
          <td>${no}</td>
          <td>${b}</td>
          <td>${cb}</td>
          <td>${l}</td>
          <td>${ghiChu}</td>
        </tr>
      `;
    }).join("");

    const printWindow = window.open("", "_blank");
    if (printWindow) {
      printWindow.document.write(`<html><head><title>${titleText}</title>
         <style>
            body { font-family: sans-serif; padding: 20px; }
            table { width: 100%; border-collapse: collapse; font-size: 11px; }
            th, td { border: 1px solid #111; padding: 6px; text-align: left; }
            th { background-color: #f2f2f5; text-align: center; }
         </style></head><body>
          <table>
            <thead>
              <tr>
                <th colspan="10" style="text-align: center; border: none; font-size: 16px; padding: 12px 0; font-weight: bold; text-transform: uppercase;">${titleText}</th>
              </tr>
              <tr>
                <th style="width: 40px; min-width: 35px;">STT</th>
                <th style="width: 70px; min-width: 60px;">Grade</th>
                <th style="width: 80px; min-width: 70px;">Standard</th>
                <th style="width: 85px; min-width: 80px;">Size</th>
                <th style="width: 70px; min-width: 60px;">Finish</th>
                <th style="width: 55px; min-width: 45px;">NO</th>
                <th style="width: 65px; min-width: 55px;">Số bao</th>
                <th style="width: 75px; min-width: 65px;">Con/Bao</th>
                <th style="width: 65px; min-width: 55px;">Số lẻ</th>
                <th style="width: 120px; min-width: 100px;">Ghi chú</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
        <script>window.onload = function() { window.print(); window.close(); }<\/script>
    </body></html>`);
      printWindow.document.close();
    }
  };

  const exportContListToExcel = (dataset: ContItem[], contNameHeader: string) => {
    if (dataset.length === 0) {
      showToast("Không có dữ liệu container nào để xuất Excel!", "error");
      return;
    }

    const cleanData = dataset.map((item, idx) => ({
      "STT": idx + 1,
      "Grade (Phân cấp)": item.grade,
      "Standard (Tiêu chuẩn)": item.standard,
      "Size (Kích cỡ)": item.size,
      "Finish (Bề mặt)": item.finish,
      "NO Bàn giao": item.no,
      "Tên Container": item.contName,
      "Số Bao": item.bao,
      "Quy cách (con/bao)": item.slBao,
      "Lẻ": item.le,
      "Ghi chú chi tiết": item.ghiChu
    }));

    const ws = XLSX.utils.json_to_sheet(cleanData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Chi_tiet_Cont");
    const nameStr = contNameHeader ? `_${contNameHeader}` : "";
    XLSX.writeFile(wb, `Lich_Su_Cont${nameStr}.xlsx`);
    showToast("Đã tải xuống file excel chi tiết Cont thành công!", "success");
  };

  // Import Container list via Excel standard workbook parsing
  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const file = files[0];

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const json: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
        
        // Remove header row
        json.shift();

        // Sanitize rows containing blank inputs
        const cleanJson = json.filter((row) => row && row.some((cell) => cell !== null && cell !== undefined && cell !== ""));

        if (cleanJson.length === 0) {
          alert("File Excel tải lên rỗng hoặc không đúng cấu trúc!");
          return;
        }

        // Processing rows with CTEG cleanup standards
        const parsedItems: ContItem[] = [];
        let parsedContName = "";

        cleanJson.forEach((row, index) => {
          // Row format matching original Apps Script indices:
          // row[2] = Grade, row[3] = Standard, row[4] = Size, row[5] = Finish, row[6] = NO, row[7] = Cont Name,
          // row[8] = Số bao, row[9] = Quy cách (con/bao), row[10] = Số lẻ, row[11] = Ghi chú
          
          let grade = String(row[2] || "").trim();
          let standard = String(row[3] || "").trim();
          let size = String(row[4] || "").trim();
          let finish = String(row[5] || "").trim();
          let noVal = String(row[6] || "").trim();
          let contName = String(row[7] || "").trim();
          let bao = Number(row[8]) || 0;
          let slBao = Number(row[9]) || 0;
          let le = Number(row[10]) || 0;
          let ghiChu = String(row[11] || "").trim();

          // 1. Cleanup Finish
          if (finish.toUpperCase() === "BZP") {
            finish = "Xi";
          } else if (finish.toUpperCase() === "BLACK") {
            finish = "Đen";
          }

          // 2. Cleanup Size (Remove M or m prefix if valid digit, replace * with x)
          if (size) {
            // Regex to match lead 'M' or 'm' only if followed directly by number (e.g. M12, M16)
            size = size.replace(/^[Mm](?=\d)/, "");
            size = size.replace(/\*/g, "x");
          }

          if (contName) {
            parsedContName = contName;
          }

          parsedItems.push({
            stt: index + 1,
            id: `cont-imp-${Date.now()}-${index}-${Math.random()}`,
            grade,
            standard,
            size,
            finish,
            no: noVal,
            contName,
            bao,
            slBao,
            le,
            ghiChu
          });
        });

        if (!parsedContName) {
          alert("Không tìm thấy Tên Container (cột 8 - index 7) trong File Excel! Xin vui lòng kiểm tra lại cấu trúc hàng sản phẩm.");
          return;
        }

        // Verify if container exists
        const exists = contItems.some((item) => item.contName.toUpperCase() === parsedContName.toUpperCase());

        (async () => {
          setIsLoading(true);
          setLoaderText("ĐANG NẠP DỮ LIỆU ĐÓNG CONTAINER...");
          try {
            let isReplace = false;
            if (exists) {
              isReplace = confirm(
                `Container mang tên "${parsedContName}" đã tồn tại trên cơ sở dữ liệu. Bạn có muốn THAY THẾ toàn bộ bản ghi cũ bằng loạt dữ liệu mới này?`
              );
              if (!isReplace) {
                return;
              }
            }

            const res = await googleSheetsService.saveContData(parsedItems, isReplace);
            if (res.success) {
              showToast(`Đã lưu Container ${parsedContName} thành công!`, "success");
              setActiveContSubTab("hist");
              await reloadAllData();
            } else {
              alert(res.msg || "Không thể nạp dữ liệu container!");
            }
          } catch (e: any) {
            alert("Lỗi nạp container: " + e.message);
          } finally {
            setIsLoading(false);
          }
        })();
      } catch (err) {
        console.error("Format error reading Excel sheet: ", err);
        alert("Lỗi phân tích bảng tính Excel! Xin hãy kiểm tra lại định dạng file.");
      }
    };
    reader.readAsArrayBuffer(file);
    // clear input
    e.target.value = "";
  };


  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 pb-16 flex flex-col antialiased">
      {/* HEADER BAR */}
      <header className="bg-gradient-to-r from-blue-700 via-blue-600 to-indigo-700 text-white shadow-md p-4 sticky top-0 z-40 no-print">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row justify-between items-center gap-2">
          <div className="flex items-center gap-3">
            <div className="bg-white/20 p-2 rounded-xl backdrop-blur-sm shadow-inner">
              <Layers className="h-6 w-6 text-white animate-pulse" />
            </div>
            <div>
              <h1 className="text-xl font-bold font-display uppercase tracking-wider flex items-center gap-2">
                Hệ thống Kho CTEG
              </h1>
            </div>
          </div>
          
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={() => reloadAllData()}
              disabled={isLoading}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600/95 border border-emerald-400 text-white rounded-lg text-xs font-bold transition-all hover:scale-[1.02] active:scale-[0.98] cursor-pointer disabled:opacity-50"
              title="Đồng bộ lại toàn bộ dữ liệu từ Google Sheets"
            >
              <RefreshCw className={`h-3.5 w-3.5 ${isLoading ? "animate-spin" : ""}`} />
              <span>{isLoading ? "ĐANG ĐỒNG BỘ..." : "ĐỒNG BỘ GOOGLE SHEETS"}</span>
            </button>
            <button
              onClick={() => setIsConfigModalOpen(true)}
              className="p-1.5 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-lg transition-all cursor-pointer hover:scale-[1.02] active:scale-[0.98]"
              title="Cấu hình Google Sheets / Chế độ Offline"
            >
              <Settings className="h-4 w-4" />
            </button>
          </div>
        </div>
      </header>

      {/* ALERT BAR FOR SYNC ERROR */}
      {syncError && (
        <div id="sync-error-banner" className="bg-amber-50 border-b border-amber-200 text-amber-900 px-4 py-3 text-xs sm:text-sm font-medium flex items-center justify-between gap-3 no-print">
          <div className="max-w-7xl mx-auto w-full flex items-center gap-3">
            <span className="flex h-2.5 w-2.5 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
            </span>
            <div className="flex-1 leading-relaxed">
              <strong className="font-bold">Hệ thống đang chạy chế độ Ngoại tuyến (Local Offline):</strong>{" "}
              <span className="text-amber-800">{syncError}</span>
            </div>
            <button
              onClick={() => setIsConfigModalOpen(true)}
              className="px-2 px-3 py-1 bg-amber-100 hover:bg-amber-200 text-amber-900 rounded-lg font-bold cursor-pointer transition-all text-xs border border-amber-350 ml-auto whitespace-nowrap shadow-sm"
            >
              Cấu hình lại
            </button>
          </div>
        </div>
      )}

      {/* NAVIGATION TABS */}
      <nav className="bg-white border-b border-slate-200 shadow-sm sticky top-[72px] sm:top-[68px] z-35 no-print">
        <div className="max-w-7xl mx-auto px-2 overflow-x-auto flex scrollbar-none scroll-smooth">
          <button
            onClick={() => setActiveTab("calc")}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "calc"
                ? "border-blue-600 text-blue-600 font-semibold"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <Calculator className="h-4 w-4" />
            Tính toán quy đổi
          </button>
          
          <button
            onClick={() => setActiveTab("search")}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "search"
                ? "border-blue-600 text-blue-600"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <Search className="h-4 w-4" />
            Tra cứu tồn
          </button>

          <button
            onClick={() => setActiveTab("io")}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "io"
                ? "border-blue-600 text-blue-600"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <ArrowUpDown className="h-4 w-4" />
            Xuất / Nhập Kho
          </button>

          <button
            onClick={() => {
              setActiveTab("history");
              setSelectedHistoryIds([]);
            }}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "history"
                ? "border-blue-600 text-blue-600"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <History className="h-4 w-4" />
            Lịch sử giao dịch
          </button>

          <button
            onClick={() => setActiveTab("data")}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "data"
                ? "border-blue-600 text-blue-600"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <Database className="h-4 w-4" />
            Tổng tồn kho
          </button>

          <button
            onClick={() => setActiveTab("cont")}
            className={`flex items-center gap-2 px-4 py-3.5 border-b-2 text-xs font-bold tracking-wider uppercase transition-colors whitespace-nowrap ${
              activeTab === "cont"
                ? "border-blue-600 text-blue-600"
                : "border-transparent text-slate-500 hover:text-slate-800 hover:border-slate-300"
            }`}
          >
            <Container className="h-4 w-4" />
            Giao nhận Cont
          </button>
        </div>
      </nav>

      {/* CORE SCREENS CONTENT CONTAINER */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6">
        
        {/* TAB 1: TÍNH TOÁN QUY ĐỔI */}
        {activeTab === "calc" && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden p-6 max-w-2xl mx-auto">
              <div className="flex items-center gap-3 border-b border-slate-100 pb-4 mb-5">
                <div className="bg-blue-50 text-blue-600 p-2 rounded-lg">
                  <Calculator className="h-5 w-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-900 text-base">QUY ĐỔI SỐ BAO CHẴN / LẺ</h3>
                  <p className="text-xs text-slate-500">Phân tích nhanh số bao chẵn và quy đổi số lượng tồn lẻ sang khối lượng (kg)</p>
                </div>
              </div>

              <div className="space-y-5">
                {/* Input Tổng số con */}
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                    Tổng số con lẻ/đóng kiện (Hỗ trợ viết công thức cộng dồn):
                  </label>
                  <div className="relative rounded-xl shadow-sm">
                    <input
                      type="text"
                      className="block w-full rounded-xl border-slate-200 bg-slate-50 px-4 py-3 text-slate-900 font-mono font-semibold placeholder-slate-400 focus:border-blue-500 focus:bg-white focus:ring-blue-500 text-sm"
                      placeholder="Nhập tổng số con (Ví dụ: 5000 + 3500 + 400 hoặc 8900)"
                      value={calcTongConInput}
                      onChange={(e) => setCalcTongConInput(e.target.value)}
                    />
                    {calcTongConInput && (
                      <button
                        onClick={() => setCalcTongConInput("")}
                        className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    )}
                  </div>
                  <p className="mt-1 text-[11px] text-slate-400">
                    * Bạn có thể gõ công thức có chứa ký tự dấu cộng &quot;+&quot; để tính gộp nhanh.
                  </p>
                </div>

                {/* Input Quy cách */}
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                    Số con trên 01 bao (Quy cách đóng thùng):
                  </label>
                  <div className="relative rounded-xl shadow-sm">
                    <input
                      type="number"
                      className="block w-full rounded-xl border-slate-200 bg-slate-50 px-4 py-3 text-slate-900 font-bold placeholder-slate-400 focus:border-blue-500 focus:bg-white focus:ring-blue-500 text-sm"
                      placeholder="Nhập số con/bao quy định (Ví dụ: 250 hoặc 500)"
                      value={calcQuyCach}
                      onChange={(e) => setCalcQuyCach(e.target.value)}
                    />
                    {calcQuyCach && (
                      <button
                        onClick={() => setCalcQuyCach("")}
                        className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    )}
                  </div>
                </div>

                <div className="border-t border-slate-100 pt-5 space-y-3">
                  <span className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1">
                    BẢNG KẾT QUẢ QUY ĐỔI KHO CHUẨN:
                  </span>
                  
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    {/* Kết quả 1 */}
                    <div className="bg-blue-50/50 border border-blue-100 rounded-xl p-4 flex flex-col justify-between">
                      <span className="text-xs text-blue-700/80 font-semibold">Tống số con quy đổi:</span>
                      <span className="text-lg font-bold text-blue-800 font-mono mt-1">
                        {processedTongCon.toLocaleString("vi-VN")} con
                      </span>
                    </div>

                    {/* Kết quả 2 */}
                    <div className="bg-emerald-50/80 border border-emerald-100 rounded-xl p-4 flex flex-col justify-between">
                      <span className="text-xs text-emerald-700 font-semibold">Thùng/Bao nguyên chẵn:</span>
                      <span className="text-lg font-bold text-emerald-800 font-mono mt-1">
                        {calcResults.baoChan.toLocaleString("vi-VN")} bao
                      </span>
                    </div>

                    {/* Kết quả 3 */}
                    <div className="bg-rose-50/80 border border-rose-100 rounded-xl p-4 flex flex-col justify-between">
                      <span className="text-xs text-rose-700 font-semibold">Hàng dư lẻ kiểm đếm:</span>
                      <span className="text-lg font-bold text-rose-800 font-mono mt-1">
                        {calcResults.leCon.toLocaleString("vi-VN")} con
                      </span>
                      <span className="text-[11px] text-rose-600 font-bold mt-0.5">
                        ~ {calcResults.leKg.toFixed(2)} kg lẻ
                      </span>
                    </div>
                  </div>
                  
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 flex gap-2.5 items-start mt-4">
                    <Info className="h-4 w-4 text-blue-600 shrink-0 mt-0.5" />
                    <p className="text-[11px] leading-relaxed text-slate-500">
                      <strong>Công thức quy đổi trọng lượng lẻ:</strong> Trọng lượng lẻ (kg) được quy ước tính dựa trên tỷ lệ con lẻ 
                      so với tổng bao chẵn (quy đổi chuẩn 25kg/bao). <br/>
                      <code>Trọng lượng = (Con Lẻ / Quy Cách) * 25kg</code>
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: TRA CỨU THÔNG MINH */}
        {activeTab === "search" && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
              <div className="max-w-3xl mx-auto space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="font-bold text-slate-900 text-base">TRA CỨU THÀNH PHẨM KHO</h3>
                    <p className="text-xs text-slate-500">Tìm kiếm nhanh vị trí lưu trữ, số lượng tồn thực tế của các loại Bulông, tán đai ốc</p>
                  </div>
                  
                  <button
                    onClick={() => openQrScanner("search")}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-50 text-blue-700 hover:bg-blue-100 active:scale-95 transition-all text-xs font-bold rounded-lg border border-blue-200"
                  >
                    <QrCode className="h-4 w-4" />
                    Quét Mã Kho QR
                  </button>
                </div>

                <div className="relative">
                  <div className="relative rounded-xl shadow-inner">
                    <div className="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3.5">
                      <Search className="h-5 w-5 text-slate-400" />
                    </div>
                    <input
                      type="text"
                      className="block w-full rounded-xl border-slate-200 pl-10 pr-4 py-3 bg-slate-50 text-sm font-semibold placeholder-slate-400 focus:border-blue-500 focus:bg-white focus:ring-blue-500"
                      placeholder="Nhập tên khóa hoặc mã kệ cần lọc, VD: H1.12, M12x30, DIN933..."
                      value={searchQuery}
                      onChange={(e) => {
                        setSearchQuery(e.target.value);
                        setShowSearchSuggestions(true);
                      }}
                      onFocus={() => setShowSearchSuggestions(true)}
                    />
                    {searchQuery && (
                      <button
                        onClick={() => {
                          setSearchQuery("");
                          setShowSearchSuggestions(false);
                        }}
                        className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-slate-400 hover:text-slate-600"
                      >
                        <X className="h-4 w-4" />
                      </button>
                    )}
                  </div>

                  {/* Autocomplete Suggestions */}
                  {showSearchSuggestions && searchSuggestions.length > 0 && (
                    <div className="absolute z-50 left-0 right-0 mt-1.5 bg-white border border-slate-200 rounded-xl shadow-xl overflow-hidden max-h-60 overflow-y-auto">
                      <div className="bg-slate-50 px-3 py-1.5 border-b border-slate-100 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                        Danh mục Gợi Ý:
                      </div>
                      {searchSuggestions.map((p, idx) => (
                        <div
                          key={`sug-${idx}`}
                          onClick={() => {
                            setSearchQuery(p.ten);
                            setShowSearchSuggestions(false);
                          }}
                          className="px-4 py-2.5 hover:bg-blue-50/50 text-xs font-semibold cursor-pointer border-b border-slate-50 text-slate-700 last:border-b-0 hover:text-blue-700 transition-colors"
                        >
                          {p.ten}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* Vùng hiển thị Kết quả Tìm kiếm */}
            {searchQuery && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                <div className="bg-slate-100 px-5 py-3 border-b border-slate-200 flex justify-between items-center">
                  <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                    Kết quả tìm thấy ({searchResults.length} bản ghi tồn)
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-800 text-white font-semibold uppercase tracking-wider text-center">
                        <th className="py-2.5 px-3 style-cell-pad max-w-[80px]">Mã Kệ</th>
                        <th className="py-2.5 px-4 text-left">Tên Thành Phẩm Hàng Hóa</th>
                        <th className="py-2.5 px-2">Số Bao</th>
                        <th className="py-2.5 px-3 text-emerald-300">Quy cách (C/B)</th>
                        <th className="py-2.5 px-2">Lẻ (con)</th>
                        <th className="py-2.5 px-3 text-cyan-300">Tổng Số Con</th>
                        <th className="py-2.5 px-3 text-yellow-300 text-right pr-6">Khối lượng (Kg)</th>
                        <th className="py-2.5 px-4 text-left">Ghi chú kệ hàng</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {searchResults.length > 0 ? (
                        searchResults.map((item) => (
                          <tr key={item.id} className="hover:bg-slate-50 transition-colors text-center font-mono">
                            <td className="py-3 px-3 font-bold text-slate-900 bg-slate-50">{item.ke}</td>
                            <td className="py-3 px-4 text-left text-slate-700 font-sans font-medium">{item.ten}</td>
                            <td className="py-3 px-2 font-semibold text-slate-600">{item.bao.toLocaleString("vi-VN")}</td>
                            <td className="py-3 px-3 font-bold text-emerald-700">{item.slBao.toLocaleString("vi-VN")}</td>
                            <td className="py-3 px-2 text-slate-500">{item.le.toLocaleString("vi-VN")}</td>
                            <td className="py-3 px-3 font-bold text-blue-700">{item.tongCon.toLocaleString("vi-VN")}</td>
                            <td className="py-3 px-3 font-bold text-rose-600 text-right pr-6">{item.khoiLuong.toFixed(2)}</td>
                            <td className="py-3 px-4 text-left font-sans text-slate-500 italic max-w-xs truncate">{item.ghiChu || "-"}</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={8} className="py-8 text-center text-slate-400 font-sans italic">
                            Không tìm thấy mặt hàng phù hợp với từ khóa tra cứu của bạn.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 3: XUẤT NHẬP KHO */}
        {activeTab === "io" && (
          <div className="space-y-6">
            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
              <div className="max-w-2xl mx-auto space-y-5">
                <div className="text-center">
                  <label className="block text-xs font-bold text-blue-700 uppercase tracking-widest mb-2.5">
                    MÃ KỆ XỬ LÝ (Ví dụ: H1.12, H2.04)
                  </label>
                  
                  <div className="flex gap-2 justify-center max-w-sm mx-auto">
                    <input
                      type="text"
                      className="block w-full rounded-xl border-blue-200 bg-blue-50/20 px-4 py-3 text-center font-bold text-blue-800 tracking-widest uppercase focus:border-blue-500 focus:bg-white text-lg placeholder-slate-300"
                      placeholder="NHẬP MÃ KỆ"
                      value={keInput}
                      onChange={(e) => setKeInput(e.target.value.toUpperCase())}
                    />
                    
                    <button
                      onClick={() => openQrScanner("io")}
                      className="bg-blue-600 text-white hover:bg-blue-700 p-3 rounded-xl transition-all cursor-pointer shadow-md shadow-blue-500/15"
                      title="Quét QR tìm mã kệ nhanh"
                    >
                      <QrCode className="h-5 w-5" />
                    </button>
                  </div>
                </div>

                {/* Danh sách các dòng mặt hàng ghi kho */}
                {ioRows.length > 0 && (
                  <div className="space-y-4 pt-4 border-t border-slate-100">
                    <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">Danh sách phiếu mặt hàng thao tác:</h4>
                    
                    {ioRows.map((row, index) => {
                      const computedWeight = getRowWeight(row.selectedProductName, row.bao, row.slBao, row.le);
                      const suggestions = getIoRowSuggestions(row.searchText);

                      return (
                        <div
                          key={row.id}
                          className={`relative border-2 rounded-xl p-4 transition-all bg-slate-50 ${
                            row.type === "NHẬP"
                              ? "border-l-8 border-emerald-500 border-slate-200"
                              : "border-l-8 border-rose-500 border-slate-200"
                          }`}
                        >
                          {/* Close / Remove Row Button */}
                          <button
                            onClick={() => setIoRows((prev) => prev.filter((item) => item.id !== row.id))}
                            className="absolute -top-2.5 -right-2.5 bg-rose-600 text-white hover:bg-rose-700 h-6 w-6 rounded-full flex items-center justify-center font-bold shadow-md cursor-pointer-item border border-white"
                          >
                            <X className="h-3 w-3" />
                          </button>

                          <div className="flex justify-between items-center mb-3">
                            <span
                              className={`text-[10px] font-bold px-2 py-0.5 rounded-full text-white ${
                                row.type === "NHẬP" ? "bg-emerald-600" : "bg-rose-600"
                              }`}
                            >
                              {row.type === "NHẬP" ? "NHẬP KHO STT" : "XUẤT KHO STT"} #{index + 1}
                            </span>
                            
                            <span className="text-xs text-slate-400 font-medium">Hàng gõ từ khóa tự động khuyên dùng</span>
                          </div>

                          {/* Autocomplete Input Search */}
                          <div className="relative mb-3">
                            <input
                              type="text"
                              className="block w-full rounded-lg border-slate-200 bg-white px-3 py-2 text-xs font-semibold placeholder-slate-400 focus:border-blue-500"
                              placeholder="Gõ tên hàng, kích thước, VD: DIN933 M12..."
                              value={row.selectedProductName || row.searchText}
                              onChange={(e) => {
                                updateIoRowValue(row.id, {
                                  searchText: e.target.value,
                                  selectedProductName: "", // invalidate old select
                                  showSuggestions: true
                                });
                              }}
                              onFocus={() => updateIoRowValue(row.id, { showSuggestions: true })}
                            />
                            
                            {/* Autocomplete Suggestions Panel */}
                            {row.showSuggestions && suggestions.length > 0 && (
                              <div className="absolute z-50 left-0 right-0 mt-1 bg-white border border-slate-200 rounded-lg shadow-lg overflow-hidden max-h-48 overflow-y-auto">
                                {suggestions.map((p, pIdx) => (
                                  <div
                                    key={`p-sug-${pIdx}`}
                                    onClick={() => {
                                      updateIoRowValue(row.id, {
                                        selectedProductName: p.ten,
                                        searchText: p.ten,
                                        showSuggestions: false
                                      });
                                    }}
                                    className="px-3 py-2 hover:bg-blue-50 text-[11px] font-semibold cursor-pointer border-b border-slate-100 last:border-b-0 text-slate-800"
                                  >
                                    {p.ten}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>

                          {/* Numeric input cells */}
                          <div className="grid grid-cols-3 gap-2 mb-3">
                            <div>
                              <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Số Bao:</label>
                              <input
                                type="number"
                                className="block w-full rounded-lg border-slate-200 px-2.5 py-1.5 text-xs text-center font-bold"
                                value={row.bao || ""}
                                placeholder="0"
                                onChange={(e) => {
                                  const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                                  updateIoRowValue(row.id, { bao: val });
                                }}
                              />
                            </div>

                            <div>
                              <label className="block text-[10px] font-bold text-blue-700 uppercase mb-1">Quy cách/Bao:</label>
                              <input
                                type="number"
                                className="block w-full rounded-lg border-slate-200 px-2.5 py-1.5 text-xs text-center font-bold text-blue-700-f font-mono"
                                value={row.slBao || ""}
                                placeholder="VD: 500"
                                onChange={(e) => {
                                  const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                                  updateIoRowValue(row.id, { slBao: val });
                                }}
                              />
                            </div>

                            <div>
                              <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Số Lẻ (con):</label>
                              <input
                                type="number"
                                className="block w-full rounded-lg border-slate-200 px-2.5 py-1.5 text-xs text-center font-semibold"
                                value={row.le || ""}
                                placeholder="0"
                                onChange={(e) => {
                                  const val = Math.max(0, parseInt(e.target.value, 10) || 0);
                                  updateIoRowValue(row.id, { le: val });
                                }}
                              />
                            </div>
                          </div>

                          {/* Ghi chú hàng */}
                          <div className="mb-2">
                            <input
                              type="text"
                              className="block w-full rounded-lg border-slate-200 px-3 py-1.5 text-xs placeholder-slate-400"
                              placeholder="Ghi chú thêm sản phẩm (không bắt buộc)..."
                              value={row.note}
                              onChange={(e) => updateIoRowValue(row.id, { note: e.target.value })}
                            />
                          </div>

                          {/* Computed metric block */}
                          <div className="flex justify-between items-center text-xs border-t border-slate-200 pt-2 text-slate-500">
                            <span>
                              Tổng số con: <strong className="font-mono text-slate-700">{(row.bao * row.slBao + row.le).toLocaleString("vi-VN")} con</strong>
                            </span>
                            <span className="font-bold text-rose-600 font-mono">
                              {computedWeight.toFixed(2)} <span className="text-[10px] font-sans font-medium text-slate-400">Kg</span>
                            </span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                {/* Bottom line action triggers */}
                <div className="flex gap-4 pt-4">
                  <button
                    onClick={() => plusIoRow("XUẤT")}
                    className="flex-1 flex justify-center items-center gap-1.5 border border-rose-200 bg-rose-50 hover:bg-rose-100/80 text-rose-700 py-3 rounded-xl text-xs font-bold transition-all shadow-sm active:scale-[0.98]"
                  >
                    <Plus className="h-4 w-4" />+ XUẤT KHO THỰC
                  </button>
                  
                  <button
                    onClick={() => plusIoRow("NHẬP")}
                    className="flex-1 flex justify-center items-center gap-1.5 border border-emerald-200 bg-emerald-50 hover:bg-emerald-100/80 text-emerald-700 py-3 rounded-xl text-xs font-bold transition-all shadow-sm active:scale-[0.98]"
                  >
                    <Plus className="h-4 w-4" />+ NHẬP KHO THÊM
                  </button>
                </div>

                {ioRows.length > 0 && (
                  <button
                    onClick={submitIoData}
                    className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 px-6 rounded-xl transition-all shadow-lg shadow-blue-600/15 text-xs uppercase tracking-wider flex justify-center items-center gap-2 mt-6"
                  >
                    <CheckCircle className="h-4.5 w-4.5" />
                    Xác nhận ghi kho thành phẩm
                  </button>
                )}
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: LỊCH SỬ GIAO DỊCH */}
        {activeTab === "history" && (
          <div className="space-y-4">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-sm no-print">
              <div>
                <h3 className="font-bold text-slate-900 text-sm uppercase">LỊCH SỬ GIAO DỊCH CHI TIẾT KHO</h3>
                <p className="text-xs text-slate-500">Nhật ký chi tiết các phiên nhập xuất kho mới nhất</p>
              </div>

              {selectedHistoryIds.length > 0 && (
                <button
                  onClick={() => triggerHistoryDelete(selectedHistoryIds)}
                  className="w-full sm:w-auto flex items-center justify-center gap-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold px-4 py-2 rounded-lg transition-all"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  XÓA PHIẾU ĐÃ CHỌN ({selectedHistoryIds.length})
                </button>
              )}
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[70vh] sticky-th-container">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-800 text-slate-100 font-semibold uppercase font-display text-center sticky top-0 z-10">
                    <tr>
                      <th className="py-2.5 px-3 style-cell-pad w-[40px] no-print">
                        <input
                          type="checkbox"
                          className="rounded text-blue-600"
                          onChange={handleSelectAllHistory}
                          checked={history.length > 0 && selectedHistoryIds.length === history.length}
                        />
                      </th>
                      <th className="py-2.5 px-3">Thời gian ghi</th>
                      <th className="py-2.5 px-2">Phân loại</th>
                      <th className="py-2.5 px-2">Vị trí Kệ</th>
                      <th className="py-2.5 px-4 text-left">Tên hàng hóa chính</th>
                      <th className="py-2.5 px-2">Bao chẵn</th>
                      <th className="py-2.5 px-3">Quy cách</th>
                      <th className="py-2.5 px-2">Con Lẻ</th>
                      <th className="py-2.5 px-3 text-cyan-300">Tổng Con</th>
                      <th className="py-2.5 px-3 text-yellow-300">Nặng kg</th>
                      <th className="py-2.5 px-4 text-left">Nội dung Ghi Chú</th>
                      <th className="py-2.5 px-3 no-print">Thao Tác</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono text-center text-slate-700">
                    {history.length > 0 ? (
                      history.map((log) => (
                        <tr key={log.id} className="hover:bg-slate-50 transition-colors">
                          <td className="py-2.5 px-3 no-print">
                            <input
                              type="checkbox"
                              className="rounded text-blue-600"
                              checked={selectedHistoryIds.includes(log.id)}
                              onChange={() => handleSelectHistoryItem(log.id)}
                            />
                          </td>
                          <td className="py-2.5 px-3 text-slate-500 font-sans text-[11px] whitespace-nowrap">{log.ngay}</td>
                          <td className="py-2.5 px-2">
                            <span
                              className={`text-[9px] font-bold uppercase px-2 py-0.5 rounded-full ${
                                log.loai === "NHẬP"
                                  ? "bg-emerald-100 text-emerald-800 border border-emerald-200"
                                  : "bg-rose-100 text-rose-800 border border-rose-200"
                              }`}
                            >
                              {log.loai}
                            </span>
                          </td>
                          <td className="py-2.5 px-2 font-bold text-slate-900 bg-slate-50">{log.ke}</td>
                          <td className="py-2.5 px-4 text-left font-sans font-medium text-slate-700">{log.tenHang}</td>
                          <td className="py-2.5 px-2">{log.bao.toLocaleString("vi-VN")}</td>
                          <td className="py-2.5 px-3 text-emerald-700 font-bold">{log.slBao.toLocaleString("vi-VN")}</td>
                          <td className="py-2.5 px-2 text-slate-400">{log.le.toLocaleString("vi-VN")}</td>
                          <td className="py-2.5 px-3 text-blue-700 font-bold">{log.tong.toLocaleString("vi-VN")}</td>
                          <td className="py-2.5 px-3 text-rose-600 font-bold">{log.kg.toFixed(2)}</td>
                          <td className="py-2.5 px-4 text-left font-sans text-slate-500 italic max-w-xs truncate">{log.ghiChu || "-"}</td>
                          <td className="py-2.5 px-3 no-print">
                            <button
                              onClick={() => triggerHistoryDelete([log.id])}
                              className="text-rose-600 hover:text-rose-800 p-1 rounded hover:bg-rose-50 transition-colors"
                              title="Xóa phiếu này"
                            >
                              <Trash2 className="h-3.5 w-3.5" />
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={12} className="py-12 text-center text-slate-400 font-sans italic">
                          Chưa ghi nhận bất kỳ dữ liệu sự kiện lịch sử gần đây nào.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 5: DỮ LIỆU TỒN KHO TỔNG */}
        {activeTab === "data" && (
          <div className="space-y-4">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-white p-4 rounded-xl border border-slate-200 shadow-sm no-print">
              <div>
                <h3 className="font-bold text-slate-900 text-sm uppercase">TỒN KHO HIỆN TẠI THÀNH PHẨM</h3>
                <p className="text-xs text-slate-500">Quản lý định lượng, sức chứa và phân bố trên từng dãy kệ hàng hóa</p>
              </div>

              <div className="w-full sm:w-auto flex flex-col sm:flex-row items-center gap-2">
                <div className="relative w-full sm:w-64">
                  <div className="pointer-events-none absolute inset-y-0 left-0 flex items-center pl-3">
                    <Search className="h-4 w-4 text-slate-400" />
                  </div>
                  <input
                    type="text"
                    className="block w-full rounded-lg border-slate-200 pl-9 pr-3 py-1.5 text-xs"
                    placeholder="Lọc hàng thô hoặc mã kệ..."
                    value={inventoryFilter}
                    onChange={(e) => setInventoryFilter(e.target.value)}
                  />
                  {inventoryFilter && (
                    <button
                      onClick={() => setInventoryFilter("")}
                      className="absolute inset-y-0 right-0 pr-2.5 flex items-center text-slate-400 hover:text-slate-600"
                    >
                      <X className="h-3.5 w-3.5" />
                    </button>
                  )}
                </div>

                <button
                  onClick={exportInventoryToExcel}
                  className="w-full sm:w-auto flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2 rounded-lg transition-all shrink-0 cursor-pointer"
                >
                  <FileSpreadsheet className="h-3.5 w-3.5" />
                  XUẤT FILE EXCEL
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[70vh] sticky-th-container">
                <table className="w-full text-xs text-left border-collapse">
                  <thead className="bg-slate-800 text-slate-100 font-semibold uppercase font-display text-center sticky top-0 z-10">
                    <tr>
                      <th className="py-2.5 px-3 style-cell-pad w-[80px]">Mã Kệ</th>
                      <th className="py-2.5 px-4 text-left">Tên Mặt Hàng Chi Tiết</th>
                      <th className="py-2.5 px-2">Bao Chẵn</th>
                      <th className="py-2.5 px-3">Quy cách/Bao</th>
                      <th className="py-2.5 px-2">Số lẻ kiểm</th>
                      <th className="py-2.5 px-3 text-cyan-300">Tổng Số Con</th>
                      <th className="py-2.5 px-3 text-yellow-300 text-right pr-6">Khối lượng (Kg)</th>
                      <th className="py-2.5 px-4 text-left">Ghi chú kệ chi tiết</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 font-mono text-center text-slate-700">
                    {filteredInventory.length > 0 ? (
                      filteredInventory.map((item) => (
                        <tr key={item.id} className="hover:bg-slate-50/50 transition-colors">
                          <td className="py-3 px-3 font-bold text-slate-900 bg-slate-50">{item.ke}</td>
                          <td className="py-3 px-4 text-left font-sans font-medium text-slate-800">{item.ten}</td>
                          <td className="py-3 px-2 font-semibold text-slate-600">{item.bao.toLocaleString("vi-VN")}</td>
                          <td className="py-3 px-3 text-emerald-700 font-bold">{item.slBao.toLocaleString("vi-VN")}</td>
                          <td className="py-3 px-2 text-slate-400">{item.le.toLocaleString("vi-VN")}</td>
                          <td className="py-3 px-3 text-blue-700 font-bold">{item.tongCon.toLocaleString("vi-VN")}</td>
                          <td className="py-3 px-3 text-rose-600 font-bold text-right pr-6">{item.khoiLuong.toFixed(2)}</td>
                          <td className="py-3 px-4 text-left font-sans text-slate-500 italic max-w-xs truncate">{item.ghiChu || "-"}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={8} className="py-12 text-center text-slate-400 font-sans italic">
                          Không tìm thấy dữ liệu tồn kho trùng khớp với từ khóa lọc.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 6: GIAO NHẬN CONT */}
        {activeTab === "cont" && (
          <div className="space-y-4">
            {/* SUBTABS PILLS NAVIGATION */}
            <div className="flex bg-slate-100 p-1.5 rounded-xl border border-slate-200 shadow-inner max-w-md mx-auto no-print">
              <button
                onClick={() => setActiveContSubTab("hist")}
                className={`flex-1 py-2 text-xs font-bold text-center rounded-lg transition-all ${
                  activeContSubTab === "hist"
                    ? "bg-blue-600 text-white shadow-md shadow-blue-600/10"
                    : "text-slate-500 hover:text-slate-800"
                }`}
              >
                Lịch sử Cont
              </button>
              
              <button
                onClick={() => setActiveContSubTab("edit")}
                className={`flex-1 py-2 text-xs font-bold text-center rounded-lg transition-all ${
                  activeContSubTab === "edit"
                    ? "bg-blue-600 text-white shadow-md shadow-blue-600/10"
                    : "text-slate-500 hover:text-slate-800"
                }`}
              >
                Chỉnh sửa Cont
              </button>

              <button
                onClick={() => setActiveContSubTab("import")}
                className={`flex-1 py-2 text-xs font-bold text-center rounded-lg transition-all ${
                  activeContSubTab === "import"
                    ? "bg-blue-600 text-white shadow-md shadow-blue-600/10"
                    : "text-slate-500 hover:text-slate-800"
                }`}
              >
                Nhập Cont Excel
              </button>
            </div>

            {/* CONT SUB-TAB 1: DANH SÁCH LỊCH SỬ CONT */}
            {activeContSubTab === "hist" && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
                <div className="flex flex-col md:flex-row gap-3 items-center justify-between no-print">
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-2 w-full md:flex-1">
                    <select
                      className="form-select rounded-lg border-slate-200 text-xs font-semibold py-2 bg-slate-50/50 w-full"
                      value={selectedContHistName}
                      onChange={(e) => {
                        const val = e.target.value;
                        setSelectedContHistName(val);
                        setAppliedContHistName(val);
                      }}
                    >
                      <option value="">-- Chọn tên Cont --</option>
                      {uniqueContNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>

                    <input
                      type="text"
                      className="form-input rounded-lg border-slate-200 text-xs py-2 bg-slate-50/50 w-full"
                      placeholder="Tìm từ khóa nhanh..."
                      value={findHistInput}
                      onChange={(e) => setFindHistInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") handleSearchHist();
                      }}
                    />

                    <button
                      onClick={handleSearchHist}
                      className="flex items-center justify-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-all shadow-md shadow-blue-500/15 cursor-pointer w-full"
                    >
                      <Search className="h-3.5 w-3.5" />
                      TÌM KIẾM
                    </button>

                    <button
                      onClick={handleResetHist}
                      className="flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg transition-all border border-slate-200 cursor-pointer w-full"
                      title="Xóa bộ lọc"
                    >
                      <X className="h-3.5 w-3.5" />
                      XÓA LỌC
                    </button>
                  </div>

                  <div className="flex gap-2 w-full md:w-auto">
                    <button
                      onClick={() => printContTable(filteredContHistList, selectedContHistName)}
                      className="flex-1 md:flex-none flex items-center justify-center gap-1 px-3.5 py-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 text-xs font-bold rounded-lg transition-all"
                    >
                      <Printer className="h-4 w-4" />
                      IN BẢNG CONT
                    </button>
                    
                    <button
                      onClick={() => exportContListToExcel(filteredContHistList, selectedContHistName)}
                      className="flex-1 md:flex-none flex items-center justify-center gap-1 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-all shadow shadow-emerald-600/10"
                    >
                      <FileSpreadsheet className="h-4 w-4" />
                      XUẤT EXCEL
                    </button>
                  </div>
                </div>

                {/* Printable container list */}
                <div className="overflow-x-auto max-h-[60vh] sticky-th-container">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-slate-800 text-slate-100 font-semibold uppercase font-display text-center sticky top-0 z-10">
                      <tr>
                        <th className="py-2.5 px-3 style-cell-pad w-[45px]">STT</th>
                        <th className="py-2.5 px-3">Grade</th>
                        <th className="py-2.5 px-3">Standard</th>
                        <th className="py-2.5 px-4 text-left">Size</th>
                        <th className="py-2.5 px-2">Finish</th>
                        <th className="py-2.5 px-2">No. Bàn Giao</th>
                        <th className="py-2.5 px-3 text-cyan-300">Tổng Số Bao Nguyên</th>
                        <th className="py-2.5 px-3 text-emerald-300">Con / Bao</th>
                        <th className="py-2.5 px-2 text-yellow-300">Con lẻ</th>
                        <th className="py-2.5 px-4 text-left">Nội dung ghi chú</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono text-center text-slate-700">
                      {filteredContHistList.length > 0 ? (
                        filteredContHistList.map((item, idx) => {
                          let qtyParts: string[] = [];
                          if (item.bao > 0) {
                            qtyParts.push(`${item.bao}`);
                            if (item.slBao > 0) {
                              qtyParts.push(`${item.slBao}`);
                            }
                          }
                          if (item.le > 0) {
                            qtyParts.push(`${item.le}`);
                          }
                          const qtyString = qtyParts.join(" x ") || "-";

                          return (
                            <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                              <td className="py-2.5 px-3 font-bold text-slate-400 bg-slate-50">{idx + 1}</td>
                              <td className="py-2.5 px-3 font-semibold text-slate-800 font-sans">{item.grade}</td>
                              <td className="py-2.5 px-3 text-slate-500">{item.standard}</td>
                              <td className="py-2.5 px-4 text-left font-bold text-slate-900">{item.size}</td>
                              <td className="py-2.5 px-2 font-medium font-sans">{item.finish}</td>
                              <td className="py-2.5 px-2 font-bold text-indigo-700">{item.no}</td>
                              <td className="py-2.5 px-3 font-bold text-slate-900">{item.bao.toLocaleString("vi-VN")} bao</td>
                              <td className="py-2.5 px-3 text-emerald-700 font-bold">{item.slBao.toLocaleString("vi-VN")}</td>
                              <td className="py-2.5 px-2 text-slate-500">{item.le.toLocaleString("vi-VN")}</td>
                              <td className="py-2.5 px-4 text-left font-sans text-slate-400 italic font-medium">{item.ghiChu || "-"}</td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={10} className="py-12 text-center text-slate-400 font-sans italic">
                            Không tìm thấy chi tiết thông tin đợt Container phù hợp với tìm kiếm của bạn.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* CONT SUB-TAB 2: CHỈNH SỬA & IN NHÃN CONT */}
            {activeContSubTab === "edit" && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
                <div className="flex flex-col md:flex-row gap-3 items-center justify-between no-print">
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-2 w-full md:flex-1">
                    <select
                      className="form-select rounded-lg border-slate-200 text-xs font-semibold py-2 bg-slate-50/50 w-full"
                      value={selectedContEditName}
                      onChange={(e) => {
                        const val = e.target.value;
                        setSelectedContEditName(val);
                        setAppliedContEditName(val);
                      }}
                    >
                      <option value="">-- Chọn tên Cont --</option>
                      {uniqueContNames.map((name) => (
                        <option key={name} value={name}>
                          {name}
                        </option>
                      ))}
                    </select>

                    <input
                      type="text"
                      className="form-input rounded-lg border-slate-200 text-xs py-2 bg-slate-50/50 w-full"
                      placeholder="Tìm từ khóa nhanh..."
                      value={findContInput}
                      onChange={(e) => setFindContInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") handleSearchEdit();
                      }}
                    />

                    <button
                      onClick={handleSearchEdit}
                      className="flex items-center justify-center gap-1.5 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-all shadow-md shadow-blue-500/15 cursor-pointer w-full"
                    >
                      <Search className="h-3.5 w-3.5" />
                      TÌM KIẾM
                    </button>

                    <button
                      onClick={handleResetEdit}
                      className="flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg transition-all border border-slate-200 cursor-pointer w-full"
                      title="Xóa bộ lọc"
                    >
                      <X className="h-3.5 w-3.5" />
                      XÓA LỌC
                    </button>
                  </div>

                  <button
                    onClick={() => printContTable(filteredContEditList, selectedContEditName)}
                    className="w-full md:w-auto flex items-center justify-center gap-1 px-3.5 py-1.5 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 text-xs font-bold rounded-lg transition-all"
                  >
                    <Printer className="h-4 w-4" />
                    IN BẢNG CONT
                  </button>
                </div>

                <div className="overflow-x-auto max-h-[60vh] sticky-th-container">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-slate-800 text-slate-100 font-semibold uppercase font-display text-center sticky top-0 z-10">
                      <tr>
                        <th className="py-2.5 px-3 style-cell-pad w-[45px]">STT</th>
                        <th className="py-2.5 px-3">Grade</th>
                        <th className="py-2.5 px-3">Standard</th>
                        <th className="py-2.5 px-4 text-left">Size</th>
                        <th className="py-2.5 px-2">Finish</th>
                        <th className="py-2.5 px-2">No. Bàn giao</th>
                        <th className="py-2.5 px-2 text-cyan-300">Số Bao</th>
                        <th className="py-2.5 px-3 text-emerald-300">Con / Bao</th>
                        <th className="py-2.5 px-2 text-yellow-300">Con lẻ</th>
                        <th className="py-2.5 px-4 text-left">Ghi chú</th>
                        <th className="py-2.5 px-3 no-print">Thao Tác</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 font-mono text-center text-slate-700">
                      {filteredContEditList.length > 0 ? (
                        filteredContEditList.map((item, idx) => {
                          const isEditing = !!editingContRows[item.id];
                          const editData = temporaryContRowData[item.id];

                          return (
                            <tr key={item.id} className="hover:bg-slate-50 transition-colors">
                              <td className="py-2.5 px-3 font-bold text-slate-400 bg-slate-50">{idx + 1}</td>
                              
                              {/* Grade Cell */}
                              <td className="py-2.5 px-3 font-sans">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs font-bold bg-white"
                                    value={editData?.grade}
                                    onChange={(e) => updateTemporaryContField(item.id, "grade", e.target.value)}
                                  />
                                ) : (
                                  <span
                                    onClick={() => toggleEditContRow(item.id, item)}
                                    className="cursor-pointer border-b border-dashed border-slate-300 hover:text-blue-600 block px-1 py-0.5 rounded"
                                    title="Nhấp để sửa nhanh"
                                  >
                                    {item.grade || "-"}
                                  </span>
                                )}
                              </td>

                              {/* Standard Cell */}
                              <td className="py-2.5 px-3 text-slate-500">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs bg-white"
                                    value={editData?.standard}
                                    onChange={(e) => updateTemporaryContField(item.id, "standard", e.target.value)}
                                  />
                                ) : (
                                  <span
                                    onClick={() => toggleEditContRow(item.id, item)}
                                    className="cursor-pointer border-b border-dashed border-slate-300 hover:text-blue-600 block px-1 py-0.5 rounded"
                                  >
                                    {item.standard || "-"}
                                  </span>
                                )}
                              </td>

                              {/* Size Cell */}
                              <td className="py-2.5 px-4 text-left font-bold text-slate-900">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-xs font-bold bg-white"
                                    value={editData?.size}
                                    onChange={(e) => updateTemporaryContField(item.id, "size", e.target.value)}
                                  />
                                ) : (
                                  <span
                                    onClick={() => toggleEditContRow(item.id, item)}
                                    className="cursor-pointer border-b border-dashed border-slate-300 hover:text-blue-600 block px-1 py-0.5 rounded font-mono"
                                  >
                                    {item.size || "-"}
                                  </span>
                                )}
                              </td>

                              {/* Finish Cell */}
                              <td className="py-2.5 px-2 font-medium font-sans">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs bg-white"
                                    value={editData?.finish}
                                    onChange={(e) => updateTemporaryContField(item.id, "finish", e.target.value)}
                                  />
                                ) : (
                                  <span
                                    onClick={() => toggleEditContRow(item.id, item)}
                                    className="cursor-pointer border-b border-dashed border-slate-300 hover:text-blue-600 block px-1 py-0.5 rounded"
                                  >
                                    {item.finish || "-"}
                                  </span>
                                )}
                              </td>

                              {/* NO Cell */}
                              <td className="py-2.5 px-2 font-bold text-indigo-700">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs font-bold bg-white text-indigo-700"
                                    value={editData?.no}
                                    onChange={(e) => updateTemporaryContField(item.id, "no", e.target.value)}
                                  />
                                ) : (
                                  <span
                                    onClick={() => toggleEditContRow(item.id, item)}
                                    className="cursor-pointer border-b border-dashed border-slate-300 hover:text-blue-600 block px-1 py-0.5 rounded"
                                  >
                                    {item.no || "-"}
                                  </span>
                                )}
                              </td>

                              {/* Số bao Cell */}
                              <td className="py-2.5 px-2">
                                {isEditing ? (
                                  <input
                                    type="number"
                                    className="block w-[60px] mx-auto border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs bg-white font-bold"
                                    value={editData?.bao}
                                    onChange={(e) => updateTemporaryContField(item.id, "bao", parseInt(e.target.value, 10) || 0)}
                                  />
                                ) : (
                                  <span>{item.bao.toLocaleString("vi-VN")}</span>
                                )}
                              </td>

                              {/* Con / Bao Cell */}
                              <td className="py-2.5 px-3 text-emerald-700 font-bold">
                                {isEditing ? (
                                  <input
                                    type="number"
                                    className="block w-[70px] mx-auto border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs bg-white font-mono"
                                    value={editData?.slBao}
                                    onChange={(e) => updateTemporaryContField(item.id, "slBao", parseInt(e.target.value, 10) || 0)}
                                  />
                                ) : (
                                  <span>{item.slBao.toLocaleString("vi-VN")}</span>
                                )}
                              </td>

                              {/* Lẻ Cell */}
                              <td className="py-2.5 px-2 font-medium">
                                {isEditing ? (
                                  <input
                                    type="number"
                                    className="block w-[60px] mx-auto border border-blue-500 rounded px-1.5 py-0.5 text-center text-xs bg-white text-rose-700 font-bold"
                                    value={editData?.le}
                                    onChange={(e) => updateTemporaryContField(item.id, "le", parseInt(e.target.value, 10) || 0)}
                                  />
                                ) : (
                                  <span>{item.le.toLocaleString("vi-VN")}</span>
                                )}
                              </td>

                              {/* Note Cell */}
                              <td className="py-2.5 px-4 text-left font-sans text-slate-500">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    className="block w-full border border-blue-500 rounded px-1.5 py-0.5 text-xs bg-white"
                                    value={editData?.ghiChu}
                                    onChange={(e) => updateTemporaryContField(item.id, "ghiChu", e.target.value)}
                                  />
                                ) : (
                                  <span>{item.ghiChu || "-"}</span>
                                )}
                              </td>

                              {/* Print Single Label */}
                              <td className="py-2.5 px-3 no-print">
                                <div className="flex gap-2 justify-center">
                                  {isEditing ? (
                                    <button
                                      onClick={() => saveContRow(item.id)}
                                      className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-2 py-1 rounded text-[10px] uppercase cursor-pointer"
                                    >
                                      Lưu
                                    </button>
                                  ) : (
                                    <button
                                      onClick={() => toggleEditContRow(item.id, item)}
                                      className="bg-blue-100 hover:bg-blue-200 text-blue-800 px-2 py-1 rounded text-[10px] font-bold cursor-pointer"
                                    >
                                      Sửa
                                    </button>
                                  )}

                                  <button
                                    onClick={() => printSingleLabel(item, isEditing)}
                                    className="bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 px-2.5 py-1 rounded text-[10px] font-bold flex items-center gap-0.5 cursor-pointer"
                                    title="In nhãn dán bao rời"
                                  >
                                    <Printer className="h-3 w-3" /> In nhãn
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={11} className="py-12 text-center text-slate-400 font-sans italic">
                            Hãy lọc một container cụ thể ở trên để thực hiện chỉnh sửa, in tem dán decal.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* CONT SUB-TAB 3: NHẬP FILE EXCEL */}
            {activeContSubTab === "import" && (
              <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden max-w-xl mx-auto p-6 space-y-5">
                <div className="flex items-center gap-3 border-b border-slate-100 pb-4 mb-4">
                  <div className="bg-emerald-50 text-emerald-600 p-2.5 rounded-xl">
                    <Upload className="h-5 w-5" />
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-900 text-base">TẢI TRỰC TIẾP FILE EXCEL CONT</h3>
                    <p className="text-xs text-slate-500">Phân tích bảng dỡ Container, làm sạch mã cỡ chuẩn hóa của CTEG</p>
                  </div>
                </div>

                <div className="space-y-4">
                  <div className="border-2 border-dashed border-slate-300 hover:border-blue-500 rounded-xl p-8 text-center bg-slate-50 transition-colors pointer-event-area relative">
                    <input
                      type="file"
                      className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                      accept=".xlsx, .xls"
                      onChange={handleExcelImport}
                    />
                    
                    <FileSpreadsheet className="h-10 w-10 text-slate-400 mx-auto mb-3.5" />
                    <p className="text-xs font-bold text-slate-700">Kéo thả file Excel của bạn vào đây</p>
                    <p className="text-[10px] text-slate-400 mt-1">hoặc click để duyệt qua ổ cứng của bạn (Chọn file .xlsx hoặc .xls)</p>
                  </div>

                  <div className="bg-slate-50 rounded-xl p-4 border border-slate-200 flex gap-3 text-slate-600 start-items-gap">
                    <Info className="h-4.5 w-4.5 text-blue-600 shrink-0 mt-0.5 animate-bounce" />
                    
                    <div className="space-y-1.5 font-sans">
                      <p className="text-xs font-bold text-slate-800">Quy chuẩn dữ liệu đóng tệp:</p>
                      
                      <ul className="text-[10px] list-disc list-inside space-y-1 text-slate-500 leading-relaxed font-medium">
                        <li>Cột thứ 8 (index 7 - H) quy định là Tên Container duy nhất.</li>
                        <li>Trường Bề mặt (Column 6 - Finish): Tự động chuyển đổi &quot;BZP&quot; thành &quot;Xi&quot;, &quot;BLACK&quot; thành &quot;Đen&quot;.</li>
                        <li>Trường Kích Cỡ (Column 5 - Size): Tự động khử chữ M (Ví dụ: &quot;M12&quot; thành &quot;12&quot;) và chuẩn hóa dấu nhân thành &quot;x&quot; (Ví dụ: &quot;12*30&quot; thành &quot;12x30&quot;).</li>
                      </ul>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>



      {/* SYSTEM GENERAL MODAL - CAMERA ACTIVE QR SCANNER */}
      {qrModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/90 flex flex-col justify-center items-center p-4">
          <div className="w-full max-w-md text-center max-h-screen flex flex-col justify-between py-6">
            <div className="px-4 text-white">
              <h4 className="text-base font-bold uppercase tracking-wider mb-1">KIỂM QUÉT MÃ QR CODE KHO</h4>
              <p className="text-xs text-white/60 mb-4">{qrStatusText}</p>
            </div>

            {/* Scanner Container Box */}
            <div className="flex-1 w-full max-w-[360px] mx-auto bg-slate-900 border-2 border-blue-500 rounded-2xl overflow-hidden shadow-2xl relative flex items-center justify-center aspect-square">
              <div id="qr-reader-target" className="absolute select-none inset-0 w-full h-full"></div>
              
              {/* Overlay laser target scanning indicator line in css */}
              <div className="absolute w-[80%] h-0.5 bg-blue-500 shadow-[0_0_10px_rgb(59,130,246,0.8)] animate-[bounce_3s_infinite] top-[10%]"></div>
            </div>

            <div className="mt-6 px-4 space-y-4">
              {/* Fallback File Select Button for restricted secure camera viewports */}
              {qrFallbackActive && (
                <div className="relative text-center">
                  <p className="text-[11px] text-yellow-300 font-medium leading-relaxed max-w-[280px] mx-auto mb-2.5">
                    ⚠️ API Camera nhúng bị chặn do cấu trúc iFrame an toàn của Safari/Chrome nhúng. Hãy chụp ảnh hoặc chọn ảnh mã QR!
                  </p>
                  
                  <div className="relative inline-block overflow-hidden bg-yellow-400 hover:bg-yellow-500 text-slate-950 font-bold px-5 py-2.5 rounded-full text-xs cursor-pointer tracking-wider shadow">
                    <span>📸 TẢI ẢNH CHỤP QUÉT QR</span>
                    <input
                      type="file"
                      className="absolute inset-0 opacity-0 cursor-pointer h-full w-full"
                      accept="image/*"
                      capture="environment"
                      onChange={handleQrFileInput}
                    />
                  </div>
                </div>
              )}

              <button
                onClick={closeAndCleanupScanner}
                className="w-40 mx-auto bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 px-6 rounded-full text-xs uppercase cursor-pointer"
              >
                HỦY BỎ QUÉT
              </button>
            </div>
          </div>
        </div>
      )}

      {/* PASSWORD SECURE MODAL - DELETE HISTORY LOOPS */}
      {pwModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 p-6 max-w-sm w-full shadow-2xl space-y-4">
            <div className="flex items-center gap-3 border-b border-slate-100 pb-3">
              <div className="bg-rose-50 text-rose-600 p-2 rounded-lg shrink-0">
                <AlertTriangle className="h-5 w-5" />
              </div>
              
              <div>
                <h4 className="font-bold text-slate-900 text-sm">XÁC THỰC BẢO MẬT HỆ THỐNG</h4>
                <p className="text-[11px] text-slate-500">Hành động này sẽ xóa vĩnh viễn dòng lịch sử, mất kiểm soát dữ liệu.</p>
              </div>
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">
                  Nhập Mật Khẩu Thủ Kho:
                </label>
                
                <input
                  type="password"
                  className="block w-full rounded-xl border-slate-200 px-3.5 py-2.5 text-center font-bold tracking-widest text-slate-800 focus:border-rose-500 placeholder-slate-300"
                  placeholder="••••••••"
                  value={pwInput}
                  onChange={(e) => setPwInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") confirmDeleteHistoryWithPassword();
                  }}
                />
                
                <span className="block text-[9px] text-slate-400 mt-1">
                  * Nhập &quot;admin&quot; hoặc &quot;123456&quot; để bỏ qua kiểm tra quyền.
                </span>
              </div>

              <div className="flex gap-3 pt-2">
                <button
                  onClick={() => setPwModalOpen(false)}
                  className="flex-1 border border-slate-200 hover:bg-slate-50 text-slate-700 py-2.5 rounded-lg text-xs font-bold"
                >
                  Bỏ qua
                </button>
                
                <button
                  onClick={confirmDeleteHistoryWithPassword}
                  className="flex-1 bg-rose-600 hover:bg-rose-700 text-white py-2.5 rounded-lg text-xs font-bold transition-all shadow"
                >
                  XÁC NHẬN XÓA
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* GOOGLE SHEETS CONFIGURATION MODAL */}
      {isConfigModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex justify-center items-center p-4">
          <div className="bg-white rounded-2xl border border-slate-200 p-6 max-w-lg w-full max-h-[90vh] overflow-y-auto shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2 text-blue-700">
                <FileSpreadsheet className="h-5 w-5" />
                <h4 className="font-bold text-slate-900 text-sm uppercase">CẤU HÌNH LIÊN KẾT GOOGLE SHEETS</h4>
              </div>
              <button
                onClick={() => {
                  setIsConfigModalOpen(false);
                  setTestResult(null);
                }}
                className="text-slate-400 hover:text-slate-600 rounded-lg p-1 hover:bg-slate-100 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="space-y-4 text-xs font-sans">
              <p className="text-[11px] text-slate-500 leading-relaxed">
                Đồng bộ hóa hoạt động kho hàng trực tiếp với trang tính Google của bạn.
                Mọi hành động nhập, xuất, đóng Cont sẽ lưu lại thời gian thực.
              </p>

              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">
                  Phương thức kết nối:
                </label>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setGsConfig(prev => ({ ...prev, connectionType: "local" }));
                      setTestResult(null);
                    }}
                    className={`p-2 rounded-xl border text-center font-semibold cursor-pointer transition-all ${
                      gsConfig.connectionType === "local"
                        ? "bg-slate-900 border-slate-900 text-white"
                        : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    Local Offline
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setGsConfig(prev => ({ ...prev, connectionType: "appscript" }));
                      setTestResult(null);
                    }}
                    className={`p-2 rounded-xl border text-center font-semibold cursor-pointer transition-all ${
                      gsConfig.connectionType === "appscript"
                        ? "bg-blue-600 border-blue-600 text-white"
                        : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    Apps Script URL
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setGsConfig(prev => ({ ...prev, connectionType: "oauth" }));
                      setTestResult(null);
                    }}
                    className={`p-2 rounded-xl border text-center font-semibold cursor-pointer transition-all ${
                      gsConfig.connectionType === "oauth"
                        ? "bg-cyan-600 border-cyan-600 text-white"
                        : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    Google OAuth API
                  </button>
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">
                  Google Sheet Spreadsheet ID:
                </label>
                <input
                  type="text"
                  className="block w-full rounded-xl border border-slate-200 px-3.5 py-2.5 font-bold text-slate-800 placeholder-slate-300"
                  placeholder="ID Trang tính..."
                  value={gsConfig.spreadsheetId}
                  onChange={(e) => {
                    setGsConfig(prev => ({ ...prev, spreadsheetId: e.target.value }));
                    setTestResult(null);
                  }}
                />
                <span className="block text-[9px] text-slate-400 mt-1">
                  * Nhập ID từ đường dẫn Google Sheets của bạn. ID là chuỗi ký tự dài nằm giữa d/ và /edit.
                </span>
              </div>

              {gsConfig.connectionType === "appscript" && (
                <div className="space-y-3">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">
                      Google Apps Script Web App URL:
                    </label>
                    <input
                      type="text"
                      className="block w-full rounded-xl border border-slate-200 px-3.5 py-2.5 font-semibold text-slate-800 placeholder-slate-300"
                      placeholder="https://script.google.com/macros/s/.../exec"
                      value={gsConfig.webAppUrl}
                      onChange={(e) => {
                        setGsConfig(prev => ({ ...prev, webAppUrl: e.target.value }));
                        setTestResult(null);
                      }}
                    />
                    <span className="block text-[9px] text-slate-400 mt-1">
                      * Dán URL Web App nhận được sau khi Triển khai dưới dạng Ứng dụng Web trong Google Sheets Apps Script.
                    </span>
                  </div>

                  {/* COLLAPSIBLE GUIDE FOR APPS SCRIPT */}
                  <div className="border border-slate-100 rounded-xl overflow-hidden bg-slate-50">
                    <button
                      type="button"
                      onClick={() => setShowScriptGuide(!showScriptGuide)}
                      className="w-full flex justify-between items-center px-4 py-3 bg-slate-100/60 font-bold text-slate-700 hover:bg-slate-100 cursor-pointer"
                    >
                      <span>{showScriptGuide ? "🔽 ẨN HƯỚNG DẪN CÀI ĐẶT" : "▶️ XEM HƯỚNG DẪN CÀI ĐẶT & CODE SCRIPT"}</span>
                      <FileSpreadsheet className="h-4 w-4 text-blue-500" />
                    </button>
                    {showScriptGuide && (
                      <div className="p-4 space-y-3 border-t border-slate-100 bg-white max-h-[300px] overflow-y-auto">
                        <div className="space-y-1.5 text-slate-600 text-[11px] leading-relaxed">
                          <p><strong>Bước 1:</strong> Mở Google Sheets của bạn, chuẩn bị 4 Sheets mang tên: <strong>Tổng hợp</strong>, <strong>Lịch sử</strong>, <strong>Cont</strong>, <strong>Mã hàng</strong>.</p>
                          <p><strong>Bước 2:</strong> Nhấp menu <strong>Tiện ích mở rộng (Extensions)</strong> &gt; <strong>Apps Script</strong>.</p>
                          <p><strong>Bước 3:</strong> Xoá sạch code cũ, sao chép toàn bộ mã nguồn bên dưới và dán vào.</p>
                          <p><strong>Bước 4:</strong> Cập nhật hằng số <code>SPREADSHEET_ID_DEFAULT</code> trong code bằng ID Trang tính của bạn.</p>
                          <p><strong>Bước 5:</strong> Nhấp <strong>Triển khai (Deploy)</strong> &gt; <strong>Triển khai mới (New Deployment)</strong>. Chọn loại ứng dụng là <strong>Web App (Ứng dụng Web)</strong>.</p>
                          <p><strong>Bước 6:</strong> Phần &quot;Ai có quyền truy cập&quot; (Who has access), chọn <strong>Bất kỳ ai (Anyone)</strong>. Tiến hành cấp quyền (Authorize) &amp; Copy URL của Web App dán vào ô cấu hình ở trên.</p>
                        </div>

                        <div className="space-y-1 bg-slate-900 text-slate-100 p-2.5 rounded-lg relative font-mono text-[9px] overflow-x-auto leading-relaxed max-h-[160px]">
                          <button
                            type="button"
                            onClick={() => {
                              const codeStr = `// GOOGLE APPS SCRIPT CODE - HỆ THỐNG KHO CTEG
var SPREADSHEET_ID_DEFAULT = "1QKueqrlHhA0xhW1ZnkE4QrpKLUVD6eDTRNYxUMBtbhc";

function doGet(e) {
  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_DEFAULT);
  var values = [];
  
  if (action === "test") {
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Kết nối thành công!" }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getProductList") {
    var s = ss.getSheetByName("Mã hàng");
    if (s && s.getLastRow() >= 2) values = s.getRange(2, 2, s.getLastRow() - 1, 3).getValues();
    return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getRealTimeInventory") {
    var s = ss.getSheetByName("Tổng hợp");
    if (s && s.getLastRow() >= 2) values = s.getRange(2, 1, s.getLastRow() - 1, 9).getValues();
    return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getHistory") {
    var s = ss.getSheetByName("Lịch sử");
    if (s && s.getLastRow() >= 2) values = s.getRange(1, 1, s.getLastRow(), 11).getValues();
    return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getContList") {
    var s = ss.getSheetByName("Cont");
    if (s && s.getLastRow() >= 2) values = s.getRange(1, 1, s.getLastRow(), 12).getValues();
    return ContentService.createTextOutput(JSON.stringify(values)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID_DEFAULT);
    
    if (action === "processMultiAdjust") {
      var ke = postData.ke || "";
      var items = postData.items || [];
      
      var sMa = ss.getSheetByName("Mã hàng");
      var ttMap = {};
      if (sMa && sMa.getLastRow() >= 2) {
        var pRaw = sMa.getRange(2, 2, sMa.getLastRow() - 1, 3).getValues();
        pRaw.forEach(function(row) {
          var name = String(row[0]).trim().toUpperCase();
          var tyTrong = parseFloat(String(row[2]).replace(/,/g, ".")) || 0;
          ttMap[name] = tyTrong;
        });
      }
      
      var sTh = ss.getSheetByName("Tổng hợp");
      var thData = [];
      if (sTh && sTh.getLastRow() >= 2) thData = sTh.getRange(2, 1, sTh.getLastRow() - 1, 9).getValues();
      var thMap = thData.map(function(row) {
        return [
          row[0] || "", row[1] || "", row[2] || "",
          parseFloat(row[3]) || 0, parseFloat(row[4]) || 0, parseFloat(row[5]) || 0,
          parseFloat(row[6]) || 0, parseFloat(row[7]) || 0, String(row[8] || "")
        ];
      });
      
      for (var i = 0; i < items.length; i++) {
        var item = items[i];
        var isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
        if (!isNhap) {
          var ten = String(item.tenHang).trim();
          var b = parseFloat(item.bao) || 0;
          var cb = parseFloat(item.slBao) || 0;
          var l = parseFloat(item.le) || 0;
          var note = String(item.ghiChu || "").trim();
          var xuatTotalCon = b * cb + l;
          
          var foundIdx = -1;
          for (var j = 0; j < thMap.length; j++) {
            if (
              String(thMap[j][0]).toUpperCase() === ke.toUpperCase() &&
              String(thMap[j][2]).toUpperCase() === ten.toUpperCase() &&
              thMap[j][4] === cb &&
              String(thMap[j][8]).toUpperCase() === note.toUpperCase()
            ) {
              foundIdx = j;
              break;
            }
          }
          if (foundIdx === -1) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "Không tìm thấy hàng '" + ten + "' quy cách " + cb + " tại kệ " + ke + " để xuất!" })).setMimeType(ContentService.MimeType.JSON);
          }
          var currentTotal = thMap[foundIdx][6];
          if (currentTotal < xuatTotalCon) {
            return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "Kệ " + ke + " không đủ hàng cho mã '" + ten + "' quy cách " + cb + "!" })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      
      var sHist = ss.getSheetByName("Lịch sử");
      var now = new Date();
      var timestamp = now.toLocaleDateString("vi-VN") + " " + now.toLocaleTimeString("vi-VN");
      
      items.forEach(function(item) {
        var ten = String(item.tenHang).trim();
        var b = parseFloat(item.bao) || 0;
        var cb = parseFloat(item.slBao) || 0;
        var l = parseFloat(item.le) || 0;
        var note = String(item.ghiChu || "").trim();
        var isNhap = item.type.toUpperCase() === "NHẬP" || item.type.toUpperCase() === "NHAP";
        
        var totalCon = b * cb + l;
        var tyTrong = ttMap[ten.toUpperCase()] || 0;
        var kg = (totalCon * tyTrong) / 1000;
        var histId = "hist-" + now.getTime() + "-" + Math.floor(Math.random() * 10000);
        
        if (sHist) {
          sHist.appendRow([timestamp, isNhap ? "NHẬP" : "XUẤT", ten, ke.toUpperCase(), b, cb, l, kg.toFixed(2) + " Kg", note, histId, totalCon]);
        }
        
        var foundIdx = -1;
        for (var j = 0; j < thMap.length; j++) {
          if (
            String(thMap[j][0]).toUpperCase() === ke.toUpperCase() &&
            String(thMap[j][2]).toUpperCase() === ten.toUpperCase() &&
            thMap[j][4] === cb &&
            String(thMap[j][8]).toUpperCase() === note.toUpperCase()
          ) {
            foundIdx = j;
            break;
          }
        }
        if (foundIdx !== -1) {
          var currBao = thMap[foundIdx][3];
          var currLe = thMap[foundIdx][5];
          var finalBao = isNhap ? currBao + b : currBao - b;
          var finalLe = isNhap ? currLe + l : currLe - l;
          if (finalLe < 0) {
            finalBao -= 1;
            finalLe += cb;
          }
          var finalTotal = finalBao * cb + finalLe;
          thMap[foundIdx][3] = finalBao;
          thMap[foundIdx][5] = finalLe;
          thMap[foundIdx][6] = finalTotal;
          thMap[foundIdx][7] = (finalTotal * tyTrong) / 1000;
        } else {
          if (isNhap) {
            thMap.push([ke.toUpperCase(), "", ten, b, cb, l, totalCon, kg, note]);
          }
        }
      });
      
      var finalTh = thMap.filter(function(r) { return r[6] > 0; });
      if (sTh) {
        if (sTh.getLastRow() >= 2) sTh.getRange(2, 1, sTh.getLastRow() - 1, 9).clearContent();
        if (finalTh.length > 0) sTh.getRange(2, 1, finalTh.length, 9).setValues(finalTh);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "deleteSelectedHistory") {
      var ids = postData.ids || [];
      var password = postData.password;
      if (password !== "Tanlh") {
        return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "Sai mật khẩu!" })).setMimeType(ContentService.MimeType.JSON);
      }
      var sHist = ss.getSheetByName("Lịch sử");
      if (!sHist || sHist.getLastRow() < 2) return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
      
      var histRows = sHist.getRange(1, 1, sHist.getLastRow(), 11).getValues();
      var remaining = [histRows[0]];
      
      var sTh = ss.getSheetByName("Tổng hợp");
      var thData = [];
      if (sTh && sTh.getLastRow() >= 2) thData = sTh.getRange(2, 1, sTh.getLastRow() - 1, 9).getValues();
      var thMap = thData.map(function(row) {
        return [
          row[0]||"", row[1]||"", row[2]||"",
          parseFloat(row[3])||0, parseFloat(row[4])||0, parseFloat(row[5])||0,
          parseFloat(row[6])||0, parseFloat(row[7])||0, String(row[8]||"")
        ];
      });
      
      var sMa = ss.getSheetByName("Mã hàng");
      var ttMap = {};
      if (sMa && sMa.getLastRow() >= 2) {
        sMa.getRange(2, 2, sMa.getLastRow() - 1, 3).getValues().forEach(function(r) {
          ttMap[String(r[0]).trim().toUpperCase()] = parseFloat(String(r[2]).replace(/,/g, ".")) || 0;
        });
      }
      
      for (var i = 1; i < histRows.length; i++) {
        var row = histRows[i];
        if (ids.indexOf(String(row[9])) !== -1) {
          var type = String(row[1]).toUpperCase().trim();
          var ke = String(row[3]).trim().toUpperCase();
          var tenHang = String(row[2]).trim().toUpperCase();
          var slRevert = parseFloat(row[10]) || 0;
          var cb = parseFloat(row[5]) || 0;
          var note = String(row[8] || "").trim().toUpperCase();
          
          for (var j = 0; j < thMap.length; j++) {
            if (
              String(thMap[j][0]).toUpperCase() === ke &&
              String(thMap[j][2]).toUpperCase() === tenHang &&
              thMap[j][4] === cb &&
              String(thMap[j][8]).toUpperCase() === note
            ) {
              var curr = thMap[j][6];
              var next = (type === "NHẬP" || type === "NHAP") ? curr - slRevert : curr + slRevert;
              if (next < 0) next = 0;
              thMap[j][3] = Math.floor(next / cb);
              thMap[j][5] = next % cb;
              thMap[j][6] = next;
              thMap[j][7] = (next * (ttMap[tenHang] || 0)) / 1000;
              break;
            }
          }
        } else {
          remaining.push(row);
        }
      }
      
      sHist.clearContent();
      if (remaining.length > 0) sHist.getRange(1, 1, remaining.length, 11).setValues(remaining);
      
      if (sTh) {
        if (sTh.getLastRow() >= 2) sTh.getRange(2, 1, sTh.getLastRow() - 1, 9).clearContent();
        var finalTh = thMap.filter(function(r) { return r[6] > 0; });
        if (finalTh.length > 0) sTh.getRange(2, 1, finalTh.length, 9).setValues(finalTh);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "saveContData") {
      var rows = postData.data || [];
      var isReplace = postData.isReplace || false;
      var sCont = ss.getSheetByName("Cont");
      if (!sCont) return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "Sheet Cont không tồn tại!" })).setMimeType(ContentService.MimeType.JSON);
      
      if (isReplace && rows.length > 0) {
        var finishedName = String(rows[0][7]).trim().toUpperCase();
        var currentCont = [];
        if (sCont.getLastRow() >= 2) currentCont = sCont.getRange(2, 1, sCont.getLastRow() - 1, 12).getValues();
        var unchanged = currentCont.filter(function(r) { return String(r[7]).trim().toUpperCase() !== finishedName; });
        var header = ["STT","ID","Phân Cấp","Tiêu chuẩn","Kích Cỡ","Bề Mặt","Số bàn giao (NO)","Tên Container","Bao","CB","Lẻ","Ghi chú"];
        var finalRows = [header];
        unchanged.forEach(function(r) { finalRows.push(r); });
        rows.forEach(function(r) {
          var tr = [].concat(r);
          tr[8] = 0; tr[9] = 0; tr[10] = 0; tr[11] = "";
          finalRows.push(tr);
        });
        sCont.clearContent();
        sCont.getRange(1, 1, finalRows.length, 12).setValues(finalRows);
      } else {
        rows.forEach(function(r) {
          var tr = [].concat(r);
          tr[8] = 0; tr[9] = 0; tr[10] = 0; tr[11] = "";
          sCont.appendRow(tr);
        });
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "updateContQuantity") {
      var sCont = ss.getSheetByName("Cont");
      if (!sCont || sCont.getLastRow() < 2) return ContentService.createTextOutput(JSON.stringify({ success: false, msg: "Không có dữ liệu trong sheet Cont!" })).setMimeType(ContentService.MimeType.JSON);
      
      var table = sCont.getRange(2, 1, sCont.getLastRow() - 1, 12).getValues();
      var foundIdx = -1;
      for (var i = 0; i < table.length; i++) {
        var row = table[i];
        if (
          String(row[2]).trim() === String(postData.grade).trim() &&
          String(row[3]).trim() === String(postData.standard).trim() &&
          String(row[4]).trim() === String(postData.size).trim() &&
          String(row[5]).trim() === String(postData.finish).trim() &&
          String(row[6]).trim() === String(postData.no).trim() &&
          String(row[7]).trim() === String(postData.contName).trim()
        ) {
          foundIdx = i;
          break;
        }
      }
      if (foundIdx !== -1) {
        var rIdx = foundIdx + 2;
        sCont.getRange(rIdx, 3).setValue(postData.newG || postData.grade);
        sCont.getRange(rIdx, 4).setValue(postData.newS || postData.standard);
        sCont.getRange(rIdx, 5).setValue(postData.newSz || postData.size);
        sCont.getRange(rIdx, 6).setValue(postData.newF || postData.finish);
        sCont.getRange(rIdx, 7).setValue(postData.newNo || postData.no);
        sCont.getRange(rIdx, 9).setValue(postData.b || 0);
        sCont.getRange(rIdx, 10).setValue(postData.cb || 0);
        sCont.getRange(rIdx, 11).setValue(postData.l || 0);
        sCont.getRange(rIdx, 12).setValue(postData.n || "");
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, msg: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}`;
                              navigator.clipboard.writeText(codeStr);
                              showToast("Đã sao chép Apps Script Code!", "success");
                            }}
                            className="absolute top-1 right-1 bg-blue-600 hover:bg-blue-700 text-white rounded px-1.5 py-0.5 text-[9px] font-bold cursor-pointer animate-pulse"
                          >
                            SAO CHÉP
                          </button>
                          <pre className="text-left text-slate-100 pointer-events-none select-all select-none">
{`var SPREADSHEET_ID = "ID_TRANG_TINH_CUA_BAN";

function doGet(e) {
  var action = e.parameter.action;
  ... (Nhấp nút Sao Chép phía trên)`}
                          </pre>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {gsConfig.connectionType === "oauth" && (
                <div className="space-y-4">
                  {/* Google Auth Integrated Button and State */}
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Thực Hiện Cấp Quyền Đọc Ghi Trực Tiếp:</p>
                    
                    {firebaseUser ? (
                      <div className="space-y-3">
                        <div className="flex items-center gap-2">
                          {firebaseUser.photoURL ? (
                            <img src={firebaseUser.photoURL} alt={firebaseUser.displayName || "Google User"} className="h-8 w-8 rounded-full border border-teal-200 shadow-sm" referrerPolicy="no-referrer" />
                          ) : (
                            <div className="h-8 w-8 rounded-full bg-slate-200 text-slate-600 font-bold flex items-center justify-center text-xs">G</div>
                          )}
                          <div className="text-[11px] leading-tight">
                            <p className="font-bold text-slate-800">{firebaseUser.displayName}</p>
                            <p className="text-slate-400 font-mono text-[9px]">{firebaseUser.email}</p>
                          </div>
                        </div>
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={handleGoogleSignOut}
                            className="text-[10px] font-bold text-rose-600 bg-rose-50 border border-rose-200 hover:bg-rose-100 rounded-lg px-2.5 py-1.5 cursor-pointer transition-colors"
                          >
                            Đăng xuất
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <p className="text-[11px] text-slate-500 leading-normal mb-2">
                          Bạn có thể đăng nhập Google để tự động định danh và cấp quyền trực tiếp cho ứng dụng.
                        </p>
                        <button
                          type="button"
                          onClick={handleGoogleSignIn}
                          className="w-full flex items-center justify-center gap-2 px-4 py-2.5 border border-slate-300 rounded-xl bg-white hover:bg-slate-100 text-slate-700 font-bold text-xs shadow-sm hover:shadow-md cursor-pointer transition-all duration-150"
                        >
                          <svg version="1.1" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" className="h-4 w-4">
                            <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"></path>
                            <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"></path>
                            <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"></path>
                            <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"></path>
                            <path fill="none" d="M0 0h48v48H0z"></path>
                          </svg>
                          <span>Liên Kết Đăng Nhập Google</span>
                        </button>
                      </div>
                    )}
                  </div>

                  <div>
                    <div className="flex items-center justify-between mb-1">
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest">
                        Nhập Token Thủ Công (Tùy Chọn):
                      </label>
                      {googleToken && (
                        <span className="text-[10px] font-bold text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded-md">
                          Đã nạp tự động ✅
                        </span>
                      )}
                    </div>
                    <input
                      type="text"
                      className="block w-full rounded-xl border border-slate-200 px-3.5 py-2.5 font-mono text-slate-800 placeholder-slate-300"
                      placeholder="ya29.a0... hoặc để trống để sử dụng Đăng nhập"
                      value={gsConfig.accessToken}
                      onChange={(e) => {
                        setGsConfig(prev => ({ ...prev, accessToken: e.target.value }));
                        setTestResult(null);
                      }}
                    />
                    <span className="block text-[9px] text-slate-400 mt-1">
                      * Nhập Access Token thủ công nếu bạn không muốn đăng nhập tài khoản trực tiếp qua popup.
                    </span>
                  </div>
                </div>
              )}

              {/* TEST RESULT SECTION */}
              {testResult && (
                <div className={`p-3 rounded-xl border text-[11px] font-medium leading-relaxed flex items-start gap-2 ${
                  testResult.success 
                    ? "bg-emerald-50 border-emerald-200 text-emerald-800" 
                    : "bg-rose-50 border-rose-200 text-rose-800"
                }`}>
                  <span className="text-lg leading-none">{testResult.success ? "✅" : "⚠️"}</span>
                  <div>
                    <p className="font-bold">{testResult.success ? "Kết nối hoạt động tốt!" : "Lỗi đồng bộ hoặc kết nối thất bại!"}</p>
                    <p className="text-[10px] opacity-90">{testResult.message}</p>
                  </div>
                </div>
              )}

              <div className="flex flex-wrap gap-2 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    const reset = {
                      spreadsheetId: "1QKueqrlHhA0xhW1ZnkE4QrpKLUVD6eDTRNYxUMBtbhc",
                      connectionType: "appscript" as const,
                      webAppUrl: "https://script.google.com/macros/s/AKfycbxuHHa2RNny5HCPyswWwRAZDFs8F6KiL0nWIE2oDs9X-dEoOLoWgiI7fCopvD7n3gtCYw/exec",
                      accessToken: "",
                      clientId: "",
                    };
                    setGsConfig(reset);
                    saveConfig(reset);
                    googleSheetsService["config"] = reset;
                    setTestResult(null);
                    showToast("Đã khôi phục thiết lập mặc định!", "info");
                    setIsConfigModalOpen(false);
                    reloadAllData();
                  }}
                  className="px-3.5 py-2.5 rounded-xl border border-slate-200 hover:bg-slate-50 font-bold text-slate-500 hover:text-slate-700 cursor-pointer text-center whitespace-nowrap text-xs transition-colors"
                >
                  Xóa cấu hình
                </button>
                <button
                  type="button"
                  disabled={isTestingConnection}
                  onClick={handleTestConnection}
                  className="px-3 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold cursor-pointer text-center whitespace-nowrap text-xs transition-colors disabled:opacity-50"
                >
                  {isTestingConnection ? "Đang kiểm tra..." : "Kiểm tra kết nối"}
                </button>
                <button
                  type="button"
                  onClick={async () => {
                    saveConfig(gsConfig);
                    googleSheetsService["config"] = gsConfig;
                    showToast("Lưu cấu hình và đang đồng bộ dữ liệu...", "success");
                    setIsConfigModalOpen(false);
                    await reloadAllData();
                  }}
                  className="flex-1 bg-blue-600 hover:bg-blue-700 text-white rounded-xl py-2.5 font-bold cursor-pointer text-center shadow text-xs transition-transform hover:scale-[1.01]"
                >
                  ĐỒNG BỘ NGAY
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* POPUP LIVE TOAST FEEDBACK notifications */}
      {toast && (
        <div className="fixed bottom-4 right-4 z-50 flex items-center gap-2.5 px-4.5 py-3 bg-slate-900/95 backdrop-blur text-white text-xs font-semibold rounded-xl border border-white/10 shadow-2xl animate-[slideIn_0.3s_ease-out-back]">
          {toast.type === "success" && <CheckCircle className="h-4.5 w-4.5 text-emerald-400 shrink-0" />}
          {toast.type === "error" && <AlertTriangle className="h-4.5 w-4.5 text-rose-500 shrink-0" />}
          {toast.type === "info" && <Info className="h-4.5 w-4.5 text-blue-400 shrink-0" />}
          <span>{toast.message}</span>
        </div>
      )}
    </div>
  );
}
