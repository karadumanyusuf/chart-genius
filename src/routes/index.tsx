import { createFileRoute } from "@tanstack/react-router";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { toPng, toJpeg, toSvg } from "html-to-image";
import jsPDF from "jspdf";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ScatterChart, Scatter, ComposedChart,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  RadialBarChart, RadialBar, LabelList, Label as RLabel,
} from "recharts";
import {
  Upload, FileSpreadsheet, Image as ImageIcon, Download, Sparkles,
  BarChart3, LineChart as LineIcon, PieChart as PieIcon, Radar as RadarIcon,
  AreaChart as AreaIcon, ScatterChart as ScatterIcon, Activity, Loader2,
  Palette, Table2, Trash2, Database, FileJson, FileText, FileImage, Code2,
  HardDrive, Save, FolderOpen, Github, Linkedin, Info, Workflow, Mail, Languages,
  Pencil, Hash, Settings2, RotateCcw, Type, Ruler,
} from "lucide-react";
import {
  Sheet, SheetContent, SheetHeader, SheetTitle, SheetDescription,
} from "@/components/ui/sheet";
import { toast } from "sonner";
import { Toaster } from "@/components/ui/sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger,
  DropdownMenuLabel, DropdownMenuSeparator,
} from "@/components/ui/dropdown-menu";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "ChartGlass — Veriden grafiğe saniyeler içinde" },
      { name: "description", content: "Excel, CSV ve görsellerden veri okuyup özelleştirilebilir grafiklere dönüştüren modern web aracı." },
    ],
  }),
  component: ChartGlass,
});

type Row = Record<string, string | number>;
type Lang = "tr" | "en";

type ChartKind =
  | "bar" | "stackedBar" | "horizontalBar"
  | "line" | "smoothLine" | "area" | "stackedArea"
  | "pie" | "donut" | "radar" | "radial" | "scatter" | "composed";

const I18N = {
  tr: {
    tagline: "Veriden grafiğe — saniyeler içinde",
    download: "İndir", chartImage: "Grafik (resim)", dataSection: "Veri",
    uploadData: "Veri Yükle", dragHint: "Excel, CSV veya görseli sürükleyin",
    orPaste: "veya doğrudan yapıştır (CSV/TSV)", process: "İşle",
    readingImage: "Görsel okunuyor...",
    apiSection: "Veritabanı / API",
    apiDesc: "JSON döndüren bir REST endpoint, Supabase/PostgREST veya herhangi bir veri API'si bağlayın.",
    urlLabel: "URL", tokenLabel: "Bearer Token (opsiyonel)", fetch: "Veriyi Çek",
    localDb: "Yerel Veritabanı",
    localDbDesc: "Verilerinizi tarayıcıda saklayın — sayfayı kapatsanız bile kalır.",
    datasetName: "Veri seti adı", save: "Kaydet", noSaved: "Henüz kayıtlı veri yok.",
    rowsSuffix: "satır",
    chartTitle: "Grafik başlığı", noData: "Görüntülemek için veri yükleyin",
    dataPreview: "Veri Önizleme", editHint: "Hücreye tıklayarak veriyi düzenleyebilirsiniz",
    chartType: "Grafik Türü", colorPalette: "Renk Paleti",
    showValues: "Değerleri göster", showAxisLabels: "Eksen adlarını göster",
    valueAxis: "Değer",
    purpose: "Amaç",
    purposeText: "ChartGlass; Excel, CSV, görsel veya API'den gelen verileri saniyeler içinde özelleştirilebilir grafiklere dönüştürmek için tasarlandı. Sunumlar, raporlar ve hızlı analizler için hazır görseller üretin.",
    howTo: "Nasıl Kullanılır?",
    step1: "Soldan veri yükleyin (dosya, API veya yapıştır).",
    step2: "Sağdan grafik türünü ve renk paletini seçin.",
    step3: "Üstten istediğiniz formatta (PNG, PDF, SVG, XLSX) indirin.",
    contact: "İletişim",
    contactText: "Geri bildirim, fikir veya katkı için ulaşabilirsiniz.",
    promptEdit: "Yeni değer:",
    editChart: "Grafiği Düzenle",
    editorDesc: "Veri değerlerini, metinleri, boyutları ve serilerin renklerini özelleştirin.",
    titleText: "Başlık metni",
    chartHeight: "Grafik yüksekliği (px)",
    titleSize: "Başlık yazı boyutu",
    axisSize: "Eksen yazı boyutu",
    valueSize: "Değer etiketi boyutu",
    seriesColors: "Seri renkleri",
    dataValues: "Veri değerleri",
    resetColors: "Paletten sıfırla",
    sectionGeneral: "Genel",
    sectionSizes: "Boyutlar",
    sectionColors: "Renkler",
    sectionData: "Veriler",
    chartLabels: {
      bar: "Sütun", stackedBar: "Yığılmış Sütun", horizontalBar: "Yatay Çubuk",
      line: "Çizgi", smoothLine: "Yumuşak Çizgi", area: "Alan", stackedArea: "Yığılmış Alan",
      pie: "Pasta", donut: "Halka", radar: "Radar", radial: "Radyal",
      scatter: "Dağılım", composed: "Karma",
    } as Record<ChartKind, string>,
  },
  en: {
    tagline: "From data to chart — in seconds",
    download: "Download", chartImage: "Chart (image)", dataSection: "Data",
    uploadData: "Upload Data", dragHint: "Drag in Excel, CSV or an image",
    orPaste: "or paste directly (CSV/TSV)", process: "Process",
    readingImage: "Reading image...",
    apiSection: "Database / API",
    apiDesc: "Connect any REST endpoint returning JSON, Supabase/PostgREST or any data API.",
    urlLabel: "URL", tokenLabel: "Bearer Token (optional)", fetch: "Fetch Data",
    localDb: "Local Storage",
    localDbDesc: "Keep your datasets in the browser — they persist after closing.",
    datasetName: "Dataset name", save: "Save", noSaved: "No saved datasets yet.",
    rowsSuffix: "rows",
    chartTitle: "Chart title", noData: "Upload data to visualize",
    dataPreview: "Data Preview", editHint: "Click a cell to edit the data",
    chartType: "Chart Type", colorPalette: "Color Palette",
    showValues: "Show values", showAxisLabels: "Show axis labels",
    valueAxis: "Value",
    purpose: "Purpose",
    purposeText: "ChartGlass turns data from Excel, CSV, images or APIs into customizable charts in seconds. Produce presentation-ready visuals for reports, slides and quick analysis.",
    howTo: "How to Use",
    step1: "Load data on the left (file, API or paste).",
    step2: "Pick a chart type and color palette on the right.",
    step3: "Download in the format you want (PNG, PDF, SVG, XLSX) from the top.",
    contact: "Contact",
    contactText: "Reach out for feedback, ideas or contributions.",
    promptEdit: "New value:",
    editChart: "Edit Chart",
    editorDesc: "Customize data values, text, sizes and series colors.",
    titleText: "Title text",
    chartHeight: "Chart height (px)",
    titleSize: "Title font size",
    axisSize: "Axis font size",
    valueSize: "Value label size",
    seriesColors: "Series colors",
    dataValues: "Data values",
    resetColors: "Reset to palette",
    sectionGeneral: "General",
    sectionSizes: "Sizes",
    sectionColors: "Colors",
    sectionData: "Data",
    chartLabels: {
      bar: "Bar", stackedBar: "Stacked Bar", horizontalBar: "Horizontal Bar",
      line: "Line", smoothLine: "Smooth Line", area: "Area", stackedArea: "Stacked Area",
      pie: "Pie", donut: "Donut", radar: "Radar", radial: "Radial",
      scatter: "Scatter", composed: "Composed",
    } as Record<ChartKind, string>,
  },
} as const;

const CHART_ICONS: { id: ChartKind; icon: React.ComponentType<{ className?: string }> }[] = [
  { id: "bar", icon: BarChart3 },
  { id: "stackedBar", icon: BarChart3 },
  { id: "horizontalBar", icon: BarChart3 },
  { id: "line", icon: LineIcon },
  { id: "smoothLine", icon: LineIcon },
  { id: "area", icon: AreaIcon },
  { id: "stackedArea", icon: AreaIcon },
  { id: "pie", icon: PieIcon },
  { id: "donut", icon: PieIcon },
  { id: "radar", icon: RadarIcon },
  { id: "radial", icon: Activity },
  { id: "scatter", icon: ScatterIcon },
  { id: "composed", icon: Activity },
];

const PALETTES: { name: { tr: string; en: string }; colors: string[] }[] = [
  { name: { tr: "Aurora", en: "Aurora" }, colors: ["#7dd3fc", "#c4b5fd", "#f0abfc", "#fda4af", "#fcd34d", "#86efac"] },
  { name: { tr: "Okyanus", en: "Ocean" }, colors: ["#0ea5e9", "#06b6d4", "#14b8a6", "#22d3ee", "#3b82f6", "#6366f1"] },
  { name: { tr: "Gün Batımı", en: "Sunset" }, colors: ["#f59e0b", "#f97316", "#ef4444", "#ec4899", "#d946ef", "#a855f7"] },
  { name: { tr: "Orman", en: "Forest" }, colors: ["#10b981", "#84cc16", "#a3e635", "#22c55e", "#14b8a6", "#059669"] },
  { name: { tr: "Neon", en: "Neon" }, colors: ["#22d3ee", "#a3e635", "#f472b6", "#facc15", "#60a5fa", "#fb7185"] },
  { name: { tr: "Mono", en: "Mono" }, colors: ["#cbd5e1", "#94a3b8", "#64748b", "#475569", "#334155", "#1e293b"] },
];

const SAMPLE: Row[] = [
  { Ay: "Oca", Satış: 420, Gider: 240, Kar: 180 },
  { Ay: "Şub", Satış: 510, Gider: 280, Kar: 230 },
  { Ay: "Mar", Satış: 680, Gider: 310, Kar: 370 },
  { Ay: "Nis", Satış: 590, Gider: 330, Kar: 260 },
  { Ay: "May", Satış: 770, Gider: 360, Kar: 410 },
  { Ay: "Haz", Satış: 860, Gider: 380, Kar: 480 },
];

function ChartGlass() {
  const [lang, setLang] = useState<Lang>("tr");
  const t = I18N[lang];
  const [rows, setRows] = useState<Row[]>(SAMPLE);
  const [chart, setChart] = useState<ChartKind>("bar");
  const [palette, setPalette] = useState(0);
  const [title, setTitle] = useState("Aylık Performans");
  const [showValues, setShowValues] = useState(true);
  const [showAxisLabels, setShowAxisLabels] = useState(true);
  const [ocrBusy, setOcrBusy] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [apiUrl, setApiUrl] = useState("");
  const [apiToken, setApiToken] = useState("");
  const [apiBusy, setApiBusy] = useState(false);
  const [saveName, setSaveName] = useState("");
  const [saved, setSaved] = useState<{ name: string; rows: Row[]; chart: ChartKind; palette: number; title: string; savedAt: number }[]>([]);
  const [editorOpen, setEditorOpen] = useState(false);
  const [chartHeight, setChartHeight] = useState(460);
  const [titleSize, setTitleSize] = useState(18);
  const [axisSize, setAxisSize] = useState(12);
  const [valueSize, setValueSize] = useState(11);
  const [seriesColors, setSeriesColors] = useState<Record<string, string>>({});
  const chartRef = useRef<HTMLDivElement>(null);

  const LS_KEY = "chartglass:datasets";
  const LS_LAST = "chartglass:last";
  const LS_LANG = "chartglass:lang";

  useEffect(() => {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (raw) setSaved(JSON.parse(raw));
      const last = localStorage.getItem(LS_LAST);
      if (last) {
        const s = JSON.parse(last);
        if (s.rows?.length) setRows(s.rows);
        if (s.chart) setChart(s.chart);
        if (typeof s.palette === "number") setPalette(s.palette);
        if (s.title) setTitle(s.title);
      }
      const l = localStorage.getItem(LS_LANG);
      if (l === "tr" || l === "en") setLang(l);
    } catch (e) { console.warn(e); }
  }, []);

  useEffect(() => {
    try { localStorage.setItem(LS_LAST, JSON.stringify({ rows, chart, palette, title })); } catch (e) { console.warn(e); }
  }, [rows, chart, palette, title]);

  useEffect(() => {
    try { localStorage.setItem(LS_LANG, lang); } catch (e) { console.warn(e); }
  }, [lang]);

  const persistSaved = (next: typeof saved) => {
    setSaved(next);
    try { localStorage.setItem(LS_KEY, JSON.stringify(next)); } catch (e) { console.warn(e); }
  };

  const saveDataset = () => {
    const name = saveName.trim() || title.trim() || `Veri ${new Date().toLocaleString()}`;
    if (!rows.length) { toast.error(lang === "tr" ? "Kaydedilecek veri yok" : "Nothing to save"); return; }
    const entry = { name, rows, chart, palette, title, savedAt: Date.now() };
    const next = [entry, ...saved.filter((s) => s.name !== name)].slice(0, 50);
    persistSaved(next);
    setSaveName("");
    toast.success(`"${name}" ${lang === "tr" ? "kaydedildi" : "saved"}`);
  };

  const loadDataset = (name: string) => {
    const s = saved.find((x) => x.name === name);
    if (!s) return;
    setRows(s.rows); setChart(s.chart); setPalette(s.palette); setTitle(s.title);
    toast.success(`"${name}" ${lang === "tr" ? "yüklendi" : "loaded"}`);
  };

  const deleteDataset = (name: string) => {
    persistSaved(saved.filter((s) => s.name !== name));
  };

  const { keys, categoryKey, numericKeys } = useMemo(() => {
    if (!rows.length) return { keys: [], categoryKey: "", numericKeys: [] as string[] };
    const k = Object.keys(rows[0]);
    const numeric = k.filter((key) => rows.every((r) => typeof r[key] === "number" || !isNaN(Number(r[key]))));
    const cat = k.find((key) => !numeric.includes(key)) ?? k[0];
    return { keys: k, categoryKey: cat, numericKeys: numeric.filter((n) => n !== cat) };
  }, [rows]);

  const colors = PALETTES[palette].colors;

  const updateCell = (rowIndex: number, key: string, value: string) => {
    setRows((prev) => {
      const next = [...prev];
      const num = Number(value.replace(/[^\d.,-]/g, "").replace(",", "."));
      next[rowIndex] = { ...next[rowIndex], [key]: !isNaN(num) && value.trim() !== "" && /\d/.test(value) ? num : value };
      return next;
    });
  };

  // Click on a bar/point → prompt to edit
  const editPoint = (rowIndex: number, key: string) => {
    const current = rows[rowIndex]?.[key];
    const next = window.prompt(`${key} — ${t.promptEdit}`, String(current ?? ""));
    if (next === null) return;
    updateCell(rowIndex, key, next);
  };

  const normalizeRows = (raw: Row[]): Row[] =>
    raw.map((r) => {
      const out: Row = {};
      for (const [k, v] of Object.entries(r)) {
        const num = typeof v === "string" ? Number(v.replace(/[^\d.,-]/g, "").replace(",", ".")) : v;
        out[k] = typeof num === "number" && !isNaN(num) && v !== "" ? num : (v as string);
      }
      return out;
    });

  const handleFile = useCallback(async (file: File) => {
    const ext = file.name.split(".").pop()?.toLowerCase();
    try {
      if (ext === "csv" || ext === "tsv") {
        const text = await file.text();
        const parsed = Papa.parse<Row>(text, { header: true, skipEmptyLines: true, dynamicTyping: true });
        setRows(normalizeRows(parsed.data as Row[]));
        toast.success(`${parsed.data.length} ${t.rowsSuffix}`);
      } else if (["xlsx", "xls", "ods"].includes(ext ?? "")) {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Row>(ws);
        setRows(normalizeRows(json));
        toast.success(`${json.length} ${t.rowsSuffix}`);
      } else if (file.type.startsWith("image/")) {
        setOcrBusy(true);
        toast.info(t.readingImage);
        const Tesseract = (await import("tesseract.js")).default;
        const { data } = await Tesseract.recognize(file, "eng+tur");
        const text = data.text.trim();
        const parsed = Papa.parse<Row>(text, { header: true, skipEmptyLines: true, delimiter: "" });
        if (parsed.data.length && Object.keys(parsed.data[0] as object).length > 1) {
          setRows(normalizeRows(parsed.data as Row[]));
          toast.success(`${parsed.data.length} ${t.rowsSuffix}`);
        } else {
          setPasteText(text);
        }
        setOcrBusy(false);
      } else {
        toast.error(lang === "tr" ? "Desteklenmeyen dosya türü" : "Unsupported file type");
      }
    } catch (e) {
      console.error(e);
      toast.error(lang === "tr" ? "Dosya işlenemedi" : "Failed to read file");
      setOcrBusy(false);
    }
  }, [lang, t.readingImage, t.rowsSuffix]);

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  };

  const parsePasted = () => {
    if (!pasteText.trim()) return;
    const parsed = Papa.parse<Row>(pasteText, {
      header: true, skipEmptyLines: true, delimitersToGuess: ["\t", ",", ";", "|"],
    });
    if (parsed.data.length) {
      setRows(normalizeRows(parsed.data as Row[]));
      toast.success(`${parsed.data.length} ${t.rowsSuffix}`);
    } else {
      toast.error(lang === "tr" ? "Veri okunamadı" : "Could not parse data");
    }
  };

  const downloadBlob = (blob: Blob | string, filename: string) => {
    const url = typeof blob === "string" ? blob : URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; a.click();
    if (typeof blob !== "string") setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const exportChart = async (format: "png" | "jpeg" | "svg" | "pdf") => {
    if (!chartRef.current) return;
    try {
      const opts = { cacheBust: true, pixelRatio: 2, backgroundColor: "#1a1f3a" };
      if (format === "png") downloadBlob(await toPng(chartRef.current, opts), `${title || "chart"}.png`);
      else if (format === "jpeg") downloadBlob(await toJpeg(chartRef.current, { ...opts, quality: 0.95 }), `${title || "chart"}.jpg`);
      else if (format === "svg") downloadBlob(await toSvg(chartRef.current, opts), `${title || "chart"}.svg`);
      else if (format === "pdf") {
        const dataUrl = await toPng(chartRef.current, opts);
        const img = new window.Image();
        img.src = dataUrl;
        await new Promise((r) => (img.onload = r));
        const orientation = img.width > img.height ? "landscape" : "portrait";
        const pdf = new jsPDF({ orientation, unit: "px", format: [img.width, img.height] });
        pdf.addImage(dataUrl, "PNG", 0, 0, img.width, img.height);
        pdf.save(`${title || "chart"}.pdf`);
      }
    } catch (e) {
      console.error(e);
      toast.error(lang === "tr" ? "Dışa aktarma başarısız" : "Export failed");
    }
  };

  const exportData = (format: "csv" | "json" | "xlsx") => {
    if (!rows.length) return;
    try {
      if (format === "csv") downloadBlob(new Blob([Papa.unparse(rows)], { type: "text/csv;charset=utf-8" }), `${title || "data"}.csv`);
      else if (format === "json") downloadBlob(new Blob([JSON.stringify(rows, null, 2)], { type: "application/json" }), `${title || "data"}.json`);
      else if (format === "xlsx") {
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");
        XLSX.writeFile(wb, `${title || "data"}.xlsx`);
      }
    } catch { toast.error(lang === "tr" ? "Dışa aktarma başarısız" : "Export failed"); }
  };

  const clearData = () => { setRows([]); setPasteText(""); };

  const fetchFromUrl = async () => {
    const url = apiUrl.trim();
    if (!url) return;
    try {
      setApiBusy(true);
      const headers: Record<string, string> = { Accept: "application/json" };
      if (apiToken.trim()) headers.Authorization = apiToken.trim().startsWith("Bearer ") ? apiToken.trim() : `Bearer ${apiToken.trim()}`;
      const res = await fetch(url, { headers });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const ct = res.headers.get("content-type") || "";
      let data: unknown;
      if (ct.includes("application/json")) data = await res.json();
      else {
        const text = await res.text();
        try { data = JSON.parse(text); } catch {
          const parsed = Papa.parse<Row>(text, { header: true, skipEmptyLines: true, dynamicTyping: true });
          data = parsed.data;
        }
      }
      let arr: Row[] | null = null;
      if (Array.isArray(data)) arr = data as Row[];
      else if (data && typeof data === "object") {
        for (const key of ["data", "results", "rows", "records", "items"]) {
          const v = (data as Record<string, unknown>)[key];
          if (Array.isArray(v)) { arr = v as Row[]; break; }
        }
      }
      if (!arr || !arr.length) throw new Error(lang === "tr" ? "Yanıtta dizi bulunamadı" : "No array in response");
      const flat = arr.map((item) => {
        const out: Row = {};
        for (const [k, v] of Object.entries(item as object)) {
          if (v === null || typeof v === "object") continue;
          out[k] = v as string | number;
        }
        return out;
      });
      setRows(flat);
      toast.success(`${flat.length} ${t.rowsSuffix}`);
    } catch (e) {
      console.error(e);
      toast.error(`${(e as Error).message}`);
    } finally { setApiBusy(false); }
  };

  return (
    <div className="min-h-screen text-foreground">
      <Toaster theme="dark" />
      <div className="mx-auto max-w-[1500px] px-4 py-8 md:px-8 md:py-12">
        <header className="mb-10 flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
          <div className="flex items-center gap-3">
            <div className="glass-panel glow-primary flex h-12 w-12 items-center justify-center">
              <Sparkles className="h-6 w-6 text-primary" />
            </div>
            <div>
              <h1 className="text-2xl font-semibold tracking-tight md:text-3xl">ChartGlass</h1>
              <p className="text-sm text-muted-foreground">{t.tagline}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <div className="glass-panel inline-flex items-center gap-1 p-1">
              <Languages className="ml-1.5 h-3.5 w-3.5 text-muted-foreground" />
              <button
                onClick={() => setLang("tr")}
                className={`px-2.5 py-1 text-xs rounded-md transition ${lang === "tr" ? "bg-primary/20 text-primary" : "text-muted-foreground hover:text-foreground"}`}
              >TR</button>
              <button
                onClick={() => setLang("en")}
                className={`px-2.5 py-1 text-xs rounded-md transition ${lang === "en" ? "bg-primary/20 text-primary" : "text-muted-foreground hover:text-foreground"}`}
              >EN</button>
            </div>
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button disabled={!rows.length} className="glass-panel border-0 bg-primary/90 text-primary-foreground hover:bg-primary">
                  <Download className="mr-2 h-4 w-4" /> {t.download}
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="glass-panel border-0 w-56">
                <DropdownMenuLabel className="text-xs text-muted-foreground">{t.chartImage}</DropdownMenuLabel>
                <DropdownMenuItem onClick={() => exportChart("png")}><FileImage className="mr-2 h-4 w-4" /> PNG</DropdownMenuItem>
                <DropdownMenuItem onClick={() => exportChart("jpeg")}><FileImage className="mr-2 h-4 w-4" /> JPEG</DropdownMenuItem>
                <DropdownMenuItem onClick={() => exportChart("svg")}><Code2 className="mr-2 h-4 w-4" /> SVG</DropdownMenuItem>
                <DropdownMenuItem onClick={() => exportChart("pdf")}><FileText className="mr-2 h-4 w-4" /> PDF</DropdownMenuItem>
                <DropdownMenuSeparator />
                <DropdownMenuLabel className="text-xs text-muted-foreground">{t.dataSection}</DropdownMenuLabel>
                <DropdownMenuItem onClick={() => exportData("csv")}><FileText className="mr-2 h-4 w-4" /> CSV</DropdownMenuItem>
                <DropdownMenuItem onClick={() => exportData("xlsx")}><FileSpreadsheet className="mr-2 h-4 w-4" /> Excel (XLSX)</DropdownMenuItem>
                <DropdownMenuItem onClick={() => exportData("json")}><FileJson className="mr-2 h-4 w-4" /> JSON</DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        </header>

        <div className="grid gap-6 lg:grid-cols-[320px_minmax(0,1fr)_300px]">
          <aside className="space-y-6 min-w-0 lg:order-1">
            <section className="glass-panel p-5" onDrop={onDrop} onDragOver={(e) => e.preventDefault()}>
              <div className="mb-3 flex items-center gap-2">
                <Upload className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{t.uploadData}</h2>
              </div>
              <label className="glass-soft flex cursor-pointer flex-col items-center justify-center gap-2 px-4 py-8 text-center transition hover:border-primary/50">
                <div className="flex gap-3 text-muted-foreground">
                  <FileSpreadsheet className="h-6 w-6" />
                  <ImageIcon className="h-6 w-6" />
                </div>
                <span className="text-sm">{t.dragHint}</span>
                <span className="text-xs text-muted-foreground">.xlsx · .csv · .png · .jpg</span>
                <input type="file" accept=".xlsx,.xls,.csv,.tsv,.ods,image/*" className="hidden"
                  onChange={(e) => e.target.files?.[0] && handleFile(e.target.files[0])} />
                {ocrBusy && (
                  <div className="mt-2 flex items-center gap-2 text-xs text-primary">
                    <Loader2 className="h-3 w-3 animate-spin" /> {t.readingImage}
                  </div>
                )}
              </label>
              <div className="mt-4">
                <Label className="text-xs text-muted-foreground">{t.orPaste}</Label>
                <Textarea value={pasteText} onChange={(e) => setPasteText(e.target.value)}
                  placeholder={"Ay,Satış,Gider\nOca,420,240\nŞub,510,280"}
                  className="glass-soft mt-2 h-28 resize-none border-0 font-mono text-xs" />
                <div className="mt-2 flex gap-2">
                  <Button size="sm" onClick={parsePasted} className="flex-1 bg-primary/90 text-primary-foreground hover:bg-primary">
                    {t.process}
                  </Button>
                  <Button size="sm" variant="ghost" onClick={clearData}>
                    <Trash2 className="h-4 w-4" />
                  </Button>
                </div>
              </div>
            </section>

            <section className="glass-panel p-5">
              <div className="mb-3 flex items-center gap-2">
                <Database className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{t.apiSection}</h2>
              </div>
              <p className="mb-3 text-xs text-muted-foreground">{t.apiDesc}</p>
              <Label className="text-xs text-muted-foreground">{t.urlLabel}</Label>
              <Input value={apiUrl} onChange={(e) => setApiUrl(e.target.value)}
                placeholder="https://api.ornek.com/satislar" className="glass-soft mt-1 mb-2 border-0 text-xs" />
              <Label className="text-xs text-muted-foreground">{t.tokenLabel}</Label>
              <Input type="password" value={apiToken} onChange={(e) => setApiToken(e.target.value)}
                placeholder="eyJhbGciOi..." className="glass-soft mt-1 mb-2 border-0 text-xs" />
              <Button size="sm" onClick={fetchFromUrl} disabled={apiBusy || !apiUrl.trim()}
                className="w-full bg-primary/90 text-primary-foreground hover:bg-primary">
                {apiBusy ? <Loader2 className="mr-2 h-3 w-3 animate-spin" /> : <Database className="mr-2 h-3 w-3" />}
                {t.fetch}
              </Button>
            </section>

            <section className="glass-panel p-5">
              <div className="mb-3 flex items-center gap-2">
                <HardDrive className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{t.localDb}</h2>
              </div>
              <p className="mb-3 text-xs text-muted-foreground">{t.localDbDesc}</p>
              <div className="flex gap-2">
                <Input value={saveName} onChange={(e) => setSaveName(e.target.value)}
                  placeholder={t.datasetName} className="glass-soft border-0 text-xs" />
                <Button size="sm" onClick={saveDataset} disabled={!rows.length}
                  className="bg-primary/90 text-primary-foreground hover:bg-primary">
                  <Save className="mr-1 h-3 w-3" /> {t.save}
                </Button>
              </div>
              {saved.length > 0 ? (
                <div className="mt-3 max-h-48 space-y-1.5 overflow-auto pr-1">
                  {saved.map((s) => (
                    <div key={s.name} className="glass-soft flex items-center justify-between gap-2 px-2.5 py-1.5">
                      <button onClick={() => loadDataset(s.name)}
                        className="flex min-w-0 flex-1 items-center gap-2 text-left text-xs hover:text-primary">
                        <FolderOpen className="h-3 w-3 shrink-0 text-primary/70" />
                        <span className="truncate">{s.name}</span>
                        <span className="ml-auto shrink-0 text-[10px] text-muted-foreground">
                          {s.rows.length} {t.rowsSuffix}
                        </span>
                      </button>
                      <button onClick={() => deleteDataset(s.name)}
                        className="shrink-0 text-muted-foreground hover:text-destructive" aria-label="Delete">
                        <Trash2 className="h-3 w-3" />
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="mt-3 text-[10px] text-muted-foreground">{t.noSaved}</p>
              )}
            </section>
          </aside>

          <main className="space-y-6 min-w-0 lg:order-2">
            <section className="glass-panel p-5 md:p-7 relative">
              <div className="mb-4 pr-32">
                <Label className="text-xs text-muted-foreground">{t.chartTitle}</Label>
                <Input value={title} onChange={(e) => setTitle(e.target.value)}
                  className="glass-soft mt-1 border-0 text-base font-medium" />
              </div>

              <Button
                size="sm"
                onClick={() => setEditorOpen(true)}
                disabled={!rows.length}
                className="absolute right-5 top-5 md:right-7 md:top-7 glass-panel border-0 bg-primary/90 text-primary-foreground hover:bg-primary z-10"
              >
                <Settings2 className="mr-1.5 h-4 w-4" /> {t.editChart}
              </Button>

              <div ref={chartRef} className="glass-soft p-4 md:p-6" style={{ background: "oklch(0.22 0.04 265 / 0.6)" }}>
                <h3 className="mb-4 text-center font-semibold" style={{ fontSize: titleSize }}>{title}</h3>
                {rows.length && numericKeys.length ? (
                  <div style={{ height: chartHeight }}>
                    <ResponsiveContainer width="100%" height="100%">
                      {renderChart(chart, rows, categoryKey, numericKeys, colors, {
                        showValues, showAxisLabels, valueAxisLabel: t.valueAxis, onEdit: editPoint,
                        seriesColors, axisSize, valueSize,
                      })}
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="flex items-center justify-center text-sm text-muted-foreground" style={{ height: chartHeight }}>
                    {t.noData}
                  </div>
                )}
              </div>

              {rows.length > 0 && (
                <p className="mt-3 flex items-center justify-center gap-1.5 text-[11px] text-muted-foreground">
                  <Pencil className="h-3 w-3" /> {t.editHint}
                </p>
              )}
            </section>

            {/* Editor Sheet */}
            <Sheet open={editorOpen} onOpenChange={setEditorOpen}>
              <SheetContent side="right" className="w-full sm:max-w-md overflow-y-auto bg-background/95 backdrop-blur-xl border-l border-white/10">
                <SheetHeader>
                  <SheetTitle className="flex items-center gap-2">
                    <Settings2 className="h-4 w-4 text-primary" /> {t.editChart}
                  </SheetTitle>
                  <SheetDescription className="text-xs">{t.editorDesc}</SheetDescription>
                </SheetHeader>

                <div className="mt-6 space-y-6">
                  {/* General */}
                  <div className="space-y-3">
                    <h3 className="text-xs font-semibold uppercase tracking-wider text-primary flex items-center gap-1.5">
                      <Type className="h-3 w-3" /> {t.sectionGeneral}
                    </h3>
                    <div>
                      <Label className="text-xs text-muted-foreground">{t.titleText}</Label>
                      <Input value={title} onChange={(e) => setTitle(e.target.value)} className="glass-soft mt-1 border-0" />
                    </div>
                  </div>

                  {/* Sizes */}
                  <div className="space-y-3">
                    <h3 className="text-xs font-semibold uppercase tracking-wider text-primary flex items-center gap-1.5">
                      <Ruler className="h-3 w-3" /> {t.sectionSizes}
                    </h3>
                    <SizeRow label={t.chartHeight} value={chartHeight} min={200} max={900} step={10} onChange={setChartHeight} />
                    <SizeRow label={t.titleSize} value={titleSize} min={10} max={40} step={1} onChange={setTitleSize} />
                    <SizeRow label={t.axisSize} value={axisSize} min={8} max={22} step={1} onChange={setAxisSize} />
                    <SizeRow label={t.valueSize} value={valueSize} min={8} max={22} step={1} onChange={setValueSize} />
                  </div>

                  {/* Series colors */}
                  {numericKeys.length > 0 && (
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <h3 className="text-xs font-semibold uppercase tracking-wider text-primary flex items-center gap-1.5">
                          <Palette className="h-3 w-3" /> {t.seriesColors}
                        </h3>
                        <Button size="sm" variant="ghost" className="h-7 text-[11px]" onClick={() => setSeriesColors({})}>
                          <RotateCcw className="mr-1 h-3 w-3" /> {t.resetColors}
                        </Button>
                      </div>
                      <div className="space-y-2">
                        {numericKeys.map((k, i) => {
                          const current = seriesColors[k] ?? colors[i % colors.length];
                          return (
                            <div key={k} className="glass-soft flex items-center gap-3 px-3 py-2">
                              <input
                                type="color"
                                value={current}
                                onChange={(e) => setSeriesColors({ ...seriesColors, [k]: e.target.value })}
                                className="h-8 w-8 cursor-pointer rounded border-0 bg-transparent"
                              />
                              <span className="flex-1 truncate text-xs">{k}</span>
                              <code className="text-[10px] text-muted-foreground">{current}</code>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* Data values */}
                  {rows.length > 0 && (
                    <div className="space-y-3">
                      <h3 className="text-xs font-semibold uppercase tracking-wider text-primary flex items-center gap-1.5">
                        <Table2 className="h-3 w-3" /> {t.dataValues}
                      </h3>
                      <div className="glass-soft max-h-72 overflow-auto p-0">
                        <table className="w-full text-xs">
                          <thead className="sticky top-0 bg-white/5 backdrop-blur">
                            <tr>
                              {keys.map((k) => (
                                <th key={k} className="px-2 py-1.5 text-left font-medium text-muted-foreground">{k}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {rows.map((r, i) => (
                              <tr key={i} className="border-t border-white/5">
                                {keys.map((k) => (
                                  <td key={k} className="px-1 py-1">
                                    <input
                                      value={String(r[k] ?? "")}
                                      onChange={(e) => updateCell(i, k, e.target.value)}
                                      className="w-full rounded bg-transparent px-1.5 py-1 outline-none focus:bg-primary/15"
                                    />
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              </SheetContent>
            </Sheet>


            {rows.length > 0 && (
              <section className="glass-panel p-5">
                <div className="mb-3 flex items-center gap-2">
                  <Table2 className="h-4 w-4 text-primary" />
                  <h2 className="text-sm font-medium">{t.dataPreview} ({rows.length} {t.rowsSuffix})</h2>
                </div>
                <div className="glass-soft max-h-64 overflow-auto p-0">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 bg-white/5 backdrop-blur">
                      <tr>
                        {keys.map((k) => (
                          <th key={k} className="px-3 py-2 text-left font-medium text-muted-foreground">{k}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows.slice(0, 50).map((r, i) => (
                        <tr key={i} className="border-t border-white/5">
                          {keys.map((k) => (
                            <td key={k} className="px-3 py-1.5">
                              <input
                                value={String(r[k] ?? "")}
                                onChange={(e) => updateCell(i, k, e.target.value)}
                                className="w-full bg-transparent outline-none focus:bg-primary/10 focus:px-1.5 focus:rounded transition-all"
                              />
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            )}
          </main>

          <aside className="space-y-6 min-w-0 lg:order-3">
            <section className="glass-panel p-5">
              <div className="mb-3 flex items-center gap-2">
                <BarChart3 className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{t.chartType}</h2>
              </div>
              <div className="grid grid-cols-3 gap-2 lg:grid-cols-2">
                {CHART_ICONS.map((opt) => {
                  const Icon = opt.icon;
                  const active = chart === opt.id;
                  return (
                    <button key={opt.id} onClick={() => setChart(opt.id)}
                      className={`glass-soft flex flex-col items-center gap-1 p-3 text-xs transition ${
                        active ? "border-primary/70 bg-primary/15 text-primary" : "text-muted-foreground hover:text-foreground"
                      }`}>
                      <Icon className="h-4 w-4" />
                      <span className="text-[10px] leading-tight text-center">{t.chartLabels[opt.id]}</span>
                    </button>
                  );
                })}
              </div>
            </section>

            <section className="glass-panel p-5">
              <div className="mb-3 flex items-center gap-2">
                <Hash className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{lang === "tr" ? "Görünüm" : "Display"}</h2>
              </div>
              <label className="flex items-center justify-between text-xs cursor-pointer py-1.5">
                <span>{t.showValues}</span>
                <input type="checkbox" checked={showValues} onChange={(e) => setShowValues(e.target.checked)}
                  className="accent-primary h-4 w-4" />
              </label>
              <label className="flex items-center justify-between text-xs cursor-pointer py-1.5">
                <span>{t.showAxisLabels}</span>
                <input type="checkbox" checked={showAxisLabels} onChange={(e) => setShowAxisLabels(e.target.checked)}
                  className="accent-primary h-4 w-4" />
              </label>
            </section>

            <section className="glass-panel p-5">
              <div className="mb-3 flex items-center gap-2">
                <Palette className="h-4 w-4 text-primary" />
                <h2 className="text-sm font-medium">{t.colorPalette}</h2>
              </div>
              <div className="space-y-2">
                {PALETTES.map((p, i) => (
                  <button key={p.name.en} onClick={() => setPalette(i)}
                    className={`glass-soft flex w-full items-center justify-between gap-3 p-2.5 transition ${
                      palette === i ? "border-primary/70 bg-primary/10" : "hover:bg-white/5"
                    }`}>
                    <span className="text-xs font-medium">{p.name[lang]}</span>
                    <div className="flex gap-1">
                      {p.colors.map((c) => (
                        <span key={c} className="h-4 w-4 rounded-full ring-1 ring-white/20" style={{ background: c }} />
                      ))}
                    </div>
                  </button>
                ))}
              </div>
            </section>
          </aside>
        </div>

        <footer className="mt-12 grid gap-6 md:grid-cols-3">
          <section className="glass-panel p-6">
            <div className="mb-3 flex items-center gap-2">
              <Info className="h-4 w-4 text-primary" />
              <h2 className="text-sm font-semibold">{t.purpose}</h2>
            </div>
            <p className="text-xs leading-relaxed text-muted-foreground">{t.purposeText}</p>
          </section>

          <section className="glass-panel p-6">
            <div className="mb-3 flex items-center gap-2">
              <Workflow className="h-4 w-4 text-primary" />
              <h2 className="text-sm font-semibold">{t.howTo}</h2>
            </div>
            <ol className="space-y-1.5 text-xs leading-relaxed text-muted-foreground">
              <li><span className="text-primary">1.</span> {t.step1}</li>
              <li><span className="text-primary">2.</span> {t.step2}</li>
              <li><span className="text-primary">3.</span> {t.step3}</li>
            </ol>
          </section>

          <section className="glass-panel p-6">
            <div className="mb-3 flex items-center gap-2">
              <Mail className="h-4 w-4 text-primary" />
              <h2 className="text-sm font-semibold">{t.contact}</h2>
            </div>
            <p className="mb-3 text-xs text-muted-foreground">{t.contactText}</p>
            <div className="flex flex-col gap-2">
              <a href="https://github.com/karadumanyusuf" target="_blank" rel="noopener noreferrer"
                className="glass-soft flex items-center gap-2 px-3 py-2 text-xs transition hover:text-primary">
                <Github className="h-4 w-4" />
                <span>github.com/karadumanyusuf</span>
              </a>
              <a href="https://www.linkedin.com/in/karadumanyusuf/" target="_blank" rel="noopener noreferrer"
                className="glass-soft flex items-center gap-2 px-3 py-2 text-xs transition hover:text-primary">
                <Linkedin className="h-4 w-4" />
                <span>linkedin.com/in/karadumanyusuf</span>
              </a>
            </div>
          </section>
        </footer>

        <p className="mt-8 text-center text-[11px] text-muted-foreground">
          © {new Date().getFullYear()} ChartGlass — Yusuf Karaduman
        </p>
      </div>
    </div>
  );
}

type RenderOpts = {
  showValues: boolean;
  showAxisLabels: boolean;
  valueAxisLabel: string;
  onEdit: (rowIndex: number, key: string) => void;
  seriesColors?: Record<string, string>;
  axisSize?: number;
  valueSize?: number;
};

function SizeRow({ label, value, min, max, step, onChange }: {
  label: string; value: number; min: number; max: number; step: number;
  onChange: (n: number) => void;
}) {
  return (
    <div>
      <div className="flex items-center justify-between mb-1">
        <Label className="text-xs text-muted-foreground">{label}</Label>
        <span className="text-xs font-mono text-primary">{value}</span>
      </div>
      <input type="range" min={min} max={max} step={step} value={value}
        onChange={(e) => onChange(Number(e.target.value))}
        className="w-full accent-primary" />
    </div>
  );
}

function renderChart(
  kind: ChartKind,
  data: Row[],
  cat: string,
  nums: string[],
  colors: string[],
  opts: RenderOpts,
): React.ReactElement {
  const tooltipStyle = {
    background: "oklch(0.2 0.04 265 / 0.95)",
    border: "1px solid oklch(1 0 0 / 0.15)",
    borderRadius: 12, color: "#fff", backdropFilter: "blur(12px)",
  };
  const axisColor = "oklch(0.75 0.03 260)";
  const gridColor = "oklch(1 0 0 / 0.08)";
  const axisFs = opts.axisSize ?? 12;
  const valueFs = opts.valueSize ?? 11;
  const labelStyle = { fill: "#fff", fontSize: valueFs, fontWeight: 600 };
  const colorFor = (name: string, i: number) => opts.seriesColors?.[name] ?? colors[i % colors.length];

  const xAxisLabel = opts.showAxisLabels
    ? <RLabel value={cat} offset={-4} position="insideBottom" fill={axisColor} fontSize={axisFs} />
    : null;
  const yAxisLabel = opts.showAxisLabels
    ? <RLabel value={opts.valueAxisLabel} angle={-90} position="insideLeft" fill={axisColor} fontSize={axisFs} style={{ textAnchor: "middle" }} />
    : null;

  const handleBarClick = (key: string) => (payload: { index?: number }) => {
    if (typeof payload?.index === "number") opts.onEdit(payload.index, key);
  };

  switch (kind) {
    case "bar":
      return (
        <BarChart data={data} margin={{ top: 24, right: 16, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} cursor={{ fill: "oklch(1 0 0 / 0.05)" }} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Bar key={n} dataKey={n} fill={colorFor(n, i)} radius={[6, 6, 0, 0]}
              onClick={handleBarClick(n)} cursor="pointer">
              {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
            </Bar>
          ))}
        </BarChart>
      );
    case "stackedBar":
      return (
        <BarChart data={data} margin={{ top: 24, right: 16, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Bar key={n} dataKey={n} stackId="a" fill={colorFor(n, i)}
              onClick={handleBarClick(n)} cursor="pointer">
              {opts.showValues && <LabelList dataKey={n} position="center" style={labelStyle} />}
            </Bar>
          ))}
        </BarChart>
      );
    case "horizontalBar":
      return (
        <BarChart data={data} layout="vertical" margin={{ top: 16, right: 40, left: 16, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis type="number" stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel ? <RLabel value={opts.valueAxisLabel} offset={-4} position="insideBottom" fill={axisColor} fontSize={axisFs} /> : null}</XAxis>
          <YAxis dataKey={cat} type="category" stroke={axisColor} tick={{ fontSize: axisFs }}>
            {opts.showAxisLabels && <RLabel value={cat} angle={-90} position="insideLeft" fill={axisColor} fontSize={axisFs} style={{ textAnchor: "middle" }} />}
          </YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Bar key={n} dataKey={n} fill={colorFor(n, i)} radius={[0, 6, 6, 0]}
              onClick={handleBarClick(n)} cursor="pointer">
              {opts.showValues && <LabelList dataKey={n} position="right" style={labelStyle} />}
            </Bar>
          ))}
        </BarChart>
      );
    case "line":
      return (
        <LineChart data={data} margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Line key={n} type="linear" dataKey={n} stroke={colorFor(n, i)} strokeWidth={2.5}
              dot={{ r: 4, cursor: "pointer" }} activeDot={{ r: 6, cursor: "pointer", onClick: ((_e: unknown, p: { index?: number }) => { if (typeof p?.index === "number") opts.onEdit(p.index, n); }) as unknown as React.MouseEventHandler }}>
              {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
            </Line>
          ))}
        </LineChart>
      );
    case "smoothLine":
      return (
        <LineChart data={data} margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Line key={n} type="monotone" dataKey={n} stroke={colorFor(n, i)} strokeWidth={3}
              dot={{ r: 4, cursor: "pointer" }} activeDot={{ r: 6, cursor: "pointer", onClick: ((_e: unknown, p: { index?: number }) => { if (typeof p?.index === "number") opts.onEdit(p.index, n); }) as unknown as React.MouseEventHandler }}>
              {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
            </Line>
          ))}
        </LineChart>
      );
    case "area":
      return (
        <AreaChart data={data} margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <defs>
            {nums.map((n, i) => (
              <linearGradient key={n} id={`g-${i}`} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={colorFor(n, i)} stopOpacity={0.7} />
                <stop offset="100%" stopColor={colorFor(n, i)} stopOpacity={0.05} />
              </linearGradient>
            ))}
          </defs>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Area key={n} type="monotone" dataKey={n} stroke={colorFor(n, i)} fill={`url(#g-${i})`} strokeWidth={2}>
              {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
            </Area>
          ))}
        </AreaChart>
      );
    case "stackedArea":
      return (
        <AreaChart data={data} margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) => (
            <Area key={n} type="monotone" dataKey={n} stackId="1" stroke={colorFor(n, i)} fill={colorFor(n, i)} fillOpacity={0.6}>
              {opts.showValues && <LabelList dataKey={n} position="center" style={labelStyle} />}
            </Area>
          ))}
        </AreaChart>
      );
    case "pie":
    case "donut":
      return (
        <PieChart>
          <Pie data={data} dataKey={nums[0]} nameKey={cat} cx="50%" cy="50%"
            outerRadius={150} innerRadius={kind === "donut" ? 80 : 0} paddingAngle={2}
            onClick={(_p, idx) => opts.onEdit(idx as number, nums[0])} cursor="pointer"
            label={opts.showValues ? (e: { name?: string; value?: number }) => `${e.name}: ${e.value}` : undefined}
            labelLine={opts.showValues}>
            {data.map((_, i) => <Cell key={i} fill={colors[i % colors.length]} />)}
          </Pie>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
        </PieChart>
      );
    case "radar":
      return (
        <RadarChart data={data}>
          <PolarGrid stroke={gridColor} />
          <PolarAngleAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }} />
          <PolarRadiusAxis stroke={axisColor} tick={{ fontSize: 10 }} />
          {nums.map((n, i) => (
            <Radar key={n} dataKey={n} stroke={colorFor(n, i)} fill={colorFor(n, i)} fillOpacity={0.4}>
              {opts.showValues && <LabelList dataKey={n} style={labelStyle} />}
            </Radar>
          ))}
          <Legend wrapperStyle={{ fontSize: 12 }} />
          <Tooltip contentStyle={tooltipStyle} />
        </RadarChart>
      );
    case "radial":
      return (
        <RadialBarChart data={data.map((d, i) => ({ ...d, fill: colors[i % colors.length] }))} innerRadius={30} outerRadius={150}>
          <RadialBar dataKey={nums[0]} background onClick={(_p, idx) => opts.onEdit(idx as number, nums[0])} cursor="pointer">
            {opts.showValues && <LabelList dataKey={nums[0]} position="insideStart" style={labelStyle} />}
          </RadialBar>
          <Legend wrapperStyle={{ fontSize: 12 }} />
          <Tooltip contentStyle={tooltipStyle} />
        </RadialBarChart>
      );
    case "scatter":
      return (
        <ScatterChart margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} />
          <XAxis dataKey={nums[0]} stroke={axisColor} tick={{ fontSize: axisFs }} name={nums[0]}>
            {opts.showAxisLabels && <RLabel value={nums[0]} offset={-4} position="insideBottom" fill={axisColor} fontSize={axisFs} />}
          </XAxis>
          <YAxis dataKey={nums[1] ?? nums[0]} stroke={axisColor} tick={{ fontSize: axisFs }} name={nums[1] ?? nums[0]}>
            {opts.showAxisLabels && <RLabel value={nums[1] ?? nums[0]} angle={-90} position="insideLeft" fill={axisColor} fontSize={axisFs} style={{ textAnchor: "middle" }} />}
          </YAxis>
          <Tooltip contentStyle={tooltipStyle} cursor={{ strokeDasharray: "3 3" }} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          <Scatter name={cat} data={data} fill={colors[0]} onClick={(_p, idx) => opts.onEdit(idx as number, nums[1] ?? nums[0])} cursor="pointer">
            {opts.showValues && <LabelList dataKey={nums[1] ?? nums[0]} position="top" style={labelStyle} />}
          </Scatter>
        </ScatterChart>
      );
    case "composed":
      return (
        <ComposedChart data={data} margin={{ top: 24, right: 24, left: 8, bottom: opts.showAxisLabels ? 24 : 8 }}>
          <CartesianGrid stroke={gridColor} strokeDasharray="3 3" />
          <XAxis dataKey={cat} stroke={axisColor} tick={{ fontSize: axisFs }}>{xAxisLabel}</XAxis>
          <YAxis stroke={axisColor} tick={{ fontSize: axisFs }}>{yAxisLabel}</YAxis>
          <Tooltip contentStyle={tooltipStyle} />
          <Legend wrapperStyle={{ fontSize: 12 }} />
          {nums.map((n, i) =>
            i % 2 === 0 ? (
              <Bar key={n} dataKey={n} fill={colorFor(n, i)} radius={[6, 6, 0, 0]}
                onClick={handleBarClick(n)} cursor="pointer">
                {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
              </Bar>
            ) : (
              <Line key={n} type="monotone" dataKey={n} stroke={colorFor(n, i)} strokeWidth={2.5}>
                {opts.showValues && <LabelList dataKey={n} position="top" style={labelStyle} />}
              </Line>
            )
          )}
        </ComposedChart>
      );
  }
}
