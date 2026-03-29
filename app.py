/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { 
  Upload, 
  FileType, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  Trash2,
  FileText,
  Save,
  History,
  RotateCcw,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface StationData {
  station: string;
  columnNum: number;
  feederNum: number;
  height: string;
  arms: string | number;
  lon: number;
  lat: number;
  details: string;
  isDuplicate?: boolean;
}

const STORAGE_KEY = 'kmz_extractor_session';
const ITEMS_PER_PAGE = 25;

const NATURAL_SORT_REGEX = /(\d+)/;
// ... (rest of the helper functions remain same)

const naturalSortKey = (s: string) => {
  if (!s) return [];
  return s.split(NATURAL_SORT_REGEX).map(text => {
    const num = parseInt(text, 10);
    return isNaN(num) ? text.toLowerCase() : num;
  });
};

const compareNatural = (a: string, b: string) => {
  const keyA = naturalSortKey(a);
  const keyB = naturalSortKey(b);
  const len = Math.max(keyA.length, keyB.length);
  
  for (let i = 0; i < len; i++) {
    if (keyA[i] === undefined) return -1;
    if (keyB[i] === undefined) return 1;
    if (keyA[i] < keyB[i]) return -1;
    if (keyA[i] > keyB[i]) return 1;
  }
  return 0;
};

export default function App() {
  const [files, setFiles] = useState<File[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [processedData, setProcessedData] = useState<StationData[]>([]);
  const [sessionStatus, setSessionStatus] = useState<'saved' | 'loaded' | 'cleared' | null>(null);
  const [currentPage, setCurrentPage] = useState(1);

  const totalPages = Math.ceil(processedData.length / ITEMS_PER_PAGE);
  const paginatedData = processedData.slice(
    (currentPage - 1) * ITEMS_PER_PAGE,
    currentPage * ITEMS_PER_PAGE
  );

  const saveSession = () => {
    if (processedData.length === 0) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(processedData));
    showStatus('saved');
  };

  const loadSession = () => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const data = JSON.parse(saved);
        setProcessedData(data);
        setCurrentPage(1);
        showStatus('loaded');
      } catch (e) {
        setError("فشل تحميل الجلسة السابقة. البيانات قد تكون تالفة.");
      }
    } else {
      setError("لا توجد جلسة محفوظة.");
    }
  };

  const clearResults = () => {
    setProcessedData([]);
    setCurrentPage(1);
    showStatus('cleared');
  };

  const showStatus = (status: 'saved' | 'loaded' | 'cleared') => {
    setSessionStatus(status);
    setTimeout(() => setSessionStatus(null), 3000);
  };

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const newFiles = Array.from(e.target.files).filter(f => f.name.toLowerCase().endsWith('.kmz'));
      setFiles(prev => [...prev, ...newFiles]);
      setError(null);
    }
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const processFiles = async () => {
    if (files.length === 0) return;
    
    setIsProcessing(true);
    setError(null);
    const allData: StationData[] = [];

    try {
      for (const file of files) {
        const zip = new JSZip();
        const contents = await zip.loadAsync(file);
        const kmlFile = Object.keys(contents.files).find(name => name.endsWith('.kml'));
        
        if (!kmlFile) {
          console.warn(`No KML found in ${file.name}`);
          continue;
        }

        const kmlText = await contents.files[kmlFile].async('string');
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(kmlText, 'text/xml');
        const placemarks = xmlDoc.getElementsByTagName('Placemark');

        for (let i = 0; i < placemarks.length; i++) {
          const pm = placemarks[i];
          
          // 1. Name Extraction
          const nameNode = pm.getElementsByTagName('name')[0];
          const fullName = nameNode?.textContent?.trim() || "";
          
          // Arabic/Numeric Station Code regex
          const stMatch = fullName.match(/(\d+[\u0600-\u06FF]+|[\u0600-\u06FF]+\d+)/);
          const stationCode = stMatch ? stMatch[1] : "غير محدد";

          const cleanName = fullName.replace(stationCode, "").trim();
          const nameNums = cleanName.match(/\d+/g) || [];
          
          let columnNum = 0;
          let feederNum = 0;

          if (nameNums.length >= 2) {
            columnNum = parseInt(nameNums[0], 10);
            feederNum = parseInt(nameNums[1], 10);
          } else if (nameNums.length === 1) {
            columnNum = parseInt(nameNums[0], 10);
          }

          // 2. Description Extraction
          const descNode = pm.getElementsByTagName('description')[0];
          const descText = descNode?.textContent || "";
          
          // Extended data values
          const valueNodes = pm.getElementsByTagName('value');
          let extVals = "";
          for (let j = 0; j < valueNodes.length; j++) {
            extVals += " " + (valueNodes[j].textContent || "");
          }
          
          const techInfo = (descText + " " + extVals).trim();
          let valHeight = "";
          let valArms: string | number = "";

          const patternMatch = techInfo.match(/(\d+)[/-](\d+)/);
          if (patternMatch) {
            valHeight = patternMatch[1];
            valArms = patternMatch[2];
          } else {
            const hSearch = techInfo.match(/\b(12|10|8|6|5)\b/);
            if (hSearch) valHeight = hSearch[1];
          }

          const techLower = techInfo.toLowerCase();
          if (techLower.includes("هاي") || techLower.includes("mast")) {
            valHeight = "هاي ماست";
            valArms = 6;
          } else if (techLower.includes("جداري")) {
            valHeight = "جداري";
            valArms = 1;
          }

          // 3. Coordinates
          const coordNode = pm.getElementsByTagName('coordinates')[0];
          let lat = 0, lon = 0;
          if (coordNode?.textContent) {
            const parts = coordNode.textContent.trim().split(',');
            if (parts.length >= 2) {
              lon = parseFloat(parts[0]);
              lat = parseFloat(parts[1]);
            }
          }

          const allTxt = (fullName + " " + techInfo).toLowerCase();
          let detail = "";
          if (allTxt.includes("مفقود")) detail = "مفقود";
          else if (allTxt.includes("مغروز")) detail = "مغروز";

          allData.push({
            station: stationCode,
            columnNum,
            feederNum,
            height: valHeight,
            arms: valArms,
            lon: Number(lon.toFixed(5)),
            lat: Number(lat.toFixed(5)),
            details: detail
          });
        }
      }

      // Sorting: Station -> Feeder -> Column
      allData.sort((a, b) => {
        const stComp = compareNatural(a.station, b.station);
        if (stComp !== 0) return stComp;
        if (a.feederNum !== b.feederNum) return a.feederNum - b.feederNum;
        return a.columnNum - b.columnNum;
      });

      // Mark duplicates (same station, feeder, and column number)
      const counts = new Map<string, number>();
      allData.forEach(d => {
        // Only count as duplicate if columnNum is not 0 (or handle 0 if needed)
        const key = `${d.station}|${d.feederNum}|${d.columnNum}`;
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      
      allData.forEach(d => {
        const key = `${d.station}|${d.feederNum}|${d.columnNum}`;
        d.isDuplicate = (counts.get(key) || 0) > 1;
      });

      setProcessedData(allData);
      setCurrentPage(1);
      setIsProcessing(false);
    } catch (err) {
      console.error(err);
      setError("حدث خطأ أثناء معالجة الملفات. تأكد من أنها ملفات KMZ صالحة.");
      setIsProcessing(false);
    }
  };

  const downloadExcel = async () => {
    if (processedData.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Report', { views: [{ rightToLeft: true }] });

    const columns = [
      { header: "المحطة", key: "station", width: 20 },
      { header: "رقم العمود", key: "columnNum", width: 15 },
      { header: "رقم الفيدر", key: "feederNum", width: 15 },
      { header: "طول العمود", key: "height", width: 15 },
      { header: "الذراع", key: "arms", width: 15 },
      { header: "الاحداثيات x", key: "lon", width: 15 },
      { header: "الاحداثيات y", key: "lat", width: 15 },
      { header: "التفاصيل", key: "details", width: 20 },
    ];

    worksheet.columns = columns;

    // Header Style
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FF000000' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    let lastStation = "";
    let currentRow = 2;

    processedData.forEach((data) => {
      // Add empty row between different stations
      if (lastStation && data.station !== lastStation) {
        currentRow++;
      }

      const row = worksheet.getRow(currentRow);
      row.values = [
        data.station,
        data.columnNum,
        data.feederNum,
        data.height,
        data.arms,
        data.lon,
        data.lat,
        data.details
      ];

      const isRed = data.details === "مفقود" || data.details === "مغروز";
      const isBlue = data.isDuplicate;

      row.eachCell((cell, colNumber) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        if (isBlue) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } }; // Blue
          cell.font = { color: { argb: 'FFFFFFFF' } }; // White text for contrast
        } else if (isRed) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } }; // Red
        } else if (colNumber === 1) { // Station column
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7F7F7F' } };
          cell.font = { bold: true };
        }

        // Coordinates format
        if (colNumber === 6 || colNumber === 7) {
          cell.numFmt = '0.00000';
        }
      });

      lastStation = data.station;
      currentRow++;
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Lighting_Report_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-[#f8f9fa] text-[#1a1a1a] font-sans selection:bg-orange-100" dir="rtl">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-orange-500 rounded-lg">
              <FileType className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">مستخرج بيانات المحطات</h1>
              <p className="text-xs text-gray-500 font-medium">معالجة KMZ وفصل الفيدر</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <span className="text-xs font-mono bg-gray-100 px-2 py-1 rounded text-gray-600">v2.1.0</span>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-10">
          
          {/* Left Column: Upload & Actions */}
          <div className="lg:col-span-5 space-y-6">
            <div className="bg-white rounded-2xl border border-gray-200 p-8 shadow-sm">
              <h2 className="text-lg font-semibold mb-6 flex items-center gap-2">
                <Upload className="w-5 h-5 text-orange-500" />
                رفع الملفات
              </h2>
              
              <label className="relative group cursor-pointer block">
                <input 
                  type="file" 
                  multiple 
                  accept=".kmz" 
                  onChange={onFileChange}
                  className="hidden"
                />
                <div className="border-2 border-dashed border-gray-200 group-hover:border-orange-400 transition-colors rounded-xl p-10 text-center bg-gray-50/50">
                  <div className="w-12 h-12 bg-white rounded-full shadow-sm flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform">
                    <Upload className="w-6 h-6 text-gray-400 group-hover:text-orange-500" />
                  </div>
                  <p className="text-sm font-medium text-gray-600">اسحب ملفات KMZ هنا أو انقر للاختيار</p>
                  <p className="text-xs text-gray-400 mt-2">يدعم ملفات KMZ المتعددة</p>
                </div>
              </label>

              <AnimatePresence>
                {files.length > 0 && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -10 }}
                    className="mt-8 space-y-3"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-gray-400 uppercase tracking-wider">الملفات المختارة ({files.length})</span>
                      <button 
                        onClick={() => setFiles([])}
                        className="text-xs text-red-500 hover:underline font-medium"
                      >
                        مسح الكل
                      </button>
                    </div>
                    <div className="max-h-48 overflow-y-auto pr-2 space-y-2 custom-scrollbar">
                      {files.map((file, idx) => (
                        <div key={idx} className="flex items-center justify-between p-3 bg-gray-50 rounded-lg border border-gray-100 group">
                          <div className="flex items-center gap-3 overflow-hidden">
                            <FileText className="w-4 h-4 text-orange-400 flex-shrink-0" />
                            <span className="text-sm font-medium truncate">{file.name}</span>
                          </div>
                          <button 
                            onClick={() => removeFile(idx)}
                            className="p-1 hover:bg-red-50 rounded text-gray-400 hover:text-red-500 transition-colors"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>

              <div className="mt-8 pt-6 border-t border-gray-100">
                <button
                  disabled={files.length === 0 || isProcessing}
                  onClick={processFiles}
                  className={`w-full py-4 rounded-xl font-bold text-sm transition-all flex items-center justify-center gap-2 shadow-lg shadow-orange-500/10
                    ${files.length === 0 || isProcessing 
                      ? 'bg-gray-100 text-gray-400 cursor-not-allowed shadow-none' 
                      : 'bg-orange-500 text-white hover:bg-orange-600 active:scale-[0.98]'}`}
                >
                  {isProcessing ? (
                    <>
                      <Loader2 className="w-5 h-5 animate-spin" />
                      جاري المعالجة...
                    </>
                  ) : (
                    <>
                      <CheckCircle2 className="w-5 h-5" />
                      بدء المعالجة
                    </>
                  )}
                </button>
              </div>
            </div>

            {error && (
              <div className="bg-red-50 border border-red-100 p-4 rounded-xl flex items-start gap-3 text-red-600">
                <AlertCircle className="w-5 h-5 flex-shrink-0 mt-0.5" />
                <p className="text-sm font-medium">{error}</p>
              </div>
            )}
          </div>

          {/* Right Column: Results & Export */}
          <div className="lg:col-span-7">
            <div className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden h-full flex flex-col">
              <div className="p-6 border-b border-gray-100 flex flex-wrap items-center justify-between gap-4 bg-gray-50/30">
                <div className="flex items-center gap-4">
                  <h2 className="text-lg font-semibold flex items-center gap-2">
                    <FileText className="w-5 h-5 text-gray-400" />
                    النتائج المستخرجة
                  </h2>
                  <AnimatePresence>
                    {sessionStatus && (
                      <motion.span
                        initial={{ opacity: 0, x: 10 }}
                        animate={{ opacity: 1, x: 0 }}
                        exit={{ opacity: 0 }}
                        className={`text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-tighter
                          ${sessionStatus === 'saved' ? 'bg-blue-100 text-blue-600' : 
                            sessionStatus === 'loaded' ? 'bg-green-100 text-green-600' : 
                            'bg-gray-100 text-gray-600'}`}
                      >
                        {sessionStatus === 'saved' ? 'تم الحفظ' : 
                         sessionStatus === 'loaded' ? 'تم التحميل' : 
                         'تم المسح'}
                      </motion.span>
                    )}
                  </AnimatePresence>
                </div>

                <div className="flex items-center gap-2">
                  {processedData.length > 0 && (
                    <>
                      <button
                        onClick={saveSession}
                        title="حفظ الجلسة الحالية"
                        className="p-2 text-gray-500 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors border border-gray-200 bg-white"
                      >
                        <Save className="w-4 h-4" />
                      </button>
                      <button
                        onClick={clearResults}
                        title="مسح النتائج"
                        className="p-2 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors border border-gray-200 bg-white"
                      >
                        <RotateCcw className="w-4 h-4" />
                      </button>
                    </>
                  )}
                  <button
                    onClick={loadSession}
                    title="تحميل آخر جلسة"
                    className="flex items-center gap-2 px-3 py-2 text-gray-600 hover:text-orange-600 hover:bg-orange-50 rounded-lg text-xs font-bold transition-colors border border-gray-200 bg-white"
                  >
                    <History className="w-4 h-4" />
                    تحميل الجلسة
                  </button>
                  {processedData.length > 0 && (
                    <button
                      onClick={downloadExcel}
                      className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg text-sm font-bold hover:bg-green-700 transition-colors shadow-lg shadow-green-600/10"
                    >
                      <Download className="w-4 h-4" />
                      تحميل (Excel)
                    </button>
                  )}
                </div>
              </div>

              <div className="flex-1 overflow-auto custom-scrollbar">
                {processedData.length > 0 ? (
                  <table className="w-full text-right border-collapse">
                    <thead className="sticky top-0 bg-white z-[5]">
                      <tr className="bg-gray-50/80 backdrop-blur-sm">
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">المحطة</th>
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">العمود</th>
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">الفيدر</th>
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">الطول</th>
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">الذراع</th>
                        <th className="p-4 text-xs font-bold text-gray-500 border-b border-gray-100">التفاصيل</th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedData.map((row, idx) => {
                        const isRed = row.details === "مفقود" || row.details === "مغروز";
                        const isBlue = row.isDuplicate;
                        
                        return (
                          <tr 
                            key={idx} 
                            className={`border-b border-gray-50 hover:bg-gray-50/50 transition-colors 
                              ${isBlue ? 'bg-blue-500 text-white' : isRed ? 'bg-red-50/30' : ''}`}
                          >
                            <td className={`p-4 text-sm font-bold ${isBlue ? 'text-white' : 'text-gray-700'}`}>{row.station}</td>
                            <td className="p-4 text-sm font-medium">{row.columnNum}</td>
                            <td className="p-4 text-sm font-medium">{row.feederNum}</td>
                            <td className="p-4 text-sm font-medium">{row.height}</td>
                            <td className="p-4 text-sm font-medium">{row.arms}</td>
                            <td className={`p-4 text-xs font-bold ${isBlue ? 'text-blue-100' : isRed ? 'text-red-600' : 'text-gray-400'}`}>
                              {row.details || '-'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full py-20 text-gray-400">
                    <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mb-4">
                      <FileText className="w-8 h-8 opacity-20" />
                    </div>
                    <p className="text-sm font-medium">لا توجد بيانات حالياً</p>
                    <p className="text-xs mt-1">قم برفع ملفات KMZ ومعالجتها للبدء</p>
                  </div>
                )}
              </div>
              
              {processedData.length > 0 && (
                <div className="p-4 bg-gray-50 border-t border-gray-100 flex flex-col gap-4">
                  {/* Pagination Controls */}
                  {totalPages > 1 && (
                    <div className="flex items-center justify-center gap-2" dir="ltr">
                      <button
                        onClick={() => setCurrentPage(1)}
                        disabled={currentPage === 1}
                        className="p-1 rounded hover:bg-gray-200 disabled:opacity-30 transition-colors"
                        title="الصفحة الأولى"
                      >
                        <ChevronsLeft className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                        disabled={currentPage === 1}
                        className="p-1 rounded hover:bg-gray-200 disabled:opacity-30 transition-colors"
                        title="الصفحة السابقة"
                      >
                        <ChevronLeft className="w-4 h-4" />
                      </button>
                      
                      <div className="flex items-center gap-1 mx-2">
                        <span className="text-xs font-bold text-gray-600">صفحة {currentPage} من {totalPages}</span>
                      </div>

                      <button
                        onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                        disabled={currentPage === totalPages}
                        className="p-1 rounded hover:bg-gray-200 disabled:opacity-30 transition-colors"
                        title="الصفحة التالية"
                      >
                        <ChevronRight className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => setCurrentPage(totalPages)}
                        disabled={currentPage === totalPages}
                        className="p-1 rounded hover:bg-gray-200 disabled:opacity-30 transition-colors"
                        title="الصفحة الأخيرة"
                      >
                        <ChevronsRight className="w-4 h-4" />
                      </button>
                    </div>
                  )}

                    <div className="flex justify-between items-center text-xs text-gray-500">
                    <div className="flex items-center gap-4">
                      <span>إجمالي السجلات: {processedData.length}</span>
                      <div className="flex items-center gap-2">
                        <div className="w-3 h-3 bg-blue-500 rounded-sm"></div>
                        <span>أرقام مكررة</span>
                      </div>
                      <div className="flex items-center gap-2">
                        <div className="w-3 h-3 bg-red-500/20 border border-red-200 rounded-sm"></div>
                        <span>مفقود/مغروز</span>
                      </div>
                    </div>
                    <span className="flex items-center gap-1">
                      <CheckCircle2 className="w-3 h-3 text-green-500" />
                      تمت المعالجة بنجاح
                    </span>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </main>

      <style>{`
        .custom-scrollbar::-webkit-scrollbar {
          width: 6px;
          height: 6px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #e2e8f0;
          border-radius: 10px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover {
          background: #cbd5e1;
        }
      `}</style>
    </div>
  );
}
