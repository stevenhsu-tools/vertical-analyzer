import React, { useState, useCallback, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Play, Download, AlertCircle, CheckCircle, Loader2, Save, Settings, XCircle, Info } from 'lucide-react';

// We need to load ExcelJS dynamically since it's a large library
// Updated to handle Hot Module Replacement (HMR) and prevent duplicate script tags
const useExcelJS = () => {
  const [isLoaded, setIsLoaded] = useState(false);
  const excelRef = useRef(null);

  useEffect(() => {
    // 1. Check if already available on window
    if (window.ExcelJS) {
      excelRef.current = window.ExcelJS;
      setIsLoaded(true);
      return;
    }

    // 2. Check if script is already injected (by another instance or HMR)
    const existingScript = document.querySelector('script[src*="exceljs"]');
    if (existingScript) {
      const checkInterval = setInterval(() => {
        if (window.ExcelJS) {
          excelRef.current = window.ExcelJS;
          setIsLoaded(true);
          clearInterval(checkInterval);
        }
      }, 500);
      return () => clearInterval(checkInterval);
    }

    // 3. Inject script if missing
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
    script.onload = () => {
      excelRef.current = window.ExcelJS;
      setIsLoaded(true);
    };
    script.onerror = () => {
      console.error("Failed to load ExcelJS");
    };
    document.body.appendChild(script);
  }, []);

  return { isLoaded, ExcelJS: excelRef.current };
};

// The 27 Verticals with Descriptions and Examples from the provided document
const VERTICAL_DEFINITIONS = [
  { 
    name: "Aerospace", 
    description: "Focuses on the design, manufacture, and maintenance of aircraft, spacecraft, and related systems for civil and military use.", 
    example: "Producing commercial jet engines or regional aircraft like the Airbus A320." 
  },
  { 
    name: "Automotive", 
    description: "Involved in the design, development, manufacturing, marketing, and selling of motorized vehicles.", 
    example: "An assembly line for electric vehicles (EVs) or a plant manufacturing heavy-duty trucks." 
  },
  { 
    name: "Defense", 
    description: "Dedicated to providing products and services to national security and military forces.", 
    example: "Manufacturing armored personnel carriers, military drones, or tactical communication equipment." 
  },
  { 
    name: "Electricity & Utilities", 
    description: "Companies involved in the generation, transmission, and distribution of electrical power.", 
    example: "Managing a regional power grid or operating a solar farm." 
  },
  { 
    name: "Electronics", 
    description: "Focused on the production of consumer and industrial electronic components and devices.", 
    example: "Manufacturing printed circuit boards (PCBs) or consumer smart home devices." 
  },
  { 
    name: "Food & Beverage", 
    description: "Covers the processing, packaging, and distribution of food products and drinks.", 
    example: "A large-scale automated bottling plant or a dairy processing facility." 
  },
  { 
    name: "Machinery", 
    description: "Focused on the production of industrial machines used by other businesses to create goods.", 
    example: "Manufacturing industrial CNC lathes or heavy excavators for construction." 
  },
  { 
    name: "Manufacturing", 
    description: "The broad sector involving the mechanical, physical, or chemical transformation of materials, substances, or components into new products.", 
    example: "A factory producing furniture, toys, or general consumer goods." 
  },
  { 
    name: "Mining & Metals", 
    description: "Extraction of valuable minerals or other geological materials from the earth.", 
    example: "Copper mining, steel production, or aluminum smelting." 
  },
  { 
    name: "Oil & Gas", 
    description: "Exploration, extraction, refining, transporting, and marketing of oil and gas products.", 
    example: "An offshore drilling platform or a refinery." 
  },
  { 
    name: "Pharmaceutical & Healthcare", 
    description: "Focused on medical research, drug manufacturing, and providing medical services.", 
    example: "A pharmaceutical lab producing vaccines or a private hospital network." 
  },
  { 
    name: "Retail", 
    description: "The sale of goods directly to consumers through physical stores or e-commerce.", 
    example: "A large supermarket chain or a high-end clothing boutique." 
  },
  { 
    name: "Semiconductor", 
    description: "The design and manufacture of integrated circuits (chips) for electronic devices.", 
    example: "A 'fab' (fabrication plant) producing microprocessors for smartphones." 
  },
  { 
    name: "Transportation", 
    description: "Moving people and goods via road, rail, or air.", 
    example: "A national railway company or a global logistics and courier service." 
  },
  { 
    name: "Water Utilities", 
    description: "The management of water supply, treatment, and sewage systems.", 
    example: "A municipal water treatment plant ensuring safe drinking water." 
  },
  { 
    name: "Others", 
    description: "A catch-all category for any industry or entity that does not fit into the 26 specific classifications above.", 
    example: "An independent artist's studio or a specialized educational tutoring center." 
  },
  { 
    name: "Others - Association", 
    description: "Organizations of people with a common interest or purpose, often non-profit.", 
    example: "Trade unions, professional associations, or chambers of commerce." 
  },
  { 
    name: "Others - Chemical", 
    description: "Companies that produce industrial chemicals.", 
    example: "Fertilizer plants, polymer manufacturers, or basic chemical producers." 
  },
  { 
    name: "Others - Construction", 
    description: "Building of infrastructure, residential, and commercial buildings.", 
    example: "A commercial construction company or a civil engineering firm." 
  },
  { 
    name: "Others - Consulting & Business Process Service", 
    description: "Professional services providing expert advice or managing business processes.", 
    example: "Management consulting firms, HR services, or call centers." 
  },
  { 
    name: "Others - Entertainment & Leisure", 
    description: "Sector focused on recreation, arts, entertainment, and tourism.", 
    example: "Movie theaters, theme parks, hotels, or sports teams." 
  },
  { 
    name: "Others - Financial", 
    description: "Banking, insurance, and financial asset management.", 
    example: "Commercial banks, insurance companies, or investment firms." 
  },
  { 
    name: "Others - Government & Public Administration", 
    description: "State-run services and administrative bodies.", 
    example: "City councils, government agencies, or public departments." 
  },
  { 
    name: "Others - IT Service & Cybersecurity", 
    description: "Information technology services, software development, and security.", 
    example: "Software development houses, managed security service providers (MSSPs), or cloud hosts." 
  },
  { 
    name: "Others - Maritime", 
    description: "Sea-related activities including shipping and vessel maintenance.", 
    example: "A shipping line or a shipyard for vessel repair." 
  },
  { 
    name: "Others - Real Estate", 
    description: "Buying, selling, and managing land and physical property assets.", 
    example: "A commercial property management firm for office buildings." 
  },
  { 
    name: "Others - Telecommunication", 
    description: "Providing the infrastructure and services for long-distance communication.", 
    example: "A mobile network operator or a company installing fiber optic cables." 
  }
];

// Retry logic for API calls
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

const App = () => {
  const { isLoaded, ExcelJS } = useExcelJS();
  const [apiKey, setApiKey] = useState('');
  const [file, setFile] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [showVerticals, setShowVerticals] = useState(false);
  const [logs, setLogs] = useState([]);

  // Automatically inject Tailwind CSS for local styling if missing
  useEffect(() => {
    const existingScript = document.querySelector('script[src*="tailwindcss"]');
    if (!existingScript) {
      const script = document.createElement('script');
      script.src = "https://cdn.tailwindcss.com";
      document.head.appendChild(script);
    }
  }, []);

  // Column Mapping State
  const [colMap, setColMap] = useState({
    name: 'DnB_Account Name',
    duns: 'DUNS Number',
    country: 'DnB_Country',
    parentName: 'DnB_Parent Company Name_Global Level',
    website: 'DnB_Company Website',
    currentVertical: 'TXOne Targeted Vertical',
    targetVertical: 'TXOne Targeted Vertical_Steven',
    notes: 'Steven Notes',
    status: 'DnB_Company Status',
    industry: 'DnB_D&B Industry Industry'
  });

  useEffect(() => {
    if (isLoaded) {
      setLogs(prev => [`[${new Date().toLocaleTimeString()}] System Ready: Excel Engine Loaded.`, ...prev]);
    }
  }, [isLoaded]);

  const addLog = (msg) => setLogs(prev => [`[${new Date().toLocaleTimeString()}] ${msg}`, ...prev].slice(0, 50));

  const handleFileUpload = async (e) => {
    if (!isLoaded) return alert("System loading Excel engine... please wait a moment.");
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;

    try {
      const workbook = new ExcelJS.Workbook();
      const buffer = await uploadedFile.arrayBuffer();
      await workbook.xlsx.load(buffer);
      
      const worksheet = workbook.worksheets[0];
      const jsonData = [];
      let fileHeaders = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) {
          fileHeaders = row.values.slice(1);
          fileHeaders = fileHeaders.map(h => h ? h.toString() : '');
        } else {
          const rowData = {};
          fileHeaders.forEach((header, index) => {
             if (!header) return;
             let cellVal = row.getCell(index + 1).value;
             if (typeof cellVal === 'object' && cellVal !== null) {
                cellVal = cellVal.text || cellVal.result || JSON.stringify(cellVal);
             }
             rowData[header] = cellVal;
          });
          rowData._rowNumber = rowNumber;
          jsonData.push(rowData);
        }
      });

      setHeaders(fileHeaders);
      setRows(jsonData);
      setFile(uploadedFile);
      addLog(`Loaded ${jsonData.length} rows.`);
    } catch (err) {
      console.error(err);
      addLog(`Error reading file: ${err.message}`);
      alert("Error reading Excel file. Please ensure it is a valid .xlsx file.");
    }
  };

  const callGemini = async (companyData, retryCount = 0) => {
    if (!apiKey) throw new Error("API Key missing");
    
    // Construct the vertical list text with descriptions for the prompt
    const verticalContext = VERTICAL_DEFINITIONS.map(v => 
      `VERTICAL: "${v.name}"\nDEFINITION: ${v.description}\nEXAMPLE: ${v.example}`
    ).join('\n\n');

    const prompt = `
      You are an expert Business Analyst.
      
      Task: Determine the single best 'Vertical' for the company below using the strict definitions provided.
      
      Company Information:
      - Name: ${companyData.name}
      - Parent Company: ${companyData.parentName || 'N/A'}
      - Country: ${companyData.country || 'N/A'}
      - Website: ${companyData.website || 'N/A'}
      - Industry Description: ${companyData.industry || 'N/A'}
      - Status: ${companyData.status || 'Active'}

      Reference Vertical List (Definitions):
      --------------------------------------------------
      ${verticalContext}
      --------------------------------------------------

      Instructions:
      1. Research the company or use your knowledge to identify its primary business activity.
      2. Match it to EXACTLY ONE vertical from the list above. Read the 'DEFINITION' and 'EXAMPLE' carefully.
      3. For 'Manufacturing', only use it if the company makes products but doesn't fit a more specific category like 'Automotive', 'Electronics', or 'Food & Beverage'.
      4. If the company status implies 'Out of business', 'Inactive', or 'M&A', verify this.
      5. Provide a short note (max 15 words) justifying your choice.
      
      Return JSON format ONLY:
      {
        "selected_vertical": "Exact Name from list",
        "note": "Short reasoning here",
        "is_out_of_business": boolean
      }
    `;

    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key=${apiKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { responseMimeType: "application/json" }
        })
      });

      if (!response.ok) {
        if (response.status === 429 && retryCount < 3) {
           await delay(2000 * (retryCount + 1));
           return callGemini(companyData, retryCount + 1);
        }
        throw new Error(`API Error: ${response.statusText}`);
      }

      const data = await response.json();
      let text = data.candidates?.[0]?.content?.parts?.[0]?.text;
      
      // Clean up Markdown formatting that often causes "Got/Hot errors"
      if (text) {
        text = text.replace(/```json/g, '').replace(/```/g, '').trim();
      }
      
      return JSON.parse(text);
    } catch (error) {
      console.error("Gemini Error", error);
      return { selected_vertical: "Error", note: "AI Request Failed", is_out_of_business: false };
    }
  };

  const processRows = async () => {
    if (!apiKey) return alert("Please enter your Gemini API Key.");
    if (rows.length === 0) return alert("Please upload a file first.");

    setProcessing(true);
    setProgress({ current: 0, total: rows.length });
    
    const newRows = [...rows];
    const BATCH_SIZE = 3;

    for (let i = 0; i < newRows.length; i += BATCH_SIZE) {
      const batch = newRows.slice(i, i + BATCH_SIZE);
      
      const promises = batch.map(async (row, idx) => {
        const globalIdx = i + idx;
        const companyData = {
          name: row[colMap.name],
          parentName: row[colMap.parentName],
          country: row[colMap.country],
          website: row[colMap.website],
          industry: row[colMap.industry],
          status: row[colMap.status]
        };

        const result = await callGemini(companyData);

        newRows[globalIdx][colMap.targetVertical] = result.selected_vertical;
        newRows[globalIdx][colMap.notes] = result.note + (result.is_out_of_business ? " [POSSIBLE OUT OF BUSINESS]" : "");
        
        if (row[colMap.status]?.toLowerCase().includes('out of business')) {
             newRows[globalIdx][colMap.notes] = `[CONFIRMED OUT OF BUSINESS] ${newRows[globalIdx][colMap.notes]}`;
        }
      });

      await Promise.all(promises);
      setProgress({ current: Math.min(i + BATCH_SIZE, newRows.length), total: newRows.length });
      setRows([...newRows]);
    }

    setProcessing(false);
    addLog("Processing complete!");
  };

  const exportExcel = async () => {
    if (!isLoaded || !file) return;
    try {
      addLog("Preparing export...");
      const workbook = new ExcelJS.Workbook();
      const buffer = await file.arrayBuffer();
      await workbook.xlsx.load(buffer);
      
      const worksheet = workbook.worksheets[0];
      const headerRow = worksheet.getRow(1);
      const colIndices = {};
      
      headerRow.eachCell((cell, colNumber) => {
        colIndices[cell.text] = colNumber;
      });

      const targetColIndex = colIndices[colMap.targetVertical];
      const notesColIndex = colIndices[colMap.notes];

      if (!targetColIndex || !notesColIndex) {
        alert("Could not find target columns in the original file. Please check headers.");
        return;
      }

      rows.forEach((row) => {
        const rowIndex = row._rowNumber;
        const excelRow = worksheet.getRow(rowIndex);
        const newVal = row[colMap.targetVertical];
        const newNote = row[colMap.notes];
        
        if (newVal) excelRow.getCell(targetColIndex).value = newVal;
        if (newNote) excelRow.getCell(notesColIndex).value = newNote;

        const currentVal = row[colMap.currentVertical];
        if (newVal && currentVal) {
          const normTarget = String(newVal).trim().toLowerCase();
          const normOriginal = String(currentVal).trim().toLowerCase();
          if (normTarget !== normOriginal) {
            excelRow.getCell(targetColIndex).font = { color: { argb: 'FFFF0000' }, bold: true };
          }
        }
      });

      const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const originalName = file.name.replace(/\.[^/.]+$/, "");
      const filename = `${originalName}_${dateStr}.xlsx`;

      const outBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([outBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      window.URL.revokeObjectURL(url);
      addLog(`File exported as: ${filename}`);
      
    } catch (err) {
      console.error(err);
      addLog(`Error exporting file: ${err.message}`);
      alert("Error generating export file.");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6">
      <div className="max-w-full mx-auto space-y-6">
        
        {/* Header */}
        <header className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold text-slate-800 flex items-center gap-2">
              <CheckCircle className="text-indigo-600" />
              Vertical Verification Agent
            </h1>
            <p className="text-slate-500 mt-1">Automated validation against TXOne Vertical List (27 Items)</p>
          </div>
          <div className="flex items-center gap-4">
            <input 
              type="password"
              placeholder="Enter Gemini API Key"
              className="px-4 py-2 border rounded-lg text-sm w-64 focus:ring-2 focus:ring-indigo-500 outline-none"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
            />
            <button 
              onClick={() => setShowVerticals(!showVerticals)}
              className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-sm font-medium flex items-center gap-2 transition-colors"
            >
              <Info size={16} />
              View Verticals
            </button>
          </div>
        </header>

        {/* Vertical List Drawer */}
        {showVerticals && (
          <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 animate-in fade-in slide-in-from-top-4">
             <div className="flex justify-between items-center mb-4">
                <h3 className="font-semibold flex items-center gap-2">
                  <Info size={18} className="text-indigo-600"/>
                  Vertical Definitions Reference ({VERTICAL_DEFINITIONS.length})
                </h3>
                <button onClick={() => setShowVerticals(false)}><XCircle size={20} className="text-slate-400 hover:text-red-500"/></button>
             </div>
             <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 h-96 overflow-y-auto custom-scrollbar pr-2">
                {VERTICAL_DEFINITIONS.map((v, i) => (
                  <div key={i} className="p-4 border border-slate-100 rounded-lg bg-slate-50 hover:border-indigo-200 transition-colors">
                    <div className="font-bold text-indigo-900 mb-1">{v.name}</div>
                    <div className="text-xs text-slate-600 mb-2">{v.description}</div>
                    <div className="text-xs text-slate-400 italic border-t border-slate-200 pt-2">Ex: {v.example}</div>
                  </div>
                ))}
             </div>
          </div>
        )}

        {/* Main Workspace */}
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          
          {/* Left Panel: Controls */}
          <div className="lg:col-span-1 space-y-6">
            
            {/* Upload Area */}
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 text-center">
              {!file ? (
                <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 hover:bg-slate-50 transition-colors relative">
                  <input 
                    type="file" 
                    accept=".xlsx" 
                    onChange={handleFileUpload}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                  />
                  <FileSpreadsheet className="w-12 h-12 text-slate-400 mx-auto mb-3" />
                  <p className="text-slate-600 font-medium">Click to Upload Excel File</p>
                  <p className="text-xs text-slate-400 mt-1">Supports .xlsx files</p>
                </div>
              ) : (
                <div className="flex items-center justify-between bg-indigo-50 p-4 rounded-lg border border-indigo-100">
                  <div className="flex items-center gap-3">
                    <FileSpreadsheet className="text-indigo-600" size={20} />
                    <div className="text-left">
                      <p className="font-medium text-sm text-indigo-900 truncate max-w-[150px]">{file.name}</p>
                      <p className="text-xs text-indigo-600">{rows.length} rows loaded</p>
                    </div>
                  </div>
                  <button onClick={() => {setFile(null); setRows([]);}} className="text-indigo-400 hover:text-indigo-700">
                    <XCircle size={18} />
                  </button>
                </div>
              )}
            </div>

            {/* Column Mapping */}
            {file && (
              <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
                <h3 className="font-semibold text-sm mb-4">Column Mapping</h3>
                <div className="space-y-3">
                  <div className="text-xs text-slate-500 mb-2">Ensure these match your Excel headers:</div>
                  {Object.entries(colMap).map(([key, val]) => (
                    <div key={key} className="flex flex-col gap-1">
                      <label className="text-xs font-medium text-slate-700 uppercase">{key.replace(/([A-Z])/g, ' $1').trim()}</label>
                      <select 
                        value={val} 
                        onChange={(e) => setColMap({...colMap, [key]: e.target.value})}
                        className="text-xs p-2 border rounded bg-slate-50"
                      >
                         <option value="">-- Select Column --</option>
                         {headers.map((h, i) => <option key={i} value={h}>{h}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Actions */}
            {file && (
              <div className="space-y-3">
                 {!processing ? (
                   <button 
                    onClick={processRows}
                    disabled={!apiKey}
                    className={`w-full py-3 px-4 rounded-lg font-medium flex items-center justify-center gap-2 transition-all ${
                      apiKey ? 'bg-indigo-600 hover:bg-indigo-700 text-white shadow-md' : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                    }`}
                   >
                    <Play size={18} />
                    Start Analysis
                   </button>
                 ) : (
                   <div className="bg-indigo-50 border border-indigo-100 p-4 rounded-lg">
                      <div className="flex justify-between text-sm font-medium text-indigo-900 mb-2">
                         <span>Processing...</span>
                         <span>{Math.round((progress.current / progress.total) * 100)}%</span>
                      </div>
                      <div className="w-full bg-indigo-200 rounded-full h-2">
                        <div 
                          className="bg-indigo-600 h-2 rounded-full transition-all duration-300" 
                          style={{ width: `${(progress.current / progress.total) * 100}%` }} 
                        />
                      </div>
                      <div className="text-xs text-indigo-600 mt-2 text-center">
                        Analyzed {progress.current} of {progress.total} companies
                      </div>
                   </div>
                 )}

                 <button 
                   onClick={exportExcel}
                   disabled={processing || rows.length === 0}
                   className="w-full py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-medium flex items-center justify-center gap-2 shadow-md disabled:opacity-50 disabled:cursor-not-allowed"
                 >
                   <Download size={18} />
                   Export Results
                 </button>
              </div>
            )}
            
            {/* Logs */}
            <div className="bg-slate-900 text-slate-300 p-4 rounded-xl font-mono text-xs h-64 overflow-y-auto custom-scrollbar">
               {logs.length === 0 && <div className="text-slate-600 italic">System ready. Waiting for input...</div>}
               {logs.map((log, i) => (
                 <div key={i} className="border-b border-slate-800 py-1 last:border-0">{log}</div>
               ))}
            </div>

          </div>

          {/* Right Panel: Data Preview */}
          <div className="lg:col-span-3 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[800px]">
              <div className="p-4 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
                <h3 className="font-semibold text-slate-700">Live Data Preview</h3>
                <span className="text-xs text-slate-500">Showing all {rows.length} rows</span>
              </div>
              <div className="overflow-auto flex-1 p-0">
                <table className="w-full text-sm text-left border-collapse">
                  <thead className="bg-slate-50 sticky top-0 z-10 text-xs font-semibold text-slate-500 uppercase tracking-wider">
                    <tr>
                      <th className="p-3 border-b">Company Name</th>
                      <th className="p-3 border-b">Current Vertical</th>
                      <th className="p-3 border-b bg-indigo-50 text-indigo-700 border-indigo-100">AI Vertical (Steven)</th>
                      <th className="p-3 border-b">Notes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {rows.map((row, idx) => {
                      const aiVertical = row[colMap.targetVertical];
                      const currentVertical = row[colMap.currentVertical];
                      const isMismatch = aiVertical && currentVertical && aiVertical.trim().toLowerCase() !== currentVertical.trim().toLowerCase();

                      return (
                        <tr key={idx} className="hover:bg-slate-50 transition-colors">
                          <td className="p-3 font-medium text-slate-700 max-w-[200px] truncate" title={row[colMap.name]}>
                            {row[colMap.name]}
                          </td>
                          <td className="p-3 text-slate-500 max-w-[150px] truncate">
                            {currentVertical}
                          </td>
                          <td className={`p-3 font-medium border-l border-r border-indigo-50 bg-indigo-50/30 ${isMismatch ? 'text-red-600' : 'text-slate-700'}`}>
                            {aiVertical || '-'}
                            {isMismatch && <span className="ml-2 text-[10px] bg-red-100 text-red-700 px-1 py-0.5 rounded">DIFF</span>}
                          </td>
                          <td className="p-3 text-slate-500 text-xs max-w-[250px] truncate" title={row[colMap.notes]}>
                            {row[colMap.notes]}
                          </td>
                        </tr>
                      );
                    })}
                    {rows.length === 0 && (
                      <tr>
                        <td colSpan={4} className="p-10 text-center text-slate-400">
                          No data loaded yet. Upload an Excel file to begin.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;