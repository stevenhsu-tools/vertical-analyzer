import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileSpreadsheet, Play, Download, AlertCircle, CheckCircle, Save, Settings, XCircle, Info, Loader2, Key, Terminal } from 'lucide-react';

// The Canvas environment injects a key here automatically if left blank. MUST be named exactly apiKey.
const apiKey = ""; 

// Dynamically load ExcelJS
const useExcelJS = () => {
  const [isLoaded, setIsLoaded] = useState(false);
  const excelRef = useRef(null);

  useEffect(() => {
    if (window.ExcelJS) {
      excelRef.current = window.ExcelJS;
      setIsLoaded(true);
      return;
    }

    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
    script.onload = () => {
      excelRef.current = window.ExcelJS;
      setIsLoaded(true);
    };
    document.body.appendChild(script);
  }, []);

  return { isLoaded, ExcelJS: excelRef.current };
};

// The 27 Verticals with Descriptions and Examples from the provided document
const VERTICAL_DEFINITIONS = [
  { name: "Aerospace", description: "Focuses on the design, manufacture, and maintenance of aircraft, spacecraft, and related systems for civil and military use.", example: "Producing commercial jet engines or regional aircraft like the Airbus A320." },
  { name: "Automotive", description: "Involved in the design, development, manufacturing, marketing, and selling of motorized vehicles.", example: "An assembly line for electric vehicles (EVs) or a plant manufacturing heavy-duty trucks." },
  { name: "Defense", description: "Dedicated to providing products and services to national security and military forces.", example: "Manufacturing armored personnel carriers, military drones, or tactical communication equipment." },
  { name: "Electricity & Utilities", description: "Companies involved in the generation, transmission, and distribution of electrical power.", example: "Managing a regional power grid or operating a solar farm." },
  { name: "Electronics", description: "Focused on the production of consumer and industrial electronic components and devices.", example: "Manufacturing printed circuit boards (PCBs) or consumer smart home devices." },
  { name: "Food & Beverage", description: "Covers the processing, packaging, and distribution of food products and drinks.", example: "A large-scale automated bottling plant or a dairy processing facility." },
  { name: "Machinery", description: "Focused on the production of industrial machines used by other businesses to create goods.", example: "Manufacturing industrial CNC lathes or heavy excavators for construction." },
  { name: "Manufacturing", description: "The broad sector involving the mechanical, physical, or chemical transformation of materials, substances, or components into new products.", example: "A factory producing furniture, toys, or general consumer goods." },
  { name: "Mining & Metals", description: "Extraction of valuable minerals or other geological materials from the earth.", example: "Copper mining, steel production, or aluminum smelting." },
  { name: "Oil & Gas", description: "Exploration, extraction, refining, transporting, and marketing of oil and gas products.", example: "An offshore drilling platform or a refinery." },
  { name: "Pharmaceutical & Healthcare", description: "Focused on medical research, drug manufacturing, and providing medical services.", example: "A pharmaceutical lab producing vaccines or a private hospital network." },
  { name: "Retail", description: "The sale of goods directly to consumers through physical stores or e-commerce.", example: "A large supermarket chain or a high-end clothing boutique." },
  { name: "Semiconductor", description: "The design and manufacture of integrated circuits (chips) for electronic devices.", example: "A 'fab' (fabrication plant) producing microprocessors for smartphones." },
  { name: "Transportation", description: "Moving people and goods via road, rail, or air.", example: "A national railway company or a global logistics and courier service." },
  { name: "Water Utilities", description: "The management of water supply, treatment, and sewage systems.", example: "A municipal water treatment plant ensuring safe drinking water." },
  { name: "Others", description: "A catch-all category for any industry or entity that does not fit into the 26 specific classifications above.", example: "An independent artist's studio or a specialized educational tutoring center." },
  { name: "Others - Association", description: "Organizations of people with a common interest or purpose, often non-profit.", example: "Trade unions, professional associations, or chambers of commerce." },
  { name: "Others - Chemical", description: "Companies that produce industrial chemicals.", example: "Fertilizer plants, polymer manufacturers, or basic chemical producers." },
  { name: "Others - Construction", description: "Building of infrastructure, residential, and commercial buildings.", example: "A commercial construction company or a civil engineering firm." },
  { name: "Others - Consulting & Business Process Service", description: "Professional services providing expert advice or managing business processes.", example: "Management consulting firms, HR services, or call centers." },
  { name: "Others - Entertainment & Leisure", description: "Sector focused on recreation, arts, entertainment, and tourism.", example: "Movie theaters, theme parks, hotels, or sports teams." },
  { name: "Others - Financial", description: "Banking, insurance, and financial asset management.", example: "Commercial banks, insurance companies, or investment firms." },
  { name: "Others - Government & Public Administration", description: "State-run services and administrative bodies.", example: "City councils, government agencies, or public departments." },
  { name: "Others - IT Service & Cybersecurity", description: "Information technology services, software development, and security.", example: "Software development houses, managed security service providers (MSSPs), or cloud hosts." },
  { name: "Others - Maritime", description: "Sea-related activities including shipping and vessel maintenance.", example: "A shipping line or a shipyard for vessel repair." },
  { name: "Others - Real Estate", description: "Buying, selling, and managing land and physical property assets.", example: "A commercial property management firm for office buildings." },
  { name: "Others - Telecommunication", description: "Providing the infrastructure and services for long-distance communication.", example: "A mobile network operator or a company installing fiber optic cables." }
];

const DEFAULT_COL_MAP = {
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
};

export default function App() {
  const { isLoaded, ExcelJS } = useExcelJS();
  const [file, setFile] = useState(null);
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [showVerticals, setShowVerticals] = useState(false);
  const [logs, setLogs] = useState([]);
  const [errorMsg, setErrorMsg] = useState('');
  
  // States for API Debugging
  // FIX: Removed the ".0" to perfectly match Google's backend model ID
  const [userApiKey, setUserApiKey] = useState('');
  const [selectedModel, setSelectedModel] = useState('gemini-3-flash-preview');

  // Auto-inject Tailwind
  useEffect(() => {
    if (!document.querySelector('script[src*="tailwindcss"]')) {
      const script = document.createElement('script');
      script.src = "https://cdn.tailwindcss.com";
      document.head.appendChild(script);
    }
  }, []);

  // Standard Column Mapping
  const [colMap, setColMap] = useState(DEFAULT_COL_MAP);

  // Display Labels for the Mapping UI
  const colLabels = {
    name: 'Company Name',
    duns: 'DUNS Number',
    country: 'Country',
    parentName: 'Parent Company',
    website: 'Website',
    currentVertical: 'Current or Should be Vertical',
    targetVertical: 'AI Assigned Vertical',
    notes: 'Analysis Notes',
    status: 'Company Status',
    industry: 'Vertical to Compare'
  };

  const addLog = (msg) => setLogs(prev => [`[${new Date().toLocaleTimeString()}] ${msg}`, ...prev].slice(0, 50));

  const handleFileUpload = async (e) => {
    setErrorMsg('');
    if (!isLoaded) return setErrorMsg("Excel engine is still loading. Please wait a moment.");
    
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
          fileHeaders = row.values.slice(1).map(h => h ? h.toString() : '');
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
      addLog(`Successfully loaded ${jsonData.length} rows from file.`);
    } catch (err) {
      setErrorMsg("Error reading Excel file. Please ensure it is a valid .xlsx file.");
    }
  };

  const callGemini = async (companyData, activeKey, targetModel, retryCount = 0) => {
    const verticalContext = VERTICAL_DEFINITIONS.map(v => `VERTICAL: "${v.name}"\nDEFINITION: ${v.description}\nEXAMPLE: ${v.example}`).join('\n\n');

    const systemPrompt = `You are an expert Business Analyst.
Task: Determine the single best 'Vertical' for the company using the strict definitions provided.

Reference Vertical List (Definitions):
--------------------------------------------------
${verticalContext}
--------------------------------------------------

Instructions:
1. CRITICAL: Research the specific company name using your knowledge to find their actual core product or service. Look closely at the Parent Company if provided.
2. SUBSIDIARY/SERVICE ARM RULE: If a company provides IT, retail, or local branch services strictly as a subsidiary of a specialized manufacturer (e.g., "Infineon Technologies IT-Services", "Stellantis & YOU"), classify it under the parent's core industry (e.g., 'Semiconductor' or 'Automotive'), NOT as 'Others - IT Service' or 'Retail'.
3. WASTE/ENVIRONMENTAL RULE: Waste management, environmental services, and recycling companies (e.g., Prezero, Urbaser) MUST be classified as 'Electricity & Utilities'.
4. BEWARE OF LEGAL SUFFIXES: Do not classify a company as "Financial" or "Others - Association" just because its name ends in "AG", "Holdings", "Group", or "LLC". (e.g., "Ronal AG" makes car wheels, so it must be "Automotive").
5. Match it to EXACTLY ONE vertical from the list above. Read the 'DEFINITION' and 'EXAMPLE' carefully.
6. Specificity Rule: If they manufacture car parts, choose 'Automotive'. If they manufacture computer chips, choose 'Semiconductor'. Only use the general 'Manufacturing' category if they don't fit a more specific one.
7. If the company status implies 'Out of business', 'Inactive', or 'M&A', verify this.
8. Provide a short note (max 15 words) justifying your choice.`;

    const userQuery = `Company Information:
- Name: ${companyData.name}
- Parent Company: ${companyData.parentName || 'N/A'}
- Country: ${companyData.country || 'N/A'}
- Website: ${companyData.website || 'N/A'}
- Industry Description: ${companyData.industry || 'N/A'}
- Status: ${companyData.status || 'Active'}`;

    try {
      const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${targetModel}:generateContent?key=${activeKey}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: userQuery }] }],
          systemInstruction: { parts: [{ text: systemPrompt }] },
          generationConfig: { 
            responseMimeType: "application/json",
            responseSchema: {
              type: "OBJECT",
              properties: {
                selected_vertical: { type: "STRING" },
                note: { type: "STRING" },
                is_out_of_business: { type: "BOOLEAN" }
              },
              required: ["selected_vertical", "note", "is_out_of_business"]
            }
          }
        })
      });

      if (!response.ok) {
        let errorDetails = "Unknown API Error";
        try {
            const errBody = await response.json();
            if (errBody?.error?.message) errorDetails = errBody.error.message;
        } catch (e) {}

        // AUTO-FALLBACK: If the model is not found (404), seamlessly drop down to 2.5 Flash
        if (response.status === 404 && targetModel !== 'gemini-2.5-flash') {
            return callGemini(companyData, activeKey, 'gemini-2.5-flash', retryCount);
        }

        // Fatal auth errors (400, 401, 403). Do not retry these.
        if (response.status === 400 || response.status === 401 || response.status === 403 || response.status === 404) {
             return { selected_vertical: "FATAL_ERROR", note: `API Error ${response.status}: ${errorDetails}`, is_out_of_business: false };
        }

        // Exponential backoff strategy for rate limits (429) or temporary server errors (500)
        if (retryCount < 4) {
           const backoffDelays = [2000, 4000, 8000, 15000];
           addLog(`[WARN] Hit rate limit (429). Pausing for ${backoffDelays[retryCount]/1000}s to refill quota...`);
           await new Promise(r => setTimeout(r, backoffDelays[retryCount]));
           return callGemini(companyData, activeKey, targetModel, retryCount + 1);
        }
        return { selected_vertical: "Error", note: `API Failed (Status ${response.status})`, is_out_of_business: false };
      }

      const data = await response.json();
      let text = data.candidates?.[0]?.content?.parts?.[0]?.text;
      
      // Safety clean for markdown blocks
      if (text) {
        text = text.replace(/```json/gi, '').replace(/```/g, '').trim();
      }
      
      return JSON.parse(text);

    } catch (error) {
      if (retryCount < 4) {
         const backoffDelays = [2000, 4000, 8000, 15000];
         await new Promise(r => setTimeout(r, backoffDelays[retryCount]));
         return callGemini(companyData, activeKey, targetModel, retryCount + 1);
      }
      return { selected_vertical: "Error", note: `Connection Failed`, is_out_of_business: false };
    }
  };

  const processRows = async () => {
    setErrorMsg('');
    
    if (rows.length === 0) {
      setErrorMsg("Please upload an Excel file first.");
      return;
    }

    if (!colMap.name || !headers.includes(colMap.name)) {
      setErrorMsg("Please correctly map the 'Company Name' column to an existing column in your file before starting.");
      return;
    }

    // Determine which key to use
    const customKey = userApiKey.trim();
    const activeKey = customKey || apiKey;
    
    // Auto-Routing: Ensure the correct model is used based on the key provided
    let safeModel = selectedModel;
    
    if (customKey && selectedModel === 'gemini-2.5-flash-preview-09-2025') {
        // Personal keys cannot access the internal Canvas model
        safeModel = 'gemini-3-flash-preview';
        addLog(`[INFO] Auto-switched to Gemini 3.0 (Internal preview model restricts personal keys)`);
    } else if (!customKey) {
        // The Canvas environment's hidden key ONLY supports 2.5 Preview
        safeModel = 'gemini-2.5-flash-preview-09-2025';
        if (selectedModel !== 'gemini-2.5-flash-preview-09-2025') {
            addLog(`[INFO] Auto-switched to Gemini 2.5 Preview (Canvas built-in key only supports this model)`);
        }
    }

    // Mask the key for debugging
    const maskedKey = customKey.length > 10 
        ? `${customKey.substring(0, 5)}...${customKey.substring(customKey.length - 4)}`
        : "Canvas Default Built-in Key";

    addLog(`[DEBUG] Active API Key: ${maskedKey}`);
    addLog(`[DEBUG] Selected Model: ${safeModel}`);
    addLog(`Starting analysis for ${rows.length} companies...`);

    setProcessing(true);
    setProgress({ current: 0, total: rows.length });
    
    const newRows = [...rows];
    const BATCH_SIZE = 2; // Keep low to avoid instant 429 Rate Limits on Free Tier
    let hasFatalError = false;

    for (let i = 0; i < newRows.length; i += BATCH_SIZE) {
      if (hasFatalError) break;

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

        const result = await callGemini(companyData, activeKey, selectedModel);

        const targetColName = colMap.targetVertical || DEFAULT_COL_MAP.targetVertical;
        const notesColName = colMap.notes || DEFAULT_COL_MAP.notes;

        if (result.selected_vertical === "FATAL_ERROR") {
            hasFatalError = true;
            newRows[globalIdx][notesColName] = result.note;
            return;
        }

        newRows[globalIdx][targetColName] = result.selected_vertical;
        newRows[globalIdx][notesColName] = result.note + (result.is_out_of_business ? " [POSSIBLE OUT OF BUSINESS]" : "");
        
        // FIX: Ensure status is treated as a string to prevent crashes if Excel reads it as a number
        const statusVal = row[colMap.status] || '';
        if (String(statusVal).toLowerCase().includes('out of business')) {
             newRows[globalIdx][notesColName] = `[CONFIRMED OUT OF BUSINESS] ${newRows[globalIdx][notesColName]}`;
        }
      });

      await Promise.all(promises);
      setProgress({ current: Math.min(i + BATCH_SIZE, newRows.length), total: newRows.length });
      setRows([...newRows]);
      
      // Delay between batches to respect free tier limit (15 Requests Per Minute)
      if (i + BATCH_SIZE < newRows.length && !hasFatalError) {
          await new Promise(r => setTimeout(r, 4000));
      }
    }

    setProcessing(false);
    
    if (hasFatalError) {
        addLog("[ERROR] Processing aborted. Invalid API Key or Model access.");
        setErrorMsg(`API Error! Look at the logs window below or the "Analysis Notes" column in the table to see the exact error message from Google.`);
    } else {
        addLog("Analysis complete! Ready for download.");
    }
  };

  const exportExcel = async () => {
    if (!isLoaded || !file) return;
    setErrorMsg('');
    try {
      addLog("Preparing final Excel file with formatting...");
      const workbook = new ExcelJS.Workbook();
      const buffer = await file.arrayBuffer();
      await workbook.xlsx.load(buffer);
      
      const worksheet = workbook.worksheets[0];
      const headerRow = worksheet.getRow(1);
      const colIndices = {};
      let maxCol = 0;
      
      headerRow.eachCell((cell, colNumber) => {
        colIndices[cell.text] = colNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      });

      // Dynamically create all mapped or default columns if they don't exist in the uploaded file
      Object.keys(DEFAULT_COL_MAP).forEach(key => {
        const colName = colMap[key] || DEFAULT_COL_MAP[key];
        if (!colIndices[colName]) {
          maxCol++;
          colIndices[colName] = maxCol;
          const newCell = headerRow.getCell(maxCol);
          newCell.value = colName;
          newCell.font = { bold: true };
        }
      });

      const targetColIndex = colIndices[colMap.targetVertical || DEFAULT_COL_MAP.targetVertical];
      const notesColIndex = colIndices[colMap.notes || DEFAULT_COL_MAP.notes];

      rows.forEach((row) => {
        const rowIndex = row._rowNumber;
        const excelRow = worksheet.getRow(rowIndex);
        
        const targetColName = colMap.targetVertical || DEFAULT_COL_MAP.targetVertical;
        const notesColName = colMap.notes || DEFAULT_COL_MAP.notes;
        const currentColName = colMap.currentVertical || DEFAULT_COL_MAP.currentVertical;

        const newVal = row[targetColName];
        const newNote = row[notesColName];
        
        if (newVal) excelRow.getCell(targetColIndex).value = newVal;
        if (newNote) excelRow.getCell(notesColIndex).value = newNote;

        const currentVal = row[currentColName];
        if (newVal && currentVal) {
          const normTarget = String(newVal).trim().toLowerCase();
          const normOriginal = String(currentVal).trim().toLowerCase();
          if (normTarget !== normOriginal) {
            // Apply Red Color if mismatch
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
      addLog(`File successfully exported: ${filename}`);
      
    } catch (err) {
      console.error(err);
      addLog(`[ERROR] Export failed: ${err.message}`);
      setErrorMsg("Error generating export file. Please try again.");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans p-6">
      <div className="max-w-full mx-auto space-y-6">
        
        {/* Header */}
        <header className="bg-white p-6 rounded-xl shadow-sm border border-slate-200 flex flex-wrap gap-4 justify-between items-center">
          <div>
            <h1 className="text-2xl font-bold text-slate-800 flex items-center gap-2">
              <CheckCircle className="text-indigo-600" />
              Vertical Verification Agent
            </h1>
            <p className="text-slate-500 mt-1">Automated validation against TXOne Vertical List</p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            
            {/* Model Dropdown */}
            <select 
              value={selectedModel}
              onChange={(e) => setSelectedModel(e.target.value)}
              className="px-3 py-2 border rounded-lg text-sm bg-slate-50 focus:ring-2 focus:ring-indigo-500 outline-none text-slate-700 font-medium cursor-pointer"
            >
              <option value="gemini-3-flash-preview">Gemini 3.0 Flash Preview</option>
            </select>

            <div className="relative">
              <Key className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
              <input 
                type="password"
                placeholder="Paste API Key Here"
                className="pl-9 pr-4 py-2 border rounded-lg text-sm w-64 focus:ring-2 focus:ring-indigo-500 outline-none"
                value={userApiKey}
                onChange={(e) => setUserApiKey(e.target.value)}
              />
            </div>
            <button 
              onClick={() => setShowVerticals(!showVerticals)}
              className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-sm font-medium flex items-center gap-2 transition-colors"
            >
              <Info size={16} />
              View Reference List
            </button>
          </div>
        </header>

        {/* Error Banner */}
        {errorMsg && (
          <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-xl flex justify-between items-center animate-in fade-in slide-in-from-top-2 shadow-sm">
            <span className="flex items-center gap-2 font-medium">
              <AlertCircle size={18} />
              {errorMsg}
            </span>
            <button onClick={() => setErrorMsg('')} className="hover:text-red-900 transition-colors">
              <XCircle size={18} />
            </button>
          </div>
        )}

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
                  {isLoaded ? (
                    <>
                      <FileSpreadsheet className="w-12 h-12 text-slate-400 mx-auto mb-3" />
                      <p className="text-slate-600 font-medium">Click to Upload Excel File</p>
                      <p className="text-xs text-slate-400 mt-1">Supports .xlsx files</p>
                    </>
                  ) : (
                    <div className="flex flex-col items-center justify-center h-full text-slate-400">
                      <Loader2 className="w-8 h-8 animate-spin mb-2" />
                      <p className="text-sm">Loading Engine...</p>
                    </div>
                  )}
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
                      <label className="text-xs font-bold text-slate-700">{colLabels[key]}</label>
                      <select 
                        value={val} 
                        onChange={(e) => setColMap({...colMap, [key]: e.target.value})}
                        className="text-xs p-2 border rounded bg-slate-50 outline-none focus:ring-1 focus:ring-indigo-500"
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
                    className="w-full py-3 px-4 rounded-lg font-medium flex items-center justify-center gap-2 transition-all bg-indigo-600 hover:bg-indigo-700 text-white shadow-md"
                   >
                    <Play size={18} />
                    Start Analysis
                   </button>
                 ) : (
                   <div className="bg-indigo-50 border border-indigo-100 p-4 rounded-lg shadow-inner">
                      <div className="flex justify-between text-sm font-medium text-indigo-900 mb-2">
                         <span className="flex items-center gap-2"><Loader2 size={14} className="animate-spin" /> Processing...</span>
                         <span>{Math.round((progress.current / progress.total) * 100)}%</span>
                      </div>
                      <div className="w-full bg-indigo-200 rounded-full h-2">
                        <div 
                          className="bg-indigo-600 h-2 rounded-full transition-all duration-300" 
                          style={{ width: `${(progress.current / progress.total) * 100}%` }} 
                        />
                      </div>
                      <div className="text-xs text-indigo-600 mt-2 text-center font-medium">
                        Analyzed {progress.current} of {progress.total} companies
                      </div>
                   </div>
                 )}

                 <button 
                   onClick={exportExcel}
                   disabled={processing || progress.current === 0}
                   className="w-full py-3 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-medium flex items-center justify-center gap-2 shadow-md disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                 >
                   <Download size={18} />
                   Export Results
                 </button>
              </div>
            )}
            
            {/* Logs Window */}
            <div className="bg-slate-900 text-slate-300 p-4 rounded-xl font-mono text-xs h-64 overflow-y-auto shadow-inner">
               <div className="flex items-center gap-2 mb-2 pb-2 border-b border-slate-700 text-slate-400 font-bold">
                 <Terminal size={14} /> System Logs
               </div>
               {logs.length === 0 && <div className="text-slate-600 italic">System ready. Waiting for input...</div>}
               {logs.map((log, i) => (
                 <div key={i} className={`py-1.5 last:border-0 border-b border-slate-800 ${log.includes('[ERROR]') ? 'text-red-400' : ''} ${log.includes('[DEBUG]') ? 'text-emerald-400' : ''}`}>
                   {log}
                 </div>
               ))}
            </div>

          </div>

          {/* Right Panel: Data Preview */}
          <div className="lg:col-span-3 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[800px]">
              <div className="p-4 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
                <h3 className="font-semibold text-slate-700">Live Data Preview</h3>
                <span className="text-xs text-slate-500 font-medium bg-white px-2 py-1 rounded border border-slate-200">
                  Showing {rows.length} rows
                </span>
              </div>
              <div className="overflow-auto flex-1 p-0">
                <table className="w-full text-sm text-left border-collapse">
                  <thead className="bg-slate-50 sticky top-0 z-10 text-xs font-semibold text-slate-500 tracking-wider">
                    <tr>
                      <th className="p-3 border-b border-slate-200">Company Name</th>
                      <th className="p-3 border-b border-slate-200">Current or Should be Vertical</th>
                      <th className="p-3 border-b border-indigo-200 bg-indigo-50 text-indigo-800">AI Assigned Vertical</th>
                      <th className="p-3 border-b border-slate-200">Analysis Notes</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {rows.map((row, idx) => {
                      const targetColName = colMap.targetVertical || DEFAULT_COL_MAP.targetVertical;
                      const currentColName = colMap.currentVertical || DEFAULT_COL_MAP.currentVertical;
                      const notesColName = colMap.notes || DEFAULT_COL_MAP.notes;

                      const aiVertical = row[targetColName];
                      const currentVertical = row[currentColName];
                      
                      // FIX: Safely convert to strings before trimming to prevent React from crashing (White Screen) 
                      // if an Excel cell accidentally contains numbers instead of text.
                      const isMismatch = aiVertical && currentVertical && String(aiVertical).trim().toLowerCase() !== String(currentVertical).trim().toLowerCase();
                      const isError = aiVertical === 'FATAL_ERROR' || aiVertical === 'Error';

                      return (
                        <tr key={idx} className="hover:bg-slate-50 transition-colors">
                          <td className="p-3 font-medium text-slate-700 max-w-[200px] truncate" title={row[colMap.name]}>
                            {row[colMap.name]}
                          </td>
                          <td className="p-3 text-slate-500 max-w-[150px] truncate" title={currentVertical}>
                            {currentVertical}
                          </td>
                          <td className={`p-3 font-medium border-l border-r border-indigo-50 bg-indigo-50/20 ${isMismatch && !isError ? 'text-red-600' : 'text-slate-700'} ${isError ? 'text-red-500 bg-red-50' : ''}`}>
                            {aiVertical || '-'}
                            {isMismatch && !isError && <span className="ml-2 text-[10px] bg-red-100 text-red-700 px-1.5 py-0.5 rounded font-bold">DIFF</span>}
                          </td>
                          <td className={`p-3 text-xs max-w-[300px] truncate ${isError ? 'text-red-600 font-medium' : 'text-slate-500'}`} title={row[notesColName]}>
                            {row[notesColName]}
                          </td>
                        </tr>
                      );
                    })}
                    {rows.length === 0 && (
                      <tr>
                        <td colSpan={4} className="p-16 text-center text-slate-400">
                          <FileSpreadsheet className="w-12 h-12 mx-auto mb-3 opacity-20" />
                          No data loaded yet. Upload an Excel file to begin previewing.
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
}
