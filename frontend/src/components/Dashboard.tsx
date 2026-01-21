import React, { useState, useEffect } from 'react';
import axios from 'axios';
import {
    Upload, FileText, Download, Settings as SettingsIcon,
    AlertCircle, Loader2, CheckCircle, ChevronDown,
    Trash2, Printer, Edit2, Save, Database
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
    return twMerge(clsx(inputs));
}

// --- Helper Functions ---
function formatDate(val: any): string {
    if (!val) return "";
    const s = val.toString();
    // Remove T00:00:00 or other time parts if it looks like an ISO date
    if (s.includes('T') && s.length >= 10 && !isNaN(Date.parse(s))) {
        return s.split('T')[0];
    }
    return s;
}

// --- Types ---
interface TestResult {
    Element?: string;
    Parameter?: string;
    Spec: string;
    heat1_val: string;
    heat2_val?: string;
    Hide?: boolean;
}

interface MTCData {
    invoice_no: string;
    date: string;
    qty: string;
    part_details: string;
    heat1: string;
    heat2: string;
    chemistry: TestResult[];
    mechanical: TestResult[];
    grade: string;
}

const API_BASE = (import.meta.env.VITE_API_URL as string) || "http://localhost:8000/api";

const ELEMENT_ALIASES: Record<string, string[]> = {
    "Carbon": ["C [%]", "C%", "CARBON"],
    "Silicon": ["Si [%]", "Si%", "SILICON"],
    "Manganese": ["Mn [%]", "Mn%", "MANGANESE"],
    "Phosphorus": ["P [%]", "P%", "PHOSPHORUS"],
    "Sulphur": ["S [%]", "S%", "SULPHUR"],
    "Chromium": ["Cr [%]", "Cr%", "CHROMIUM", "cr%"],
    "Copper": ["Cu [%]", "Cu%", "COPPER", "cu%"],
    "Tin": ["Sn [%]", "Sn%", "TIN", "sn%", "sn"],
    "Magnesium": ["Mg [%]", "Mg", "Mg%", "MAGNESIUM", "mg %"],
    "Nickle": ["Ni [%]", "Ni", "NICKEL", "NICKLE", "Ni%"],
    "Moly": ["Mo [%]", "Mo", "MOLYBDENUM", "MOLY", "Mo%"],
};

const STANDARD_ELEMENTS = ["Carbon", "Silicon", "Manganese", "Phosphorus", "Sulphur", "Chromium", "Copper", "Tin", "Magnesium", "Nickle", "Moly"];
const STANDARD_MECHANICAL = ["2.1 Hardness", "2.2 Tensile Strength", "2.3 Yield Strength", "2.4 % Of Elongation"];

export default function Dashboard() {
    const [activeTab, setActiveTab] = useState<'viewer' | 'mtc' | 'grade-master'>('viewer');
    const [data, setData] = useState<any[]>([]);
    const [availableHeats, setAvailableHeats] = useState<string[]>([]);
    const [mtcData, setMtcData] = useState<MTCData>({
        invoice_no: "",
        date: new Date().toLocaleDateString(),
        qty: "no's",
        part_details: "Part Name: ; Part No: ; Date code:",
        heat1: "",
        heat2: "",
        chemistry: [
            { Element: "Carbon", Spec: "3.25 - 4.10%", heat1_val: "" },
            { Element: "Silicon", Spec: "1.80 - 3.00%", heat1_val: "" },
            { Element: "Manganese", Spec: "0.1 - 1.00%", heat1_val: "" },
            { Element: "Phosphorus", Spec: "0.050% Max", heat1_val: "" },
            { Element: "Sulphur", Spec: "0.035% Max", heat1_val: "" },
            { Element: "Chromium", Spec: "-", heat1_val: "" },
            { Element: "Copper", Spec: "-", heat1_val: "" },
            { Element: "Tin", Spec: "-", heat1_val: "" },
            { Element: "Magnesium", Spec: "0.035 - 0.060%", heat1_val: "" },
            { Element: "Nickle", Spec: "-", heat1_val: "" },
            { Element: "Moly", Spec: "-", heat1_val: "" },
        ],
        mechanical: [
            { Parameter: "2.1 Hardness", Spec: "156-217 HB", heat1_val: "197/197/197/207/207 BHN" },
            { Parameter: "2.2 Tensile Strength", Spec: "Min 450 Mpa", heat1_val: "515.28 Mpa" },
            { Parameter: "2.3 Yield Strength", Spec: "Min 295 Mpa", heat1_val: "326.02 Mpa" },
            { Parameter: "2.4 % Of Elongation", Spec: "Min 12 %", heat1_val: "14.00%" }
        ],
        grade: ""
    });
    const [settings, setSettings] = useState<any>(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [deduplicate, setDeduplicate] = useState(false);
    const [savedFormats, setSavedFormats] = useState<Record<string, string[]>>({});
    const [selectedFormat, setSelectedFormat] = useState("");
    const [masterColumns, setMasterColumns] = useState<string[]>([]);
    const [visibleColumns, setVisibleColumns] = useState<string[]>([]);
    const [saveFormatSuccess, setSaveFormatSuccess] = useState(false);
    const [availableSampleIds, setAvailableSampleIds] = useState<string[]>([]);
    const [sampleIdFilter, setSampleIdFilter] = useState("");
    const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>({});
    const [filterSearch, setFilterSearch] = useState("");
    const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
    const [gradeMaster, setGradeMaster] = useState<Record<string, { chemistry: any[], mechanical: any[] }>>({});
    const [selectedGradeMaster, setSelectedGradeMaster] = useState<string>("");
    const filterRef = React.useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (filterRef.current && !filterRef.current.contains(event.target as Node)) {
                setActiveFilterCol(null);
            }
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    useEffect(() => {
        fetchSettings();
        fetchFormats();
        fetchGradeMaster();
    }, []);

    const fetchGradeMaster = async () => {
        try {
            const res = await axios.get(`${API_BASE}/grade-master`);
            setGradeMaster(res.data);
        } catch (err) {
            console.error("Failed to load grade master");
        }
    };

    const fetchFormats = async () => {
        try {
            const res = await axios.get(`${API_BASE}/formats`);
            setSavedFormats({ ...res.data });
        } catch (err) {
            console.error("Failed to load formats");
        }
    };

    const fetchSettings = async () => {
        try {
            const res = await axios.get(`${API_BASE}/settings`);
            setSettings(res.data);
        } catch (err) {
            setError("Failed to load settings");
        }
    };

    const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        setError(null);
        const formData = new FormData();
        formData.append('file', file);

        try {
            const res = await axios.post(`${API_BASE}/analyze`, formData);
            const analyzedData = res.data.data;
            setData(analyzedData);
            setAvailableHeats(res.data.heats || []);
            setSelectedFormat(""); // Reset format selection on new upload
            setColumnFilters({}); // Reset filters on new upload

            // Extract Sample IDs
            if (analyzedData.length > 0) {
                const sIds = Array.from(new Set(analyzedData.map((d: any) => (d['Sample Id'] || d['SAMPLE ID'] || d['sample_id'] || '').toString()).filter((s: any) => s))) as string[];
                setAvailableSampleIds(sIds);
            }
            // Set master columns from backend provided list
            if (res.data.columns && res.data.columns.length > 0) {
                const backendCols = res.data.columns;
                const newMaster = ['S.No', ...backendCols.filter((k: string) => k !== 'S.No')];
                setMasterColumns(newMaster);

                // Persist visibility: if already have selections, keep them. 
                // Otherwise, show all.
                if (visibleColumns.length === 0) {
                    setVisibleColumns(newMaster);
                } else {
                    // Filter out columns that don't exist in new master, 
                    // and maybe add new mandatory columns?
                    // actually, just keep current selections.
                }
            }
            // Auto-select first heats if available
            if (res.data.heats && res.data.heats.length > 0) {
                const h1 = res.data.heats[0];
                const h2 = res.data.heats.length > 1 ? res.data.heats[1] : res.data.heats[0];
                autoMapHeats(h1, h2, res.data.data);
            }
        } catch (err: any) {
            const msg = err.response?.data?.detail || "Error analyzing report. Please check the file format.";
            setError(msg);
        }
    };

    const handleSampleIdFilterChange = (id: string) => {
        setSampleIdFilter(id);
        if (!id) return;

        // Auto-map Grade if found in the first row associated with this Sample Id
        const sampleRow = data.find(d => (d['Sample Id'] || d['SAMPLE ID'] || d['sample_id'] || '').toString().trim() === id.trim());
        if (sampleRow) {
            // Priority 1: Direct matches for "Grade" or "Cast Grade"
            // Priority 2: "Spec" or "Reference"
            // Priority 3: Anything containing "Grade"
            const gradeAliases = [
                "Grade", "GRADE", "Grade No", "Cast Grade", "Grade / Spec",
                "Reference", "Ref / Grade", "Material Grade", "Spec / Grade", "MATERIAL", "SPEC", "Standard"
            ];

            // Try to find the BEST column match
            let foundGradeVal = "";
            let foundCol = "";

            // First pass: try to find an exact alias match or very close
            for (const alias of gradeAliases) {
                const col = Object.keys(sampleRow).find(k => k.trim().toUpperCase() === alias.toUpperCase());
                if (col && sampleRow[col]) {
                    foundGradeVal = sampleRow[col].toString().trim();
                    foundCol = col;
                    break;
                }
            }

            // Second pass: if nothing found, try partial match
            if (!foundGradeVal) {
                const col = Object.keys(sampleRow).find(k =>
                    gradeAliases.some(alias => k.toUpperCase().includes(alias.toUpperCase()))
                );
                if (col && sampleRow[col]) {
                    foundGradeVal = sampleRow[col].toString().trim();
                    foundCol = col;
                }
            }

            if (foundGradeVal) {
                console.log(`Found grade '${foundGradeVal}' for Sample ID '${id}' from column '${foundCol}'`);
                setMtcData(prev => ({ ...prev, grade: foundGradeVal }));
            } else {
                console.log(`No grade found for Sample ID '${id}'. Columns available: ${Object.keys(sampleRow).join(', ')}`);
            }
        }
    };

    const autoMapHeats = (h1: string, h2: string, currentData?: any[]) => {
        const activeData = currentData || filteredData;
        const row1 = activeData.find(d => (d['Heat No'] || d['HEAT NO'] || d['heat_no'] || '').toString() === h1);
        const row2 = activeData.find(d => (d['Heat No'] || d['HEAT NO'] || d['heat_no'] || '').toString() === h2);

        const newChem = mtcData.chemistry.map(item => {
            const aliases = ELEMENT_ALIASES[item.Element || ""] || [];
            let v1 = "";
            let v2 = "";
            if (row1) {
                const col = Object.keys(row1).find(k => aliases.includes(k));
                if (col) v1 = row1[col]?.toString() || "";
            }
            if (row2) {
                const col = Object.keys(row2).find(k => aliases.includes(k));
                if (col) v2 = row2[col]?.toString() || "";
            }
            return { ...item, heat1_val: v1, heat2_val: v2 };
        });

        const newMech = mtcData.mechanical.map(m => {
            if (!row1 || !m.Parameter) return m;
            // Mechanical mapping strategy: try to find column that matches parameter name
            const col = Object.keys(row1).find(k => k.toUpperCase().includes(m.Parameter!.toUpperCase()));
            return col ? { ...m, heat1_val: row1[col]?.toString() || "" } : m;
        });

        setMtcData(prev => ({ ...prev, heat1: h1, heat2: h2, chemistry: newChem, mechanical: newMech }));
    };

    // Auto-sync mapping when filters change
    useEffect(() => {
        if (filteredData.length > 0) {
            const heatsInView = Array.from(new Set(filteredData.map(d => (d['Heat No'] || d['HEAT NO'] || d['heat_no'] || '').toString()).filter(v => v)));
            if (heatsInView.length > 0) {
                let h1 = mtcData.heat1;
                let h2 = mtcData.heat2;
                if (!h1 || !heatsInView.includes(h1)) h1 = heatsInView[0];
                if (!h2 || !heatsInView.includes(h2)) h2 = heatsInView.length > 1 ? heatsInView[1] : heatsInView[0];
                autoMapHeats(h1, h2, filteredData);
            }
        }
    }, [columnFilters, data]);

    // Auto-apply specs when grade changes
    useEffect(() => {
        const currentGrade = mtcData.grade?.trim();
        if (!currentGrade || !gradeMaster[currentGrade]) return;

        const masterEntry = gradeMaster[currentGrade];
        const savedChem = masterEntry.chemistry || [];
        const savedMech = masterEntry.mechanical || [];

        setMtcData(prev => {
            let nextChem = [...prev.chemistry];
            let nextMech = [...prev.mechanical];
            let changed = false;

            // Apply Chemistry Specs
            savedChem.forEach((s: any) => {
                const idx = nextChem.findIndex(c => c.Element?.trim().toUpperCase() === s.Element?.trim().toUpperCase());
                if (idx !== -1 && nextChem[idx].Spec !== s.Spec) {
                    nextChem[idx] = { ...nextChem[idx], Spec: s.Spec };
                    changed = true;
                }
            });

            // Apply Mechanical Standards
            savedMech.forEach((s: any) => {
                const idx = nextMech.findIndex(m => m.Parameter?.trim().toUpperCase() === s.Parameter?.trim().toUpperCase());
                if (idx !== -1 && nextMech[idx].Spec !== s.Spec) {
                    nextMech[idx] = { ...nextMech[idx], Spec: s.Spec };
                    changed = true;
                }
            });

            if (changed) return { ...prev, chemistry: nextChem, mechanical: nextMech };
            return prev;
        });
    }, [mtcData.grade, gradeMaster]);

    const handleSaveToGradeMaster = async () => {
        if (!mtcData.grade) {
            alert("Please enter a Grade first");
            return;
        }
        try {
            const specs = {
                chemistry: mtcData.chemistry.map(c => ({ Element: c.Element, Spec: c.Spec })),
                mechanical: mtcData.mechanical.map(m => ({ Parameter: m.Parameter, Spec: m.Spec }))
            };
            await axios.post(`${API_BASE}/grade-master`, { grade: mtcData.grade, specs });
            await fetchGradeMaster();
            alert(`Full specifications (Chemistry & Mechanical) for Grade "${mtcData.grade}" saved to Grade Master!`);
        } catch (err) {
            setError("Failed to save to Grade Master");
        }
    };

    const handleSaveFormat = async () => {
        let name = prompt("Enter a name for this format view:");
        if (!name) return;
        name = name.trim();
        if (!name) return;
        try {
            await axios.post(`${API_BASE}/formats`, { name, columns: visibleColumns });
            await fetchFormats();
            setSelectedFormat(name);
            setSaveFormatSuccess(true);
            setTimeout(() => setSaveFormatSuccess(false), 2000);
        } catch (err) {
            setError("Failed to save format");
        }
    };

    const handleLoadFormat = (name: string) => {
        if (!name) {
            setSelectedFormat("");
            return;
        }
        const cols = savedFormats[name];
        if (cols) {
            setSelectedFormat(name);
            // setMasterColumns(cols); // REMOVED: Do not overwrite master list
            setVisibleColumns(cols);
            setColumnFilters({});
        }
    };

    const handleDeleteFormat = async (name: string, e: React.MouseEvent) => {
        e.stopPropagation();
        if (!confirm(`Are you sure you want to delete "${name}"?`)) return;
        try {
            await axios.delete(`${API_BASE}/formats/${encodeURIComponent(name.trim())}`);
            await fetchFormats();
            if (selectedFormat === name) setSelectedFormat("");
        } catch (err) {
            console.error("Delete failed", err);
            setError("Failed to delete format");
            await fetchFormats(); // Refresh to clear stale items
        }
    };

    const handleRenameFormat = async (name: string, e: React.MouseEvent) => {
        e.stopPropagation();
        let new_name = prompt(`Rename "${name}" to:`, name);
        if (!new_name) return;
        new_name = new_name.trim();
        if (!new_name || new_name === name) return;
        try {
            await axios.put(`${API_BASE}/formats/${encodeURIComponent(name.trim())}`, { new_name });
            await fetchFormats();
            if (selectedFormat === name) setSelectedFormat(new_name);
        } catch (err) {
            console.error("Rename failed", err);
            setError("Failed to rename format");
            await fetchFormats();
        }
    };

    const toggleColumn = (col: string) => {
        setVisibleColumns(prev =>
            prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
        );
    };

    const filteredData = data.filter(row => {
        return Object.entries(columnFilters).every(([col, selectedValues]) => {
            if (!selectedValues) return true;
            const val = (row[col] ?? "").toString();
            return selectedValues.includes(val);
        });
    });

    const getUniqueValues = (col: string) => {
        const vals = data.map(row => (row[col] ?? "").toString());
        return Array.from(new Set(vals)).filter(v => v !== "").sort();
    };

    const toggleFilterValue = (col: string, val: string) => {
        const current = columnFilters[col] || []; // Initialize with empty array if no filter exists
        const next = current.includes(val)
            ? current.filter(v => v !== val)
            : [...current, val];
        setColumnFilters({ ...columnFilters, [col]: next });
    };

    const clearFilter = (col: string) => {
        const next = { ...columnFilters };
        delete next[col];
        setColumnFilters(next);
    };

    const deleteElement = (index: number) => {
        const next = mtcData.chemistry.filter((_, i) => i !== index);
        setMtcData({ ...mtcData, chemistry: next });
    };

    const addChemistryRow = () => {
        const next = [...mtcData.chemistry, { Element: "New Element", Spec: "-", heat1_val: "", heat2_val: "" }];
        setMtcData({ ...mtcData, chemistry: next });
    };

    const handleSelectAll = (checked: boolean) => {
        if (checked) {
            setVisibleColumns([...masterColumns]);
        } else {
            setVisibleColumns([]);
        }
    };

    const handleDownload = async () => {
        setLoading(true);
        setError(null);
        try {
            const downloadSettings = settings || {
                mtc_template_path: "Final correct.xlsx"
            };
            const response = await axios.post(`${API_BASE}/generate-excel`, {
                data: mtcData,
                settings: downloadSettings
            }, {
                responseType: 'blob'
            });

            const url = window.URL.createObjectURL(new Blob([response.data]));
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', `MTC_Report_${mtcData.invoice_no || 'Export'}.xlsx`);
            document.body.appendChild(link);
            link.click();
            link.remove();
        } catch (err) {
            setError("Failed to generate Excel report.");
        } finally {
            setLoading(false);
        }
    };
    const gradesInRange = Array.from(new Set(data.flatMap((row: any) => {
        const gradeAliases = [
            "Grade", "GRADE", "Grade No", "Cast Grade", "Grade / Spec",
            "Reference", "Ref / Grade", "Material Grade", "Spec / Grade", "MATERIAL", "SPEC", "Standard"
        ];
        return gradeAliases.map(alias => row[alias] || row[alias.toUpperCase()] || "").filter(v => v);
    })))
        .map(v => v.toString().trim())
        .filter(v => v !== "");

    return (
        <div className="min-h-screen bg-[#0e1117] text-white flex font-sans">
            {/* Sidebar - Matching User Screenshot */}
            <aside className="w-80 bg-[#262730] h-screen fixed left-0 top-0 overflow-y-auto p-4 border-r border-gray-700">
                <div className="space-y-6">
                    <nav className="space-y-2 no-print">
                        <button onClick={() => setActiveTab('viewer')}
                            className={clsx("w-full flex items-center gap-4 px-6 py-4 rounded-2xl transition-all duration-300 font-bold tracking-tight shadow-lg active:scale-95",
                                activeTab === 'viewer' ? "bg-red-600 text-white shadow-red-900/40" : "bg-[#262730] text-gray-100 hover:bg-[#31333f] border border-gray-700/50")}>
                            <FileText className="w-5 h-5" />
                            Report Viewer
                        </button>
                        <button onClick={() => setActiveTab('mtc')}
                            className={clsx("w-full flex items-center gap-4 px-6 py-4 rounded-2xl transition-all duration-300 font-bold tracking-tight shadow-lg active:scale-95",
                                activeTab === 'mtc' ? "bg-red-600 text-white shadow-red-900/40" : "bg-[#262730] text-gray-100 hover:bg-[#31333f] border border-gray-700/50")}>
                            <Download className="w-5 h-5" />
                            MTC Production
                        </button>
                    </nav>
                    <div>
                        <h3 className="text-sm font-semibold uppercase tracking-wider text-gray-400 mb-2 font-mono">Report Settings</h3>
                        <div className="bg-[#31333f] p-3 rounded-lg flex items-center justify-between cursor-pointer hover:bg-[#3d3f4b]">
                            <span className="text-sm">Formatting Settings</span>
                            <SettingsIcon className="w-4 h-4" />
                        </div>
                        <label className="flex items-center gap-3 mt-4 cursor-pointer group">
                            <input
                                type="checkbox"
                                className="w-4 h-4 rounded bg-gray-700 border-gray-600 text-brand-blue"
                                checked={deduplicate}
                                onChange={e => setDeduplicate(e.target.checked)}
                            />
                            <span className="text-sm group-hover:text-white text-gray-300 transition-colors">Clean & Deduplicate (Max 2)</span>
                        </label>
                    </div>

                    <div>
                        <h3 className="text-sm font-semibold uppercase tracking-wider text-gray-400 mb-2 font-mono">Saved Formats</h3>
                        <div
                            className="bg-[#31333f] p-3 rounded-lg flex items-center justify-between cursor-pointer hover:bg-[#3d3f4b] mb-4 transition-all group border border-gray-700/50"
                            onClick={handleSaveFormat}
                        >
                            <span className="text-sm group-hover:text-white font-bold">
                                {saveFormatSuccess ? "Format Saved!" : "Save Current View"}
                            </span>
                            <CheckCircle className={cn(
                                "w-4 h-4 text-green-500 transition-all",
                                saveFormatSuccess ? "opacity-100 scale-110" : "opacity-0 group-hover:opacity-100"
                            )} />
                        </div>

                        <div className="space-y-1 max-h-60 overflow-y-auto custom-scrollbar pr-1">
                            {Object.keys(savedFormats).length === 0 ? (
                                <p className="text-[10px] text-gray-500 italic px-2">No saved formats yet</p>
                            ) : (
                                Object.keys(savedFormats).map(name => (
                                    <div
                                        key={name}
                                        onClick={() => handleLoadFormat(name)}
                                        className={cn(
                                            "flex items-center justify-between p-2 pl-4 rounded-xl cursor-pointer transition-all group hover:bg-[#3d3f4b]",
                                            selectedFormat === name ? "bg-red-600/20 border border-red-500/30" : "bg-transparent"
                                        )}
                                    >
                                        <span className={cn(
                                            "text-xs truncate transition-colors",
                                            selectedFormat === name ? "text-red-400 font-bold" : "text-gray-400 group-hover:text-gray-200"
                                        )}>
                                            {name}
                                        </span>
                                        <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                            <button
                                                onClick={(e) => handleRenameFormat(name, e)}
                                                className="p-1.5 hover:bg-white/10 rounded-lg text-gray-400 hover:text-blue-400 transition-colors"
                                                title="Rename"
                                            >
                                                <Edit2 className="w-3 h-3" />
                                            </button>
                                            <button
                                                onClick={(e) => handleDeleteFormat(name, e)}
                                                className="p-1.5 hover:bg-white/10 rounded-lg text-gray-400 hover:text-red-400 transition-colors"
                                                title="Delete"
                                            >
                                                <Trash2 className="w-3 h-3" />
                                            </button>
                                        </div>
                                    </div>
                                ))
                            )}
                        </div>
                    </div>

                    <div>
                        <label className="flex items-center gap-3 mb-4 cursor-pointer group">
                            <input
                                type="checkbox"
                                className="w-4 h-4 rounded bg-gray-700 border-gray-600 accent-red-500"
                                checked={visibleColumns.length > 0 && masterColumns.length > 0 && visibleColumns.length >= masterColumns.length}
                                onChange={(e) => handleSelectAll(e.target.checked)}
                            />
                            <span className="text-sm group-hover:text-white transition-colors">Select All Columns</span>
                        </label>
                        <h3 className="text-sm font-semibold uppercase tracking-wider text-gray-400 mb-2 font-mono">Toggle columns visibility:</h3>
                        <div className="border border-gray-700 rounded-lg overflow-hidden bg-[#0e1117] max-h-60 overflow-y-auto custom-scrollbar">
                            <table className="w-full text-[10px]">
                                <thead className="sticky top-0 z-10 bg-[#31333f]">
                                    <tr className="text-gray-400 shadow-sm">
                                        <th className="p-2 text-left font-bold uppercase">Column</th>
                                        <th className="p-2 text-center w-12 font-bold uppercase">Show?</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {masterColumns.map(col => (
                                        <tr key={col} className="border-t border-gray-800 hover:bg-gray-900/50 group transition-colors">
                                            <td className="p-2 text-gray-300 group-hover:text-white">{col}</td>
                                            <td className="p-2 text-center">
                                                <input
                                                    type="checkbox"
                                                    checked={visibleColumns.includes(col)}
                                                    onChange={() => toggleColumn(col)}
                                                    className="w-3 h-3 rounded bg-gray-800 border-gray-600 text-red-500 cursor-pointer accent-red-500"
                                                />
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                        <p className="text-[10px] text-gray-500 mt-2 italic px-1">* Changes reflect instantly in the sheet</p>
                    </div>
                </div>
            </aside>

            {/* Main Content */}
            <main className="flex-1 ml-80 p-12">
                <div className="flex justify-between items-start mb-10">
                    <h1 className="text-4xl font-bold text-white tracking-tight">Spectro Report Viewer</h1>
                    <div className="flex items-center gap-4">
                        <input type="file" className="hidden" onChange={handleFileUpload} id="spectro-upload-header" />
                        <label htmlFor="spectro-upload-header" className="px-6 py-2 bg-red-600 text-white text-xs font-black uppercase tracking-widest rounded-xl hover:bg-red-700 cursor-pointer shadow-lg shadow-red-900/20 transition-all active:scale-95 flex items-center gap-2">
                            <Upload className="w-4 h-4" />
                            Upload Report
                        </label>
                    </div>
                </div>

                {/* Tabs - Red Underline Style */}
                <div className="flex gap-10 border-b border-gray-800 mb-10">
                    <button
                        className={cn("pb-4 px-1 transition-all relative text-sm font-bold uppercase tracking-widest", activeTab === 'viewer' ? "text-red-500" : "text-gray-500 hover:text-white")}
                        onClick={() => setActiveTab('viewer')}
                    >
                        Report Viewer
                        {activeTab === 'viewer' && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-red-500 shadow-lg shadow-red-500/20" />}
                    </button>
                    <button
                        className={cn("pb-4 px-1 transition-all relative text-sm font-bold uppercase tracking-widest", activeTab === 'grade-master' ? "text-red-500" : "text-gray-500 hover:text-white")}
                        onClick={() => setActiveTab('grade-master')}
                    >
                        Grade Master
                        {activeTab === 'grade-master' && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-red-500 shadow-lg shadow-red-500/20" />}
                    </button>
                    <button
                        className={cn("pb-4 px-1 transition-all relative text-sm font-bold uppercase tracking-widest", activeTab === 'mtc' ? "text-red-500" : "text-gray-500 hover:text-white")}
                        onClick={() => setActiveTab('mtc')}
                    >
                        MTC
                        {activeTab === 'mtc' && <div className="absolute bottom-0 left-0 w-full h-[3px] bg-red-500 shadow-lg shadow-red-500/20" />}
                    </button>
                </div>

                {activeTab === 'viewer' ? (
                    <div className="space-y-8 animate-in fade-in duration-300">
                        {data.length === 0 ? (
                            <section className="bg-[#262730] p-8 rounded-3xl border border-gray-700 shadow-2xl">
                                <h2 className="text-2xl font-bold mb-6 text-white">Daily Report</h2>
                                <div className="border-2 border-dashed border-gray-600 p-16 rounded-3xl flex flex-col items-center justify-center bg-[#0e1117] hover:border-red-500 transition-all group cursor-pointer relative overflow-hidden">
                                    <div className="absolute inset-0 bg-red-500/5 opacity-0 group-hover:opacity-100 transition-opacity" />
                                    <Upload className="w-16 h-16 text-gray-600 mb-6 group-hover:text-red-400 group-hover:-translate-y-1 transition-all duration-300" />
                                    <div className="text-center relative z-10">
                                        <p className="text-gray-200 text-lg font-semibold">Drag and drop file here</p>
                                        <p className="text-sm text-gray-500 mt-2">Limit 200MB per file • XLSX, CSV</p>
                                    </div>
                                    <input type="file" className="hidden" onChange={handleFileUpload} id="spectro-upload" />
                                    <label htmlFor="spectro-upload" className="mt-10 px-10 py-3 bg-red-600 text-white font-bold rounded-2xl hover:bg-red-700 cursor-pointer shadow-xl shadow-red-900/20 transition-all active:scale-95 z-10">Browse files</label>
                                </div>
                            </section>
                        ) : (
                            <section className="bg-[#262730] p-6 rounded-3xl border border-gray-700 shadow-2xl overflow-hidden animate-in slide-in-from-bottom-8">
                                <div className="flex items-center gap-6 mb-6">
                                    <div className="flex items-center gap-3 text-[10px] text-blue-400 font-mono bg-blue-500/10 px-4 py-2 rounded-full border border-blue-500/20 uppercase tracking-widest">
                                        <CheckCircle className="w-4 h-4" />
                                        {filteredData.length === data.length
                                            ? `${data.length} entries analyzed`
                                            : `Showing ${filteredData.length} of ${data.length} entries`}
                                    </div>
                                    <h3 className="text-xl font-bold">Report Data: Loaded File</h3>
                                </div>
                                <div className="overflow-x-auto rounded-xl border border-gray-700 bg-[#0e1117] shadow-inner max-h-[600px] overflow-y-auto custom-scrollbar">
                                    <table className="w-full text-[11px] text-left border-collapse">
                                        <thead className="sticky top-0 z-20 bg-[#31333f]">
                                            <tr>
                                                {masterColumns.filter(c => visibleColumns.includes(c)).map(col => (
                                                    <th key={col} className="p-3 border-b border-gray-700 whitespace-nowrap font-black text-gray-400 uppercase tracking-tighter hover:bg-gray-800 transition-colors cursor-default relative">
                                                        <div className="flex items-center justify-between gap-2">
                                                            <span>{col}</span>
                                                            <button
                                                                onClick={() => {
                                                                    setActiveFilterCol(activeFilterCol === col ? null : col);
                                                                    setFilterSearch("");
                                                                }}
                                                                className={cn(
                                                                    "p-1 hover:bg-white/10 rounded transition-colors",
                                                                    columnFilters[col]?.length > 0 ? "text-red-500" : "text-gray-600"
                                                                )}
                                                            >
                                                                <ChevronDown className="w-3 h-3" />
                                                            </button>
                                                        </div>
                                                        {activeFilterCol === col && (
                                                            <div ref={filterRef} className="fixed mt-8 w-72 bg-[#1f2128] border border-gray-700 rounded-2xl shadow-[0_20px_50px_rgba(0,0,0,0.8)] z-[9999] p-5 normal-case font-medium animate-in fade-in zoom-in-95 duration-200 ring-1 ring-white/20">
                                                                <input
                                                                    autoFocus
                                                                    className="w-full bg-[#0e1117] border border-gray-700 p-3 rounded-xl text-xs outline-none focus:border-red-600 mb-4 text-white placeholder-gray-500 shadow-inner transition-all"
                                                                    placeholder={`Search ${col}...`}
                                                                    value={filterSearch}
                                                                    onChange={e => setFilterSearch(e.target.value)}
                                                                />
                                                                <div className="mb-3 pb-3 border-b border-gray-800">
                                                                    <label className="flex items-center gap-3 p-2 hover:bg-white/5 rounded-lg cursor-pointer text-xs transition-colors">
                                                                        <input type="checkbox"
                                                                            className="w-4 h-4 rounded border-gray-600 bg-transparent text-red-600 focus:ring-red-600 transition-all"
                                                                            checked={!columnFilters[col] || columnFilters[col].length === getUniqueValues(col).length || columnFilters[col].length === 0}
                                                                            onChange={() => {
                                                                                if (!columnFilters[col] || columnFilters[col].length === getUniqueValues(col).length) {
                                                                                    setColumnFilters({ ...columnFilters, [col]: [] });
                                                                                } else {
                                                                                    setColumnFilters({ ...columnFilters, [col]: getUniqueValues(col) });
                                                                                }
                                                                            }}
                                                                        />
                                                                        <span className="font-bold text-gray-200">Select All</span>
                                                                    </label>
                                                                </div>
                                                                <div className="max-h-64 overflow-y-auto space-y-1 custom-scrollbar pr-1">
                                                                    {getUniqueValues(col)
                                                                        .filter(v => v.toLowerCase().includes(filterSearch.toLowerCase()))
                                                                        .map(val => (
                                                                            <label key={val} className="flex items-center gap-3 p-2 hover:bg-white/5 rounded-lg cursor-pointer text-xs transition-colors group/item">
                                                                                <input type="checkbox"
                                                                                    className="w-4 h-4 rounded border-gray-600 bg-transparent text-red-600 focus:ring-red-600 transition-all"
                                                                                    checked={!columnFilters[col] || columnFilters[col].includes(val)}
                                                                                    onChange={() => toggleFilterValue(col, val)}
                                                                                />
                                                                                <span className="truncate text-gray-400 group-hover/item:text-gray-200 transition-colors uppercase font-mono tracking-tight">{val}</span>
                                                                            </label>
                                                                        ))}
                                                                </div>
                                                                <div className="flex justify-between mt-5 pt-4 border-t border-gray-800">
                                                                    <button onClick={() => clearFilter(col)} className="text-[10px] uppercase font-black text-gray-500 hover:text-white transition-colors">Clear All</button>
                                                                    <button onClick={() => setActiveFilterCol(null)} className="text-[10px] uppercase font-black text-red-500 hover:text-red-400 transition-colors">Apply & Close</button>
                                                                </div>
                                                            </div>
                                                        )}
                                                    </th>
                                                ))}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filteredData.map((row, i) => (
                                                <tr key={i} className="hover:bg-red-500/10 transition-colors border-b border-gray-800/80 group odd:bg-[#1a1c23]/30 even:bg-transparent">
                                                    {masterColumns.filter(c => visibleColumns.includes(c)).map((col, j) => {
                                                        const val = col === 'S.No' ? (i + 1) : row[col];
                                                        return (
                                                            <td key={j} className="p-4 text-gray-100 whitespace-nowrap font-bold text-xs group-hover:text-white transition-colors">
                                                                {formatDate(val)}
                                                            </td>
                                                        );
                                                    })}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </section>
                        )}
                    </div>
                ) : activeTab === 'mtc' ? (
                    <div className="space-y-12 animate-in fade-in duration-500 max-w-5xl">
                        <h2 className="text-3xl font-black border-l-8 border-red-600 pl-6 mb-12 uppercase tracking-tighter">MTC Production Studio</h2>
                        {/* Sec 1: Details */}
                        <div className="space-y-8">
                            <h3 className="text-xl font-bold text-gray-200 flex items-center gap-3">
                                <span className="bg-red-600 text-white w-10 h-10 rounded-2xl flex items-center justify-center text-lg font-black shadow-lg shadow-red-900/20 italic">1</span>
                                Report Metadata
                            </h3>
                            <div className="grid grid-cols-2 gap-12 bg-[#262730] p-10 rounded-3xl border border-gray-700 shadow-2xl">
                                <div className="space-y-8">
                                    <div className="group">
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3 group-focus-within:text-red-500 transition-colors">Part Details Line</label>
                                        <input className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm focus:border-red-600 outline-none transition-all shadow-inner"
                                            value={mtcData.part_details} onChange={e => setMtcData({ ...mtcData, part_details: e.target.value })} />
                                    </div>
                                    <div>
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3">Customer Name</label>
                                        <input className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm opacity-50 cursor-not-allowed italic" readOnly defaultValue="SUNDRAM FASTENERS LTD" />
                                    </div>
                                    <div className="group">
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3 group-focus-within:text-red-500 transition-colors focus:border-red-600">Reference / Grade</label>
                                        <input className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm transition-all focus:border-red-600 outline-none"
                                            value={mtcData.grade} onChange={e => setMtcData({ ...mtcData, grade: e.target.value })} />
                                    </div>
                                </div>
                                <div className="space-y-8">
                                    <div className="group">
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3 group-focus-within:text-red-500 transition-colors">Invoice No</label>
                                        <input className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm placeholder-gray-800 transition-all focus:border-red-600 outline-none"
                                            value={mtcData.invoice_no} onChange={e => setMtcData({ ...mtcData, invoice_no: e.target.value })} placeholder="e.g. INV-2024-001" />
                                    </div>
                                    <div className="group">
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3 group-focus-within:text-red-500 transition-colors">Dispatch Date</label>
                                        <input type="date" className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm transition-all focus:border-red-600 outline-none text-gray-200"
                                            value={mtcData.date} onChange={e => setMtcData({ ...mtcData, date: e.target.value })} />
                                    </div>
                                    <div className="group">
                                        <label className="block text-[11px] font-black text-gray-500 uppercase tracking-[0.2em] mb-3 group-focus-within:text-red-500 transition-colors">Quantity</label>
                                        <input className="w-full bg-[#0e1117] border border-gray-700 p-4 rounded-2xl text-sm transition-all focus:border-red-600 outline-none"
                                            value={mtcData.qty} onChange={e => setMtcData({ ...mtcData, qty: e.target.value })} />
                                    </div>
                                </div>
                            </div>
                        </div>

                        {/* Sec 2: Results */}
                        <div className="space-y-8">
                            <h3 className="text-xl font-bold text-gray-200 flex items-center gap-3">
                                <span className="bg-red-600 text-white w-10 h-10 rounded-2xl flex items-center justify-center text-lg font-black shadow-lg shadow-red-900/20 italic">2</span>
                                Chemical Analysis
                            </h3>
                            <div className="bg-[#262730] p-10 rounded-3xl border border-gray-700 shadow-2xl space-y-10">
                                <div className="bg-[#0e1117] p-8 rounded-2xl border border-gray-700 shadow-inner">
                                    <div className="flex justify-between items-center mb-6 pb-4 border-b border-gray-800">
                                        <h4 className="text-[10px] font-black text-gray-500 uppercase tracking-[0.3em]">Mapping Configuration</h4>
                                        <div className="flex items-center gap-6">
                                            <button
                                                onClick={handleSaveToGradeMaster}
                                                title="Save current chemistry specs for this grade"
                                                className="flex items-center gap-2 px-4 py-2 bg-green-600/10 hover:bg-green-600/20 text-green-500 rounded-xl text-[10px] font-black uppercase tracking-widest border border-green-600/20 transition-all active:scale-95 group"
                                            >
                                                <Database className="w-3 h-3 group-hover:animate-pulse" />
                                                Save to Grade Master
                                            </button>
                                            <div className="flex items-center gap-4">
                                                <span className="text-[10px] font-black text-red-500 uppercase tracking-widest">Filter by Sample Id (Optional)</span>
                                                <select
                                                    className="bg-[#262730] border border-gray-700 px-4 py-2 rounded-xl text-xs font-bold outline-none focus:border-red-600 cursor-pointer shadow-lg"
                                                    value={sampleIdFilter}
                                                    onChange={e => handleSampleIdFilterChange(e.target.value)}
                                                >
                                                    <option value="">All Samples</option>
                                                    {availableSampleIds.map(s => <option key={s} value={s}>{s}</option>)}
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                    <div className="grid grid-cols-2 gap-10">
                                        <div className="space-y-3">
                                            <label className="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Master Heat (Column 1)</label>
                                            <select className="w-full bg-[#262730] border border-gray-700 p-4 rounded-2xl text-sm font-bold outline-none focus:border-red-600 hover:bg-[#31333f] transition-all cursor-pointer"
                                                value={mtcData.heat1} onChange={e => autoMapHeats(e.target.value, mtcData.heat2)}>
                                                <option value="">Select Heat No 1</option>
                                                {availableHeats
                                                    .filter(h => filteredData.some(d => (d['Heat No'] || d['HEAT NO'] || d['heat_no'] || '').toString() === h))
                                                    .map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                        <div className="space-y-3">
                                            <label className="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Secondary Heat (Column 2)</label>
                                            <select className="w-full bg-[#262730] border border-gray-700 p-4 rounded-2xl text-sm font-bold outline-none focus:border-red-600 hover:bg-[#31333f] transition-all cursor-pointer"
                                                value={mtcData.heat2} onChange={e => autoMapHeats(mtcData.heat1, e.target.value)}>
                                                <option value="">Select Heat No 2</option>
                                                {availableHeats
                                                    .filter(h => filteredData.some(d => (d['Heat No'] || d['HEAT NO'] || d['heat_no'] || '').toString() === h))
                                                    .map(h => <option key={h} value={h}>{h}</option>)}
                                            </select>
                                        </div>
                                    </div>
                                </div>
                                <div className="overflow-hidden rounded-2xl border border-gray-700 shadow-2xl">
                                    <table className="w-full text-sm text-left">
                                        <thead>
                                            <tr className="bg-[#31333f] text-gray-300 font-black">
                                                <th className="p-5 uppercase tracking-tighter text-sm">Spec Element</th>
                                                <th className="p-5 uppercase tracking-tighter text-sm text-gray-400">Requirement</th>
                                                <th className="p-5 uppercase tracking-tighter font-bold text-white text-sm">{mtcData.heat1 || 'H1'}</th>
                                                <th className="p-5 uppercase tracking-tighter font-bold text-white text-sm">{mtcData.heat2 || 'H2'}</th>
                                                <th className="p-5 text-center text-sm">Action</th>
                                            </tr>
                                        </thead>
                                        <tbody className="bg-[#0e1117]">
                                            {mtcData.chemistry.map((item, i) => (
                                                <tr key={i} className="border-b border-gray-800 hover:bg-white/[0.03] transition-colors group">
                                                    <td className="p-5 font-black text-white uppercase text-sm">
                                                        <input
                                                            className="bg-transparent border-none w-full p-0 outline-none focus:text-red-500 transition-colors font-black uppercase"
                                                            value={item.Element}
                                                            onChange={e => {
                                                                const next = [...mtcData.chemistry];
                                                                next[i].Element = e.target.value;
                                                                setMtcData({ ...mtcData, chemistry: next });
                                                            }}
                                                        />
                                                    </td>
                                                    <td className="p-5 bg-white/5">
                                                        <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 text-white font-bold text-sm transition-all"
                                                            value={item.Spec} onChange={e => {
                                                                const next = [...mtcData.chemistry];
                                                                next[i].Spec = e.target.value;
                                                                setMtcData({ ...mtcData, chemistry: next });
                                                            }} />
                                                    </td>
                                                    <td className="p-5 bg-red-500/[0.02]">
                                                        <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 transition-all font-mono text-gray-100 text-sm font-bold"
                                                            value={item.heat1_val} onChange={e => {
                                                                const next = [...mtcData.chemistry];
                                                                next[i].heat1_val = e.target.value;
                                                                setMtcData({ ...mtcData, chemistry: next });
                                                            }} />
                                                    </td>
                                                    <td className="p-5 bg-red-500/[0.02]">
                                                        <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 transition-all font-mono text-gray-100 text-sm font-bold"
                                                            value={item.heat2_val} onChange={e => {
                                                                const next = [...mtcData.chemistry];
                                                                next[i].heat2_val = e.target.value;
                                                                setMtcData({ ...mtcData, chemistry: next });
                                                            }} />
                                                    </td>
                                                    <td className="p-5 text-center flex items-center justify-center gap-4">
                                                        <input type="checkbox" title="Hide/Show" className="w-5 h-5 rounded-lg bg-[#262730] border-gray-600 text-red-600 focus:ring-red-600 transition-all cursor-pointer"
                                                            checked={item.Hide} onChange={() => {
                                                                const next = [...mtcData.chemistry];
                                                                next[i].Hide = !next[i].Hide;
                                                                setMtcData({ ...mtcData, chemistry: next });
                                                            }} />
                                                        <button onClick={() => deleteElement(i)} className="p-2 hover:bg-red-500/10 rounded-lg transition-colors group/del">
                                                            <Trash2 className="w-4 h-4 text-gray-600 group-hover/del:text-red-500" />
                                                        </button>
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                                <div className="flex justify-end mt-4 px-2">
                                    <button
                                        onClick={addChemistryRow}
                                        className="flex items-center gap-2 px-6 py-2 bg-red-600/10 hover:bg-red-600/20 text-red-500 rounded-xl text-xs font-black uppercase tracking-widest border border-red-600/20 transition-all active:scale-95"
                                    >
                                        + Add Spec Element
                                    </button>
                                </div>
                            </div>
                        </div>

                        {/* Sec 3: Mechanical */}
                        <div className="space-y-8">
                            <h3 className="text-xl font-bold text-gray-200 flex items-center gap-3">
                                <span className="bg-red-600 text-white w-10 h-10 rounded-2xl flex items-center justify-center text-lg font-black shadow-lg shadow-red-900/20 italic">3</span>
                                Mechanical Properties
                            </h3>
                            <div className="bg-[#262730] p-10 rounded-3xl border border-gray-700 shadow-2xl overflow-hidden">
                                <table className="w-full text-sm text-left">
                                    <thead>
                                        <tr className="bg-[#31333f] text-gray-300 font-black">
                                            <th className="p-5 uppercase tracking-tighter text-sm">Test Category</th>
                                            <th className="p-5 uppercase tracking-tighter text-sm text-gray-400">Standard Spec</th>
                                            <th className="p-5 uppercase tracking-tighter text-center italic text-red-500 text-sm">Observations</th>
                                        </tr>
                                    </thead>
                                    <tbody className="bg-[#0e1117]">
                                        {mtcData.mechanical.map((m, i) => (
                                            <tr key={i} className="border-b border-gray-800 hover:bg-white/[0.03] transition-colors">
                                                <td className="p-5">
                                                    <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 font-black text-white uppercase text-sm transition-all"
                                                        value={m.Parameter} onChange={e => {
                                                            const next = [...mtcData.mechanical];
                                                            next[i].Parameter = e.target.value;
                                                            setMtcData({ ...mtcData, mechanical: next });
                                                        }} />
                                                </td>
                                                <td className="p-5 bg-white/5">
                                                    <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 text-white font-bold text-sm transition-all"
                                                        value={m.Spec} onChange={e => {
                                                            const next = [...mtcData.mechanical];
                                                            next[i].Spec = e.target.value;
                                                            setMtcData({ ...mtcData, mechanical: next });
                                                        }} />
                                                </td>
                                                <td className="p-5" colSpan={2}>
                                                    <input className="bg-transparent border-b border-gray-700 w-full p-2 outline-none focus:border-red-600 font-mono text-gray-100 text-sm font-bold transition-all"
                                                        value={m.heat1_val} onChange={e => {
                                                            const next = [...mtcData.mechanical];
                                                            next[i].heat1_val = e.target.value;
                                                            setMtcData({ ...mtcData, mechanical: next });
                                                        }} />
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        {/* Footer Buttons */}
                        <div className="flex justify-between items-center py-12 border-t border-gray-800 mt-20 no-print">
                            <div className="flex gap-8">
                                <button
                                    onClick={() => window.print()}
                                    className="px-10 py-4 bg-[#262730] hover:bg-gray-700 rounded-2xl text-sm font-black uppercase tracking-widest border border-gray-600 transition-all flex items-center gap-3 shadow-xl active:scale-95"
                                >
                                    <Printer className="w-5 h-5 text-red-500" />
                                    Print Review
                                </button>
                                <button className="px-6 py-2 bg-transparent hover:bg-white/5 rounded-2xl text-xs transition-all flex items-center gap-2 underline text-red-400 font-bold tracking-tighter decoration-red-900 active:opacity-50">
                                    Direct Export to PDF
                                </button>
                            </div>
                            <button
                                onClick={handleDownload}
                                disabled={loading}
                                className="px-14 py-5 bg-red-600 hover:bg-red-700 text-white rounded-3xl transition-all shadow-2xl shadow-red-900/50 flex items-center gap-4 font-black uppercase tracking-[0.2em] italic active:scale-90 disabled:opacity-50 ring-4 ring-red-500/10"
                            >
                                {loading ? <Loader2 className="w-6 h-6 animate-spin text-white" /> : <Download className="w-6 h-6 text-white" />}
                                Export MTC (.xlsx)
                            </button>
                        </div>
                    </div>
                ) : activeTab === 'grade-master' ? (
                    <div className="space-y-12 animate-in fade-in duration-500">
                        <div className="flex justify-between items-center mb-12">
                            <h2 className="text-3xl font-black border-l-8 border-red-600 pl-6 uppercase tracking-tighter">Grade Master</h2>
                            <button
                                onClick={() => {
                                    const name = prompt("Enter new Grade name:");
                                    if (name && name.trim()) {
                                        const trimmedName = name.trim();
                                        setGradeMaster(prev => ({
                                            ...prev,
                                            [trimmedName]: {
                                                chemistry: STANDARD_ELEMENTS.map(el => ({ Element: el, Spec: "-" })),
                                                mechanical: STANDARD_MECHANICAL.map(prop => ({ Parameter: prop, Spec: "-" }))
                                            }
                                        }));
                                        setSelectedGradeMaster(trimmedName);
                                    }
                                }}
                                className="bg-red-600 text-white px-8 py-3 rounded-2xl text-xs font-black uppercase tracking-widest shadow-xl shadow-red-900/30 hover:bg-red-700 active:scale-95 transition-all flex items-center gap-2"
                            >
                                <Database className="w-4 h-4" />
                                Add New Grade
                            </button>
                        </div>

                        <div className="grid grid-cols-[300px_1fr] gap-12 items-start">
                            {/* Grade List */}
                            <div className="bg-[#262730] rounded-3xl border border-gray-700 overflow-hidden shadow-2xl sticky top-8">
                                <div className="p-6 bg-gray-800/30 border-b border-gray-700">
                                    <h3 className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Available Grades ({Object.keys(gradeMaster).length})</h3>
                                </div>
                                <div className="max-h-[600px] overflow-y-auto custom-scrollbar">
                                    {Object.keys(gradeMaster).length === 0 ? (
                                        <div className="p-12 text-center space-y-4">
                                            <Database className="w-12 h-12 text-gray-700 mx-auto" />
                                            <p className="text-xs text-gray-500 font-bold italic">No grades saved yet.</p>
                                        </div>
                                    ) : (
                                        Object.keys(gradeMaster).sort().map(gradeName => (
                                            <div key={gradeName} className="group flex items-center justify-between p-4 border-b border-gray-800/50 hover:bg-red-600/5 transition-all">
                                                <button
                                                    onClick={() => setSelectedGradeMaster(gradeName)}
                                                    className={cn("text-sm font-bold truncate flex-1 text-left py-1", selectedGradeMaster === gradeName ? "text-red-500" : "text-gray-300 group-hover:text-white")}
                                                >
                                                    {gradeName}
                                                </button>
                                                <button
                                                    onClick={async () => {
                                                        if (confirm(`Delete "${gradeName}" from Grade Master?`)) {
                                                            await axios.delete(`${API_BASE}/grade-master/${encodeURIComponent(gradeName)}`);
                                                            await fetchGradeMaster();
                                                            if (selectedGradeMaster === gradeName) setSelectedGradeMaster("");
                                                        }
                                                    }}
                                                    className="opacity-0 group-hover:opacity-100 p-2 text-gray-600 hover:text-red-500 transition-all"
                                                >
                                                    <Trash2 className="w-4 h-4" />
                                                </button>
                                            </div>
                                        ))
                                    )}
                                </div>

                                {/* Grades in Current Report Section */}
                                <div className="p-6 bg-gray-800/30 border-b border-t border-gray-700 mt-0">
                                    <h3 className="text-[10px] font-black text-gray-400 uppercase tracking-widest">In Current Report ({gradesInRange.length})</h3>
                                </div>
                                <div className="max-h-[300px] overflow-y-auto custom-scrollbar">
                                    {gradesInRange.length === 0 ? (
                                        <div className="p-8 text-center">
                                            <p className="text-[10px] text-gray-500 font-bold italic">No grades found in report.</p>
                                        </div>
                                    ) : (
                                        gradesInRange.map(gName => (
                                            <div key={gName} className="group flex items-center justify-between p-4 border-b border-gray-800/50 hover:bg-white/[0.02] transition-all">
                                                <span className={cn("text-xs font-bold truncate flex-1", gradeMaster[gName] ? "text-gray-500" : "text-gray-300")}>
                                                    {gName}
                                                </span>
                                                {!gradeMaster[gName] ? (
                                                    <button
                                                        onClick={() => {
                                                            setGradeMaster(prev => ({
                                                                ...prev,
                                                                [gName]: { chemistry: [], mechanical: [] }
                                                            }));
                                                            setSelectedGradeMaster(gName);
                                                        }}
                                                        className="px-3 py-1 bg-red-600/10 text-red-500 rounded-lg hover:bg-red-600 hover:text-white transition-all font-black text-[9px] uppercase tracking-tighter"
                                                    >
                                                        + Master
                                                    </button>
                                                ) : (
                                                    <button
                                                        onClick={() => setSelectedGradeMaster(gName)}
                                                        className="p-2 text-green-500/50 hover:text-green-500 transition-all"
                                                    >
                                                        <CheckCircle className="w-4 h-4" />
                                                    </button>
                                                )}
                                            </div>
                                        ))
                                    )}
                                </div>
                            </div>

                            {/* Spec Editor */}
                            <div className="space-y-12">
                                {selectedGradeMaster ? (
                                    <div className="animate-in fade-in slide-in-from-right-8">
                                        <div className="bg-[#262730] p-10 rounded-3xl border border-gray-700 shadow-2xl space-y-10">
                                            <div className="flex justify-between items-center border-b border-gray-800 pb-8">
                                                <div>
                                                    <h3 className="text-2xl font-black text-white italic tracking-tighter mb-1 uppercase leading-none">{selectedGradeMaster}</h3>
                                                    <p className="text-gray-500 text-[10px] font-black uppercase tracking-widest">Master Specification Editor</p>
                                                </div>
                                                <div className="flex flex-col gap-2">
                                                    <button
                                                        onClick={() => {
                                                            const currentChem = gradeMaster[selectedGradeMaster].chemistry || [];
                                                            const existingNames = currentChem.map(c => c.Element?.trim().toUpperCase());
                                                            const missing = STANDARD_ELEMENTS.filter(el => !existingNames.includes(el.toUpperCase()));

                                                            if (missing.length === 0) {
                                                                alert("All standard elements are already present.");
                                                                return;
                                                            }

                                                            const newChem = [...currentChem, ...missing.map(el => ({ Element: el, Spec: "-" }))];
                                                            setGradeMaster(prev => ({
                                                                ...prev,
                                                                [selectedGradeMaster]: { ...prev[selectedGradeMaster], chemistry: newChem }
                                                            }));
                                                        }}
                                                        className="bg-blue-600/10 border border-blue-600/20 text-blue-500 px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-600/20 active:scale-95 transition-all flex items-center gap-2"
                                                    >
                                                        <Database className="w-3.5 h-3.5" />
                                                        Fill Standard Elements
                                                    </button>
                                                    <button
                                                        onClick={() => {
                                                            const currentMech = gradeMaster[selectedGradeMaster].mechanical || [];
                                                            const existingParams = currentMech.map(m => m.Parameter?.trim().toUpperCase());
                                                            const missing = STANDARD_MECHANICAL.filter(prop => !existingParams.includes(prop.toUpperCase()));

                                                            if (missing.length === 0) {
                                                                alert("All standard mechanical properties are already present.");
                                                                return;
                                                            }

                                                            const newMech = [...currentMech, ...missing.map(prop => ({ Parameter: prop, Spec: "-" }))];
                                                            setGradeMaster(prev => ({
                                                                ...prev,
                                                                [selectedGradeMaster]: { ...prev[selectedGradeMaster], mechanical: newMech }
                                                            }));
                                                        }}
                                                        className="bg-blue-600/10 border border-blue-600/20 text-blue-500 px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-600/20 active:scale-95 transition-all flex items-center gap-2"
                                                    >
                                                        <Database className="w-3.5 h-3.5" />
                                                        Fill Standard Mechanical
                                                    </button>
                                                </div>
                                                <button
                                                    onClick={async () => {
                                                        await axios.post(`${API_BASE}/grade-master`, {
                                                            grade: selectedGradeMaster,
                                                            specs: gradeMaster[selectedGradeMaster]
                                                        });
                                                        await fetchGradeMaster();
                                                        alert("Grade specifications updated successfully!");
                                                    }}
                                                    className="bg-green-600/10 border border-green-600/20 text-green-500 px-8 py-3 rounded-2xl text-xs font-black uppercase tracking-widest hover:bg-green-600/20 active:scale-95 transition-all flex items-center gap-2"
                                                >
                                                    <Save className="w-4 h-4" />
                                                    Save All Changes
                                                </button>
                                            </div>

                                            {/* Chemistry Master */}
                                            <div className="space-y-6">
                                                <div className="flex items-center gap-3">
                                                    <div className="w-8 h-8 rounded-xl bg-red-600/20 text-red-500 flex items-center justify-center font-black italic text-sm">C</div>
                                                    <h4 className="text-sm font-black text-gray-200 uppercase tracking-widest">Chemical Composition</h4>
                                                </div>
                                                <div className="bg-[#0e1117] rounded-2xl border border-gray-800 overflow-hidden shadow-inner font-mono">
                                                    <table className="w-full text-xs">
                                                        <thead className="bg-gray-800/30 text-[10px] text-gray-500 font-black uppercase tracking-widest">
                                                            <tr>
                                                                <th className="p-4 text-left border-r border-gray-800 w-1/3">Spec Element</th>
                                                                <th className="p-4 text-left">Default Requirement (%)</th>
                                                                <th className="p-4 w-12 text-center"></th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {(gradeMaster[selectedGradeMaster].chemistry || []).map((c, idx) => (
                                                                <tr key={idx} className="border-t border-gray-800/50 group hover:bg-white/[0.02]">
                                                                    <td className="p-2 border-r border-gray-800">
                                                                        <input
                                                                            className="w-full bg-transparent p-2 outline-none text-white font-bold"
                                                                            value={c.Element}
                                                                            onChange={e => {
                                                                                const newChem = [...gradeMaster[selectedGradeMaster].chemistry];
                                                                                newChem[idx].Element = e.target.value;
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], chemistry: newChem } }));
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td className="p-2">
                                                                        <input
                                                                            className="w-full bg-transparent p-2 outline-none text-red-500 font-black"
                                                                            value={c.Spec}
                                                                            onChange={e => {
                                                                                const newChem = [...gradeMaster[selectedGradeMaster].chemistry];
                                                                                newChem[idx].Spec = e.target.value;
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], chemistry: newChem } }));
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td className="p-2 text-center">
                                                                        <button
                                                                            onClick={() => {
                                                                                const newChem = gradeMaster[selectedGradeMaster].chemistry.filter((_, i) => i !== idx);
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], chemistry: newChem } }));
                                                                            }}
                                                                            className="p-2 text-gray-500 hover:text-red-500 hover:bg-red-500/10 rounded-xl transition-all group/gm-del"
                                                                            title="Delete Spec Element"
                                                                        >
                                                                            <Trash2 className="w-3.5 h-3.5 transition-colors" />
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            ))}
                                                            <tr>
                                                                <td colSpan={3} className="p-2 border-t border-gray-800 text-center">
                                                                    <button
                                                                        onClick={() => {
                                                                            const newChem = [...(gradeMaster[selectedGradeMaster].chemistry || []), { Element: "New Spec Element", Spec: "0.00%" }];
                                                                            setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], chemistry: newChem } }));
                                                                        }}
                                                                        className="text-[10px] font-black text-gray-600 hover:text-red-500 uppercase tracking-widest py-2 flex items-center gap-2 mx-auto"
                                                                    >
                                                                        + Add Spec Element
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>

                                            {/* Mechanical Master */}
                                            <div className="space-y-6">
                                                <div className="flex items-center gap-3">
                                                    <div className="w-8 h-8 rounded-xl bg-blue-600/20 text-blue-500 flex items-center justify-center font-black italic text-sm">M</div>
                                                    <h4 className="text-sm font-black text-gray-200 uppercase tracking-widest">Mechanical Properties</h4>
                                                </div>
                                                <div className="bg-[#0e1117] rounded-2xl border border-gray-800 overflow-hidden shadow-inner font-mono">
                                                    <table className="w-full text-xs">
                                                        <thead className="bg-gray-800/30 text-[10px] text-gray-500 font-black uppercase tracking-widest">
                                                            <tr>
                                                                <th className="p-4 text-left border-r border-gray-800 w-1/3">Parameter</th>
                                                                <th className="p-4 text-left">Standard Spec</th>
                                                                <th className="p-4 w-12 text-center"></th>
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {(gradeMaster[selectedGradeMaster].mechanical || []).map((m, idx) => (
                                                                <tr key={idx} className="border-t border-gray-800/50 group hover:bg-white/[0.02]">
                                                                    <td className="p-2 border-r border-gray-800">
                                                                        <input
                                                                            className="w-full bg-transparent p-2 outline-none text-white font-bold"
                                                                            value={m.Parameter}
                                                                            onChange={e => {
                                                                                const newMech = [...gradeMaster[selectedGradeMaster].mechanical];
                                                                                newMech[idx].Parameter = e.target.value;
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], mechanical: newMech } }));
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td className="p-2">
                                                                        <input
                                                                            className="w-full bg-transparent p-2 outline-none text-blue-400 font-black"
                                                                            value={m.Spec}
                                                                            onChange={e => {
                                                                                const newMech = [...gradeMaster[selectedGradeMaster].mechanical];
                                                                                newMech[idx].Spec = e.target.value;
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], mechanical: newMech } }));
                                                                            }}
                                                                        />
                                                                    </td>
                                                                    <td className="p-2 text-center">
                                                                        <button
                                                                            onClick={() => {
                                                                                const newMech = gradeMaster[selectedGradeMaster].mechanical.filter((_, i) => i !== idx);
                                                                                setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], mechanical: newMech } }));
                                                                            }}
                                                                            className="p-2 text-gray-500 hover:text-red-500 hover:bg-red-500/10 rounded-xl transition-all group/gm-mech-del"
                                                                            title="Delete Parameter"
                                                                        >
                                                                            <Trash2 className="w-3.5 h-3.5 transition-colors" />
                                                                        </button>
                                                                    </td>
                                                                </tr>
                                                            ))}
                                                            <tr>
                                                                <td colSpan={3} className="p-2 border-t border-gray-800 text-center">
                                                                    <button
                                                                        onClick={() => {
                                                                            const newMech = [...(gradeMaster[selectedGradeMaster].mechanical || []), { Parameter: "New Parameter", Spec: "0-0" }];
                                                                            setGradeMaster(prev => ({ ...prev, [selectedGradeMaster]: { ...prev[selectedGradeMaster], mechanical: newMech } }));
                                                                        }}
                                                                        className="text-[10px] font-black text-gray-600 hover:text-blue-500 uppercase tracking-widest py-2 flex items-center gap-2 mx-auto"
                                                                    >
                                                                        + Add Property
                                                                    </button>
                                                                </td>
                                                            </tr>
                                                        </tbody>
                                                    </table>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                ) : (
                                    <div className="bg-[#262730] p-32 rounded-3xl border border-gray-700 shadow-2xl text-center space-y-8 animate-in zoom-in-95">
                                        <Database className="w-24 h-24 text-gray-800 mx-auto animate-pulse" />
                                        <div>
                                            <h3 className="text-2xl font-black text-gray-500 italic uppercase">Select a Grade to Edit</h3>
                                            <p className="text-gray-600 text-xs font-bold uppercase tracking-widest mt-2">Manage all default specifications in one place</p>
                                        </div>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>
                ) : (
                    null
                )}

                {
                    error && (
                        <div className="fixed bottom-12 right-12 p-6 bg-red-950/90 text-white rounded-3xl shadow-[0_0_50px_rgba(239,68,68,0.3)] border border-red-500/50 flex items-center gap-4 backdrop-blur-xl animate-in slide-in-from-right-12 z-50">
                            <div className="bg-red-600 p-2 rounded-xl">
                                <AlertCircle className="w-6 h-6 text-white" />
                            </div>
                            <div>
                                <p className="text-red-400 font-black text-[10px] uppercase tracking-widest mb-1">System Alert</p>
                                <span className="text-sm font-bold text-gray-200">{error}</span>
                            </div>
                            <button onClick={() => setError(null)} className="ml-8 hover:bg-white/10 p-2 rounded-xl transition-all">✕</button>
                        </div>
                    )
                }
            </main >
        </div >
    );
}

