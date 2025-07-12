import React, { useState } from 'react';
import * as XLSX from 'xlsx';

type ProcessedRow = { 'Sub-item': string; Data: string; Sheet: string };

type FileSheetMap = {
    [fileName: string]: {
        file: File;
        sheets: string[];
    };
};

type CheckedSheets = {
    [fileName: string]: {
        [sheetName: string]: boolean;
    };
};

const ExcelViewer: React.FC = () => {
    const [fileSheetMap, setFileSheetMap] = useState<FileSheetMap>({});
    const [checkedSheets, setCheckedSheets] = useState<CheckedSheets>({});
    const [processedData, setProcessedData] = useState<ProcessedRow[]>([]);
    const [loading, setLoading] = useState(false);

    // Extract sheet names from files
    const handleFileSelection = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const files = event.target.files;
        if (!files) return;
        const fileArr = Array.from(files);
        const map: FileSheetMap = {};
        const checked: CheckedSheets = {};

        for (const file of fileArr) {
            const data = await file.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(data), { type: 'array' });
            map[file.name] = { file, sheets: wb.SheetNames };
            checked[file.name] = {};
            wb.SheetNames.forEach((sheet) => (checked[file.name][sheet] = false)); // set to false by default
        }
        setFileSheetMap(map);
        setCheckedSheets(checked);
        setProcessedData([]);
    };

    // Handle checkbox change for sheets
    const handleSheetCheckboxChange = (fileName: string, sheetName: string) => {
        setCheckedSheets((prev) => ({
            ...prev,
            [fileName]: {
                ...prev[fileName],
                [sheetName]: !prev[fileName][sheetName],
            },
        }));
    };

    // Process only checked sheets in checked files
    const handleProcessFiles = async () => {
        setLoading(true);
        const allData: ProcessedRow[] = [];
        for (const fileName in fileSheetMap) {
            const { file, sheets } = fileSheetMap[fileName];
            const data = await file.arrayBuffer();
            const wb = XLSX.read(new Uint8Array(data), { type: 'array' });
            let fileHasData = false;
            for (const sheetName of sheets) {
                if (checkedSheets[fileName]?.[sheetName]) {
                    const sheet = wb.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
                    const maxCol = Math.max(...rows.map((row: any[]) => row.length));
                    let sheetHasData = false;
                    for (let groupIndex = 0; groupIndex * 2 < maxCol; groupIndex++) {
                        const colIndex = groupIndex * 2;
                        for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
                            const subItem = rows[rowIndex][colIndex];
                            const dataValue = rows[rowIndex][colIndex + 1] || '';
                            if (!subItem && !dataValue) continue;
                            if (subItem !== undefined && subItem !== null) {
                                allData.push({
                                    'Sub-item': subItem,
                                    Data: dataValue,
                                    Sheet: sheetName,
                                });
                                fileHasData = true;
                                sheetHasData = true;
                            }
                        }
                    }
                    // Add a blank row after each sheet if any data was added for this sheet
                    if (sheetHasData) {
                        allData.push({
                            'Sub-item': '',
                            Data: '',
                            Sheet: '',
                        });
                    }
                }
            }
            // Optionally, you can remove the file-level blank row if not needed anymore
            // if (fileHasData) {
            //     allData.push({
            //         'Sub-item': '',
            //         Data: '',
            //         Sheet: '',
            //     });
            // }
        }
        setProcessedData(allData);
        setLoading(false);
    };

    // Download as XLSX
    const handleDownloadXLSX = () => {
        if (processedData.length === 0) return;
        const ws = XLSX.utils.json_to_sheet(processedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'ProcessedData');
        XLSX.writeFile(wb, 'processed_data.xlsx');
    };

    // Download as CSV
    const handleDownloadCSV = () => {
        if (processedData.length === 0) return;
        const ws = XLSX.utils.json_to_sheet(processedData);
        const csv = XLSX.utils.sheet_to_csv(ws);
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'processed_data.csv';
        a.click();
        URL.revokeObjectURL(url);
    };

    return (
        <div className="min-h-screen bg-gray-50 py-8 font-sans">
            <div className="bg-white rounded-xl shadow-lg p-8 w-full">
            <input
                type="file"
                accept=".xlsx"
                multiple
                onChange={handleFileSelection}
                className="mb-5 p-2 rounded-md border border-gray-300 bg-gray-100 text-base"
            />
            {Object.keys(fileSheetMap).length > 0 && (
                <div className="mb-5">
                <p className="font-medium mb-3">Select sheets to include:</p>
                <div className="flex gap-5 overflow-x-auto pb-2">
                    {Object.entries(fileSheetMap).map(([fileName, { sheets }]) => (
                    <div
                        key={fileName}
                        className="w-full bg-gray-100 rounded-lg shadow-sm p-4 border border-gray-200 flex-shrink-0"
                    >
                        <div className="font-semibold text-sm mb-2 text-gray-800 break-all">
                        {fileName}
                        </div>
                        <div className='w-full grid grid-cols-8'>
                        {sheets.map((sheet) => (
                            <label
                            key={sheet}
                            className="flex items-center mb-2 text-sm cursor-pointer text-gray-800"
                            >
                            <input
                                type="checkbox"
                                checked={!!checkedSheets[fileName]?.[sheet]}
                                onChange={() => handleSheetCheckboxChange(fileName, sheet)}
                                className="accent-blue-600 mr-2 w-4 h-4"
                            />
                            {sheet}
                            </label>
                        ))}
                        </div>
                    </div>
                    ))}
                </div>
                <button
                    onClick={handleProcessFiles}
                    disabled={
                    loading ||
                    !Object.entries(checkedSheets).some(([file, sheets]) =>
                        Object.values(sheets).some(Boolean)
                    )
                    }
                    className={`mt-5 px-6 py-2 bg-blue-600 text-white rounded-lg font-semibold text-base shadow transition
                    ${loading ? "opacity-70 cursor-not-allowed" : "hover:bg-blue-700"}
                    ${
                        !Object.entries(checkedSheets).some(([file, sheets]) =>
                        Object.values(sheets).some(Boolean)
                        )
                        ? "opacity-60 cursor-not-allowed"
                        : ""
                    }
                    `}
                >
                    {loading ? 'Processing...' : 'Process Selected Sheets'}
                </button>
                </div>
            )}
            <div className="mb-5 mt-5">
                <button
                onClick={handleDownloadXLSX}
                disabled={processedData.length === 0}
                className={`mr-3 px-5 py-2 rounded-md font-medium text-base shadow
                    ${
                    processedData.length === 0
                        ? "bg-gray-200 text-gray-400 cursor-not-allowed"
                        : "bg-green-500 text-white hover:bg-green-600"
                    }
                `}
                >
                Download XLSX
                </button>
                <button
                onClick={handleDownloadCSV}
                disabled={processedData.length === 0}
                className={`px-5 py-2 rounded-md font-medium text-base shadow
                    ${
                    processedData.length === 0
                        ? "bg-gray-200 text-gray-400 cursor-not-allowed"
                        : "bg-orange-400 text-white hover:bg-orange-500"
                    }
                `}
                >
                Download CSV
                </button>
            </div>
            {processedData.length > 0 && (
                <div className="overflow-x-auto bg-white rounded-lg shadow mt-5">
                <table className="min-w-full border-collapse">
                    <thead>
                    <tr>
                        <th className="border border-gray-200 px-4 py-3 bg-gray-100 font-semibold">Sub-item</th>
                        <th className="border border-gray-200 px-4 py-3 bg-gray-100 font-semibold">Data</th>
                        <th className="border border-gray-200 px-4 py-3 bg-gray-100 font-semibold">Sheet</th>
                    </tr>
                    </thead>
                    <tbody>
                    {processedData.map((row, index) => (
                        <tr key={index}>
                        <td className="border border-gray-200 px-4 py-3">{row['Sub-item']}</td>
                        <td className="border border-gray-200 px-4 py-3">{row.Data}</td>
                        <td className="border border-gray-200 px-4 py-3">{row.Sheet}</td>
                        </tr>
                    ))}
                    </tbody>
                </table>
                </div>
            )}
            {processedData.length === 0 && !loading && (
                <p className="text-gray-400 mt-6 text-center">
                Please upload one or more XLSX files.
                </p>
            )}
            </div>
        </div>
    );
};

export default ExcelViewer;
