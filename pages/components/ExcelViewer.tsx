import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

const ExcelViewer: React.FC = () => {
    // State variables
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState<string | null>(null);
    const [sheetData, setSheetData] = useState<
        { 'Sub-item': string; Data: string; Filename: string }[] | null
    >(null);
    const [fileName, setFileName] = useState<string>('');

    // Handle file upload
    const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file) {
            setFileName(file.name);
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target?.result as ArrayBuffer);
                const wb = XLSX.read(data, { type: 'array' });
                setWorkbook(wb);
                setSelectedSheet(null); // Reset sheet selection
                setSheetData(null); // Reset processed data
            };
            reader.readAsArrayBuffer(file);
        }
    };

    // Process sheet data when workbook and sheet are selected
    useEffect(() => {
        if (workbook && selectedSheet) {
            const sheet = workbook.Sheets[selectedSheet];
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

            const processedData: { 'Sub-item': string; Data: string; Filename: string }[] = [];
            const maxCol = Math.max(...data.map((row: any[]) => row.length));

            for (let groupIndex = 0; groupIndex * 2 < maxCol; groupIndex++) {
                const colIndex = groupIndex * 2; // Odd columns: 0 (A), 2 (C), 4 (E), etc.

                // Process rows starting from index 1
                for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
                    const subItem = data[rowIndex][colIndex]; // Sub-item number from odd column
                    const dataValue = data[rowIndex][colIndex + 1] || ''; // Data from even column

                    // Skip the row if both subItem and dataValue are blank
                    if (!subItem && !dataValue) {
                        continue;
                    }

                    // Add a row for the group when the group changes
                    // const currentGroup = subItem ? subItem.split('.')[0] + '.' + subItem?.split('.')[1] : null;
                    // if (currentGroup !== previousGroup) {
                    //     processedData.push({
                    //         'Sub-item': '', // Extracted 2.1 from 2.1.1
                    //         Data: currentGroup || '',
                    //         Filename: '',
                    //     });
                    //     previousGroup = currentGroup;
                    // }

                    // Include the actual data row
                    if (subItem !== undefined && subItem !== null) {
                        processedData.push({
                            'Sub-item': subItem,
                            Data: dataValue,
                            Filename: selectedSheet,
                        });
                    }
                }
            }
            setSheetData(processedData);
        } else {
            setSheetData(null);
        }
    }, [workbook, selectedSheet, fileName]);

    // Render the component
    return (
        <div style={{ padding: '20px' }}>
            <input
                type="file"
                accept=".xlsx"
                onChange={handleFileUpload}
                style={{ marginBottom: '10px' }}
            />

            {workbook ? (
                <div>
                    <select
                        value={selectedSheet || ''}
                        onChange={(e) => {
                            const value = e.target.value;
                            setSelectedSheet(value === '' ? null : value);
                        }}
                        style={{ marginBottom: '10px', padding: '5px' }}
                    >
                        <option value="">Select a sheet</option>
                        {workbook.SheetNames.map((name) => (
                            <option key={name} value={name}>
                                {name}
                            </option>
                        ))}
                    </select>

                    {sheetData && sheetData.length > 0 ? (
                        <table style={{ borderCollapse: 'collapse', width: '100%' }}>
                            <thead>
                                <tr>
                                    <th style={{ border: '1px solid #ccc', padding: '8px' }}>Sub-item</th>
                                    <th style={{ border: '1px solid #ccc', padding: '8px' }}>Data</th>
                                    <th style={{ border: '1px solid #ccc', padding: '8px' }}>Filename</th>
                                </tr>
                            </thead>
                            <tbody>
                                {sheetData.map((row, index) => (
                                    <tr key={index}>
                                        <td style={{ border: '1px solid #ccc', padding: '8px' }}>{row['Sub-item']}</td>
                                        <td style={{ border: '1px solid #ccc', padding: '8px' }}>{row.Data}</td>
                                        <td style={{ border: '1px solid #ccc', padding: '8px' }}>{row.Filename}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    ) : (
                        <p>Please select a sheet and ensure it has data.</p>
                    )}
                </div>
            ) : (
                <p>Please upload an XLSX file.</p>
            )}
        </div>
    );
};

export default ExcelViewer;