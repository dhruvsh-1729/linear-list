import React, { useState } from "react";
import * as XLSX from "xlsx";

const SubsheetProcessor: React.FC = () => {
  const [loading, setLoading] = useState(false);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      const output: { topic: string; subInfo: string }[] = [];

      workbook.SheetNames.forEach((sheetName) => {
        if (!sheetName.toLowerCase().includes("sub")) return;

        const sheet = workbook.Sheets[sheetName];
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (rows.length < 3) return; // skip if not enough rows

        const subNumberMatch = sheetName.match(/\d+/);
        const subNumber = subNumberMatch ? subNumberMatch[0] : "?";
        const subTitle = rows[0]?.[1] || ""; // first row, second column

        for (let i = 2; i < rows.length; i++) {
          const topic = rows[i]?.[1]; // 2nd column (index 1)
          if (topic && typeof topic === "string" && topic.trim() !== "") {
            output.push({
              topic: topic.trim(),
              subInfo: `${subNumber}. ${subTitle}`,
            });
          }
        }
      });

      if (output.length === 0) {
        alert("No valid subsheet topics found.");
        setLoading(false);
        return;
      }

      // Convert to worksheet and export
      const newSheet = XLSX.utils.json_to_sheet(output, {
        header: ["topic", "subInfo"],
      });
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Processed");

      XLSX.writeFile(newWorkbook, "processed_subsheets.xlsx");
    } catch (err) {
      console.error(err);
      alert("Error processing file");
    }

    setLoading(false);
  };

  return (
    <div className="p-4 border rounded-lg">
      <h2 className="text-lg font-bold mb-2">Subsheet Processor</h2>
      <input
        type="file"
        accept=".xlsx"
        onChange={handleFileUpload}
        disabled={loading}
      />
      {loading && <p className="mt-2 text-sm">Processing...</p>}
    </div>
  );
};

export default SubsheetProcessor;
