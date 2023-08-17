import React, { useState } from 'react';
import * as Excel from 'exceljs';
import { saveAs } from "file-saver";

export default function App() {
  const handleChange = async (e) => {
    const file = e.target.files[0];
    const workbook = new Excel.Workbook();
    const reader = new FileReader();
    const data = {
    
      B7: 61197,
      B9: -2672,
      B12: 603,
      B17: 123606

    };
    
    reader.readAsArrayBuffer(file);
    reader.onload = async () => {
      let buffer = reader.result;
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.getWorksheet('hoja1');
      Object.entries(data).forEach(([cell, value]) => {
        worksheet.getCell(cell).value = value;
      });
 
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (cell.type === Excel.ValueType.Formula) {
            cell.value = cell.value;
          }
        });
      });
      const newBuffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([newBuffer], { type: "application/octet-stream" }), "nuevo-archivo.xlsx");
    };
  };

  return (
    <div>
      <input type="file" onChange={(e) => handleChange(e)} />
    </div>
  );
}
