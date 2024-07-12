import { useState } from "react";
import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import FileUploader from "./fileUpload";
import ExcelJS from "exceljs";

function App() {
  const [worksheet, setWorksheet] = useState(null);

  const handleFileRead = (worksheet) => {
    setWorksheet(worksheet);
  };

  const parseRange = (range) => {
    const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) return null;
    console.log(match);
    const [, startCol, startRow, endCol, endRow] = match;
    return {
      startCol: ExcelJS.utils.columnToNumber(startCol),
      startRow: parseInt(startRow, 10),
      endCol: ExcelJS.utils.columnToNumber(endCol),
      endRow: parseInt(endRow, 10),
    };
  };

  const obtenerHeaders = (worksheet, rangeObj) => {
    const row = worksheet.getRow(rangeObj.startRow);
    const headers = [];
    for (let j = rangeObj.startCol; j <= rangeObj.endCol; j++) {
      headers.push(row.getCell(j).value);
    }
    console.log("headers", headers);
    return headers;
  };

  const obtenerListaComparativa = (worksheet, rangeObj, index) => {
    const lista = [];
    for (let i = rangeObj.startRow + 1; i <= rangeObj.endRow; i++) {
      const row = worksheet.getRow(i);
      lista.push(row.getCell(rangeObj.startCol + index).value);
    }
    console.log(lista);
  };

  const obtenerTabla = (worksheet, rangeObj) => {
    const tabla = [];
    for (let i = rangeObj.startRow; i <= rangeObj.endRow; i++) {
      const row = worksheet.getRow(i);
      const rowCells = [];
      for (let j = rangeObj.startCol; j <= rangeObj.endCol; j++) {
        rowCells.push(row.getCell(j).value);
      }
      rows.push(rowCells);
    }
  };

  const procesarArchivo = () => {
    console.log("holi");
    console.log(worksheet);
    const rango = parseRange("B22:D179");
    console.log(rango);
    const headers = obtenerHeaders(worksheet, rango);
    obtenerListaComparativa(worksheet, rango, 1);
    const rangoTabla = parseRange("F22:U179");
    console.log({ rangoTabla });

    // const range = worksheet.getRange('B22:D178');
    // console.log(range)
  };

  return (
    <>
      <div>
        <a href="https://vitejs.dev" target="_blank">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <button onClick={procesarArchivo}>Procesar</button>
      <FileUploader onFileRead={handleFileRead} />
    </>
  );
}

ExcelJS.utils = {
  columnToNumber: (col) => {
    let number = 0;
    let pow = 1;
    for (let i = col.length - 1; i >= 0; i--) {
      number += (col.charCodeAt(i) - 64) * pow;
      pow *= 26;
    }
    return number;
  },
};

export default App;
