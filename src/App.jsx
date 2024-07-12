import { useState } from "react";
import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import FileUploader from "./fileUpload";
import ExcelJS from "exceljs";

function App() {
  const [worksheet, setWorksheet] = useState(null);
  const [listRango, setlistRango] = useState("");
  const [headersLista, setHeadersLista] = useState([]);
  const [listaComparacion, setListaComparacion] = useState([]);
  const [listRangoObj, setlistRangoObj] = useState(null);


  const hangleListaRango = (e) => {
    console.log(e.target.value);
    setlistRango(e.target.value);
  };

  const handleFileRead = (ws) => {
    setWorksheet(ws);
  };

  const handleSelectFilaOrden = (e) => {
    // console.log("valor",Number(e.target.value.split("mydot")[1]));
    // console.log(listRangoObj)
    const lista = obtenerListaComparativa(worksheet, listRangoObj, Number(e.target.value.split("mydot")[1]));
    // console.log("lista",lista)
    setListaComparacion(lista)
  }

  const cargarHeaders = () => {
    console.log("Procesando archivo...");
    console.log(worksheet);
    const rango = parseRange(listRango);
    setlistRangoObj(rango)
    const headers = obtenerHeaders(worksheet, rango);
    setHeadersLista(headers)
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
    console.log("rangeObj",rangeObj)
    console.log("worksheet",worksheet)
    console.log("index",index)

    for (let i = rangeObj.startRow + 1; i <= rangeObj.endRow; i++) {
      const row = worksheet.getRow(i);
      // console.log(row,  "algo", rangeObj.startCol)
      // console.log(row.getCell(rangeObj.startCol+1).value)
      lista.push(row.getCell(rangeObj.startCol + index).value);
    }
    console.log(lista);
    return lista;
  };

  const obtenerTabla = (worksheet, rangeObj, index) => {
    const tabla = [];
    const mapa = new Map();
    for (let i = rangeObj.startRow + 1; i <= rangeObj.endRow; i++) {
      const row = worksheet.getRow(i);
      const rowCells = [];
      let key = "";
      for (let j = rangeObj.startCol; j <= rangeObj.endCol; j++) {
        rowCells.push(row.getCell(j).value);
        if (j === rangeObj.startCol + index) {
          key = row.getCell(j).value;
        }
      }
      tabla.push({
        campos: rowCells,
        usado: 0,
        key,
      });
      mapa.set(key, tabla.length - 1);
    }
    return { tabla, mapa };
  };
  function createEmptyArray(cuantity, index, key) {
    const fila = new Array(cuantity).fill("");
    fila[index] = key;
    return fila;
  }

  const ordenarTabla = (lista, tablaMapa, colIndex) => {
    const tablaOrdenada = [];
    for (const item of lista) {
      const index = tablaMapa.mapa.get(item);
      // tablaMapa.tabla[index].usado = 1
      if (index) {
        // console.log(item, index, tablaMapa.tabla[index]);
        tablaMapa.tabla[index].usado = 1;
        tablaOrdenada.push(tablaMapa.tabla[index]);
      } else {
        tablaOrdenada.push({
          campos: createEmptyArray(
            tablaMapa.tabla[0].campos.length,
            colIndex,
            item
          ),
          usado: 1,
          key: item,
        });
      }
    }
    const nousado = tablaMapa.tabla.filter((fila) => fila.usado === 0);
    console.log("no usados", nousado);
    tablaOrdenada.push(...nousado);
    return tablaOrdenada;
  };

  function initialiceFileAndsheet() {
    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet("data", {
      properties: { defaultColWidth: 12 },
    });
    return { sheet, book };
  }

  const crearArchivo = async (headersTabla, tablaOrdenada) => {
    return new Promise(async (resolve, reject) => {
      try {
        const { sheet, book } = initialiceFileAndsheet();
        sheet.addRow(headersTabla);
        for (const item of tablaOrdenada) {
          sheet.addRow(item.campos);
        }
        const buffer = await book.xlsx.writeBuffer();
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "excel_file.xlsx";
        a.click();
        resolve();
      } catch (error) {
        reject(error);
      }
    });
  };

  const procesarArchivo = async () => {
    try {
      console.log("Procesando archivo...");
      console.log(worksheet);
      // const rango = parseRange("B22:D179");
      // const headers = obtenerHeaders(worksheet, rango);
      // const lista = obtenerListaComparativa(worksheet, rango, 1);
      const rangoTabla = parseRange("F22:U179");
      const headersTabla = obtenerHeaders(worksheet, rangoTabla);
      const tabla = obtenerTabla(worksheet, rangoTabla, 1);
      const tablaOrdenada = ordenarTabla(lista, tabla, 1);
      console.log({ tablaOrdenada });
      await crearArchivo(headersTabla, tablaOrdenada);
      console.log("Archivo creado y descargado correctamente.");
    } catch (error) {
      console.error("Error al procesar y descargar el archivo:", error);
    }
  };

  return (
    <>
      <div>
        <img src={viteLogo} className="logo" alt="Vite logo" />
        <h1>Pordenador de excel</h1>
        <p>Cargue su archivo</p>
        <FileUploader onFileRead={handleFileRead} />
      </div>
      <div>
        {worksheet ? (
          <>
            <span>
              Ingrese el rango de la tabla de comparacion (la que esta ordenada)
            </span>
            <input type="text" onChange={hangleListaRango} />
            <br />
            <button onClick={cargarHeaders}>Buscar tabla de comparación</button>
            <br />

            <span>Seleccione la fila ordenada</span>
            <select name="cars" id="cars" onChange={handleSelectFilaOrden}>
              {headersLista.map((item, index)=> {
                return (<option value={`${item}mydot${index}`} >{item}</option>)
                
              })}
            </select>
          </>
        ) : null}
      </div>
      <button onClick={procesarArchivo}>Procesar</button>
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
