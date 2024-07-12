import { useState } from "react";
import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import FileUploader from "./fileUpload";
import ExcelJS from "exceljs";

function App() {
  const [worksheet, setWorksheet] = useState(null);
  const [hojasLista, setHojasLista] = useState([]);

  const [fileworkbook, setWorkbook] = useState(null);
  const [listRango, setlistRango] = useState("");
  const [headersLista, setHeadersLista] = useState([]);
  const [listaComparacion, setListaComparacion] = useState([]);
  const [listRangoObj, setlistRangoObj] = useState(null);
  const [tablaRango, setTablaRango] = useState("");
  const [tablaRangoObj, setTablaRangoObj] = useState(null);
  const [headerTabla, setHeadersTabla] = useState([]);
  const [IndexComparacion, setIndexComparacion] = useState(0);
  const [tabla, setTabla] = useState(null);

  const hangleListaRango = (e) => {
    console.log(e.target.value);
    setlistRango(e.target.value);
  };

  const hangleTablaRango = (e) => {
    console.log(e.target.value);
    setTablaRango(e.target.value);
  };

  const getSheets = (wb) => {
    const hojas = wb._worksheets.map((hoja) => {
      return hoja._name
    })
    console.log(hojas)
    setHojasLista(hojas)
  }

  const handleSelectSheet= (e) => {
    console.log("HOJA",e.target.value);
    const ws = fileworkbook.getWorksheet(e.target.value);
    // console.log("lista",lista)
    setWorksheet(ws);
  };

  const handleFileRead = (wb) => {
    setWorkbook(wb);
    console.log(wb._worksheets)
    getSheets(wb)
  };

  const handleSelectFilaOrden = (e) => {
    // console.log("valor",Number(e.target.value.split("mydot")[1]));
    // console.log(listRangoObj)
    const lista = obtenerListaComparativa(
      worksheet,
      listRangoObj,
      Number(e.target.value.split("mydot")[1])
    );
    // console.log("lista",lista)
    setListaComparacion(lista);
  };

  const cargarHeaders = () => {
    console.log("Procesando archivo...");
    console.log(worksheet);
    const rango = parseRange(listRango);
    setlistRangoObj(rango);
    const headers = obtenerHeaders(worksheet, rango);
    setHeadersLista(headers);
  };

  const cargarHeaderTabla = () => {
    // console.log("Procesando archivo...");
    // console.log(worksheet);
    const rangoTabla = parseRange(tablaRango);
    setTablaRangoObj(rangoTabla);
    const headersTabla = obtenerHeaders(worksheet, rangoTabla);
    setHeadersTabla(headersTabla);
    // const tabla = obtenerTabla(worksheet, rangoTabla, 1);

    // const headers = obtenerHeaders(worksheet, rango);
  };

  const handleSelectFilaTabla = (e) => {
    setIndexComparacion(Number(e.target.value.split("nodot")[1]))
    const tabla = obtenerTabla(worksheet, tablaRangoObj, Number(e.target.value.split("nodot")[1]));
    setTabla(tabla)
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
    console.log("rangeObj", rangeObj);
    console.log("worksheet", worksheet);
    console.log("index", index);

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
      // const rangoTabla = parseRange("F22:U179");

      console.log(listaComparacion)
      console.log(tabla)
      console.log(IndexComparacion)
      const tablaOrdenada = ordenarTabla(listaComparacion, tabla, IndexComparacion);
      console.log({ tablaOrdenada });
      await crearArchivo(headerTabla, tablaOrdenada);
      console.log("Archivo creado y descargado correctamente.");
    } catch (error) {
      console.error("Error al procesar y descargar el archivo:", error);
    }
  };

  return (
    <>
      <div>
        <img src={viteLogo} className="logo" alt="Vite logo" />
        <h1>Ordenador de excel</h1>
        <p><strong>Cargue su archivo</strong></p>
        <FileUploader onFileRead={handleFileRead} />
      </div>
      <div className="contenedor-form">
      {hojasLista.length > 0 ? (
              <>
                <br />
                <span>Seleccione la hoja del excel que tiene los datos</span>
                <select name="lista" id="lista" onChange={handleSelectSheet}>
                  {hojasLista.map((item, index) => {
                    return (
                      <option value={`${item}`}>{item}</option>
                    );
                  })}
                </select>
                <br />
              </>
            ) : (
              ""
            )}
        {worksheet ? (
          <>
            <span>
              Ingrese el rango de la tabla de comparacion (la que esta ordenada)
            </span>
            <input type="text" onChange={hangleListaRango} />
            <br />
            <button className="myboton" onClick={cargarHeaders}>Buscar tabla de comparación</button>
            {headersLista.length > 0 ? (
              <>
                <br />

                <span>Seleccione la fila ordenada</span>
                <select name="lista" id="lista" onChange={handleSelectFilaOrden}>
                  {headersLista.map((item, index) => {
                    return (
                      <option value={`${item}mydot${index}`}>{item}</option>
                    );
                  })}
                </select>
              </>
            ) : (
              ""
            )}
            <br />
            <span>Ingrese el rango de la tabla que quiere ordenar</span>
            <input type="text" onChange={hangleTablaRango} />
            <br />
            <button className="myboton" onClick={cargarHeaderTabla}>Buscar tabla a ordenar</button>
            {headerTabla.length > 0 ? (
              <>
                <br />

                <span>Seleccione la fila ordenada</span>
                <select name="tabla" id="tabla" onChange={handleSelectFilaTabla}>
                  {headerTabla.map((item, index) => {
                    return (
                      <option value={`${item}nodot${index}`}>{item}</option>
                    );
                  })}
                </select>
              </>
            ) : (
              ""
            )}
          </>
        ) : null}
      </div>
      {(tablaRangoObj &&listRangoObj && listaComparacion.length > 0 && tabla)?<button className="myboton" onClick={procesarArchivo}>Procesar</button>: ""}
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
