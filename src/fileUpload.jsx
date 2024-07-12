import React, { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import ExcelJS from 'exceljs';

const FileUploader = ({ onFileRead }) => {

  const onDrop = useCallback(acceptedFiles => {
    const file = acceptedFiles[0];
    const reader = new FileReader();

    reader.onload = async (e) => {
      const arrayBuffer = e.target.result;
      if (arrayBuffer) {
        const uint8Array = new Uint8Array(arrayBuffer);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(uint8Array);

        const worksheet = workbook.getWorksheet("6. Verificación partidas"); // Assuming you want to read the first sheet
        let content = '';
        worksheet.eachRow((row, rowNumber) => {
          row.eachCell((cell, colNumber) => {
            content += `${cell.value}\t`; // Add tab between cell values
          });
          content += '\n'; // Add newline at the end of each row
        });
        onFileRead(worksheet);
      } else {
        console.error('Error reading file');
      }
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const { getRootProps, getInputProps } = useDropzone({ onDrop });

  return (
    <div>
      <div {...getRootProps({ style: dropzoneStyle })}>
        <input {...getInputProps()} />
        <p>Arrastra un archivo aquí, o haz clic para seleccionar uno</p>
      </div>
    </div>
  );
};

const dropzoneStyle = {
  border: '2px dashed #cccccc',
  borderRadius: '4px',
  padding: '20px',
  textAlign: 'center',
  cursor: 'pointer'
};

export default FileUploader;
