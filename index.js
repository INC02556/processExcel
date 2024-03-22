import * as XLSX from "xlsx";

function processExcel(file) {
  return new Promise((resolve, reject) => {
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX?.read(data, { type: "array" });
        const sheet = workbook?.Sheets[workbook.SheetNames[0]];
        const headerData = {};
        let header = XLSX.utils?.sheet_to_json(sheet, { header: 1 })?.[0];
        let valueHeader = XLSX?.utils.sheet_to_json(sheet, { header: 1 })?.[1];
        header?.forEach((head, index) => {
          if (index !== 0) {
            headerData[`${head}`] = String(valueHeader[index] || "") ;
          }
        });
        let tableHeader = XLSX.utils.sheet_to_json(sheet, { header: 1 })?.[2];
        let tableRow = XLSX.utils.sheet_to_json(sheet, { header: 1 }).slice(3);
        let tableData = tableRow?.map((item) => {
            let eachObj = {}
             tableHeader?.forEach((head, headindex) => {
              if (headindex !== 0) {
                eachObj[`${head}`] = String(item[headindex] || "");
              }
            });
            return eachObj
        });
        const jsonResult = {
          header: [headerData],
          table: tableData,
        };
        resolve(jsonResult);
      };

      reader.readAsArrayBuffer(file);
    } catch (error) {
      reject("Error processing Excel file: " + error.message);
    }
  });
}

module.exports = processExcel