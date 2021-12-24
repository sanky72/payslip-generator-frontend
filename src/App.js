import React, { useEffect, useState } from "react";
import "./App.css";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import "jspdf-autotable";

function App() {
  const [items, setItems] = useState([]);
  const [sheetName, setSheetname] = useState("");
  const [error, setError] = useState(null);
  const [uploadedFile, setUploadedFile] = useState(null);
  const [resultGenerated, setResultGenerated] = useState(false);
  const [result, setResult] = useState(null);

  const setErrorNull = () => {
    setError(null);
  };

  const readExcel = () => {
    setResultGenerated(false);
    const promise = new Promise((resolve, reject) => {
      const fileReader = new FileReader();
      fileReader.readAsArrayBuffer(uploadedFile);

      fileReader.onload = (e) => {
        const bufferArray = e.target.result;

        const wb = XLSX.read(bufferArray, { type: "buffer" });

        const indexOfSheet = wb.SheetNames.indexOf(sheetName);

        if (indexOfSheet < 0) {
          return reject(
            new Error("The sheet is not present, please check your name")
          );
        }

        const wsname = wb.SheetNames[indexOfSheet];

        const ws = wb.Sheets[wsname];

        const data = XLSX.utils.sheet_to_json(ws);
        console.log("data -> ", data);

        resolve(data);
      };

      fileReader.onerror = (error) => {
        reject(error);
      };
    });

    promise.then(
      (d) => {
        setItems(d);
        setErrorNull();
      },
      (error) => {
        setError(error);
      }
    );
  };

  const findRowWhichContainsKey = (text, checkForWholeText = true) => {
    let indexOfRowOfWhichContainText;
    if (checkForWholeText) {
      items.forEach((row, index) => {
        if (Object.values(row).indexOf(text) >= 0) {
          indexOfRowOfWhichContainText = index;
        }
      });
    } else {
      items.forEach((row, index) => {
        Object.values(row).forEach((value) => {
          if (typeof value === "string" && value.indexOf(text) >= 0) {
            indexOfRowOfWhichContainText = index;
          }
        });
      });
    }

    return indexOfRowOfWhichContainText;
  };

  const createSalaryObj = (keys, empObj) => {
    const salaryObj = {};
    for (let it in empObj) {
      salaryObj[keys[it]] = empObj[it];
    }
    return salaryObj;
  };

  useEffect(() => {
    const indexOfRowWhichContainKeys = findRowWhichContainsKey("Income Tax");
    const indexOfLastRow = findRowWhichContainsKey("Grand Total", false);
    if (
      indexOfRowWhichContainKeys !== undefined &&
      indexOfRowWhichContainKeys !== null &&
      indexOfLastRow
    ) {
      const correctedSalaryJson = [];
      for (let i = indexOfRowWhichContainKeys + 1; i < indexOfLastRow; i++) {
        correctedSalaryJson.push(
          createSalaryObj(items[indexOfRowWhichContainKeys], items[i])
        );
      }
      console.log(correctedSalaryJson);
      setResult(correctedSalaryJson);
      setResultGenerated(true);
    }
  }, [items]);

  function generatePdfs() {
    result.forEach((emp) => {
      let doc = new jsPDF();
      const tableColumn = Object.keys(emp);
      const tableRows = [];
      tableColumn.forEach((key) => {
        tableRows.push(emp[key]);
      });
      console.log(tableRows);
      doc.text(`Payslip for the month ${sheetName}`, 14, 10);

      let start = 0;
      while (start < tableColumn.length) {
        let left = start;
        let right = start + 2;
        console.log(
          tableColumn.slice(left, right + 1),
          tableRows.slice(left, right + 1)
        );
        doc.autoTable({
          head: [tableColumn.slice(left, right + 1)],
          body: [tableRows.slice(left, right + 1)],
          theme: "grid",
        });
        start += 2;
      }
      doc.save(`${emp["Name Of Employee"]}-${sheetName}.pdf`);
    });
  }

  return (
    <div>
      <input
        type="file"
        onChange={(e) => {
          const file = e.target.files[0];
          setUploadedFile(file);
          setErrorNull();
        }}
      />
      <h3>
        Please give the sheet name for the month you want to generate payslips
      </h3>
      <input
        type="input"
        placeholder="eg. -> MAY-2021"
        onChange={(e) => {
          setSheetname(e.target.value);
          setErrorNull();
        }}
      ></input>
      <button
        onClick={() => {
          readExcel();
        }}
      >
        Submit{" "}
      </button>
      {resultGenerated ? (
        <button onClick={() => generatePdfs()}>Generate Pdf</button>
      ) : (
        ""
      )}

      {error ? <div>There is an error</div> : ""}
    </div>
  );
}

export default App;
