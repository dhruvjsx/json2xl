import React from "react";
import * as XLSX from "xlsx";
import * as XlsxPopulate from "xlsx-populate/browser/xlsx-populate";

const ExcelExportHelper = ({ data }) => {
  const createDownLoadData = () => {
    handleExport().then((url) => {
      // console.log(url);
      const downloadAnchorNode = document.createElement("a");
      downloadAnchorNode.setAttribute("href", url);
      downloadAnchorNode.setAttribute("download", "student_report.xlsx");
      downloadAnchorNode.click();
      downloadAnchorNode.remove();
    });
  };

  const workbook2blob = (workbook) => {
    const wopts = {
      bookType: "xlsx",
      bookSST: false,
      type: "binary",
    };

    const wbout = XLSX.write(workbook, wopts);

    // The application/octet-stream MIME type is used for unknown binary files.
    // It preserves the file contents, but requires the receiver to determine file type,
    // for example, from the filename extension.
    const blob = new Blob([s2ab(wbout)], {
      type: "application/octet-stream",
    });

    return blob;
  };

  const s2ab = (s) => {
    // The ArrayBuffer() constructor is used to create ArrayBuffer objects.
    // create an ArrayBuffer with a size in bytes
    const buf = new ArrayBuffer(s.length);

    // console.log(buf);

    //create a 8 bit integer array
    const view = new Uint8Array(buf);

    // console.log(view);
    //charCodeAt The charCodeAt() method returns an integer between 0 and 65535 representing the UTF-16 code
    for (let i = 0; i !== s.length; ++i) {
      // console.log(s.charCodeAt(i));
      view[i] = s.charCodeAt(i);
    }

    return buf;
  };

  const handleExport = () => {
    const title = [{ A: "Audit Report" }, {}];

    let table1 = [
      {
        A: "leadName",
        B: "Lead Email",
        C: "Phone Numbers",
        D: "lead Source",
        E: "Initial Status",
        F: "Final Status",
        G: "Initial Task",
        H: "final Task",
        I: "lead Source",
        // J: "leadId"
      },
    ];

    data.forEach((row) => {
      const leadStatuses = row.leadStatuses.map((item) => item);

      leadStatuses.sort(
        (a, b) => new Date(a.modifiedOn) - new Date(b.modifiedOn)
      );
      const task1 = row?.tasks?.map((item) => item)||"";
      console.log(task1?task1:"dhruv", "task1");
      if(task1)
      task1?.sort((a, b) => new Date(a.modifiedOn) - new Date(b.modifiedOn));
      // Get the status with the smallest modifiedOn value
      const smallestModifiedOnStatus = leadStatuses[0];
      const smallestTask = task1[0];
      console.log(smallestTask?.taskTitle,"smallestTask")
      const largestTask = task1[task1?.length - 1]||" ";
      // Get the status with the largest modifiedOn value
      const largestModifiedOnStatus = leadStatuses[leadStatuses.length - 1];

      // console.log(smallestModifiedOnStatus, "smallestModifiedOnStatus")
      // console.log(leadStatuses,"status")
      table1.push({
        A: row.leadName,
        B: row.email,
        C: row.phoneNumber,
        D: row.leadSource,
        E: smallestModifiedOnStatus.status,
        F: largestModifiedOnStatus.status,
        G: smallestTask?.taskTitle,
        H: largestTask?.taskTitle,
        I: row.leadSource,
        // J: row.leadId
      });
    });
    // studentDetails.forEach( (studentIndex)=>{

    // })

    const finalData = [...title, ...table1];

    // console.log(finalData);

    //create a new workbook
    const wb = XLSX.utils.book_new();

    const sheet = XLSX.utils.json_to_sheet(finalData, {
      skipHeader: true,
    });

    XLSX.utils.book_append_sheet(wb, sheet, "student_report");

    // binary large object
    // Since blobs can store binary data, they can be used to store images or other multimedia files.

    const workbookBlob = workbook2blob(wb);

    var headerIndexes = [];
    finalData.forEach((data, index) =>
      data["A"] === "leadName" ? headerIndexes.push(index) : null
    );
console.log(headerIndexes,"headerIndexes")
    const dataInfo = {
      titleRange: "A1:I2",
      // theadRange:
      //   headerIndexes?.length >= 1
      //     ? `A${headerIndexes[0] + 1}:J${headerIndexes[0] + 1}`
      //     : null,
      theadRange:
      headerIndexes?.length >= 1
        ? `A${headerIndexes[0] + 1}:I${headerIndexes[0] + 1}`
        : null,

    };

    return addStyle(workbookBlob, dataInfo);
  };

  const addStyle = (workbookBlob, dataInfo) => {
    return XlsxPopulate.fromDataAsync(workbookBlob).then((workbook) => {
      workbook.sheets().forEach((sheet) => {
        sheet.usedRange().style({
          fontFamily: "Arial",
          verticalAlignment: "center",
        });

        sheet.column("A").width(25);
        sheet.column("B").width(35);
        sheet.column("C").width(25);
        sheet.column("D").width(25);
        sheet.column("E").width(25);
        sheet.column("F").width(25);
        sheet.column("G").width(25);
        sheet.column("H").width(25);
        sheet.column("I").width(35);

        sheet.range(dataInfo.titleRange).merged(true).style({
          bold: true,
          horizontalAlignment: "center",
          verticalAlignment: "center",
        });
        if (dataInfo.theadRange) {
          sheet.range(dataInfo.theadRange).style({
            fill: "FFFD04",
            bold: true,
            horizontalAlignment: "center",
          });
  
        }

      });

      return workbook
        .outputAsync()
        .then((workbookBlob) => URL.createObjectURL(workbookBlob));
    });
  };

  return (
    <button
      onClick={() => {
        createDownLoadData();
      }}
      className="btn btn-primary float-end"
    >
      Export
    </button>
  );
};

export default ExcelExportHelper;
