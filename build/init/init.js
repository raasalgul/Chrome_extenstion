// const xl = require('excel4node');
function init() {
  console.log("Initializing....");
  var link = document.createElement('link');
  link.setAttribute('rel', 'stylesheet');
  link.setAttribute('type', 'text/css');
  link.setAttribute('href', 'https://fonts.googleapis.com/css2?family=Special+Elite&display=swap');
  document.head.appendChild(link);
  var inputs = getInputsByValue();
  const div = document.createElement('div');
  var finalHTML = '<div id="cheta-flt-dv"><p class="cheta-flt-p">ChETA</p>';
  finalHTML += '<p class="cheta-pfnt">Words: <span id="cheta-data-wordcount">'+inputs+'</span></p>';
  finalHTML += '</div>';
  div.innerHTML = finalHTML;
  document.body.appendChild(div);
}

function getInputsByValue()
{
    let excelData=[];
    let issueList = document.getElementsByClassName("issue-list")[0].innerText;
    console.log(`issueList ${issueList}`);
    let totalIssues=document.getElementsByClassName("showing")[0]
    .getElementsByTagName("span")[0].innerText.split(" of ")[1];
    console.log(`issueList total count ${totalIssues}`);
    for(let i=0;i<totalIssues;i++)
    {
     task(i,totalIssues);
    }
    // const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
    // const fileExtension = '.xlsx';

    // const exportToCSV = (csvData, fileName) => {
    //     const ws = XLSX.utils.json_to_sheet(csvData);
    //     const wb = { Sheets: { 'data': ws }, SheetNames: ['data'] };
    //     const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    //     const data = new Blob([excelBuffer], {type: fileType});
    //     FileSaver.saveAs(data, fileName + fileExtension);
    // }
    // exportToCSV(excelData,"prod_data");

//     const wb = new xl.Workbook();
//     const ws = wb.addWorksheet('Worksheet Name');
//     const headingColumnNames = [
//       "email","status","totalSalesForces",
//       "description","comments"
//   ]
//   let headingColumnIndex = 1;
//   headingColumnNames.forEach(heading => {
//       ws.cell(1, headingColumnIndex++)
//           .string(heading)
//   });
//   let rowIndex = 2;
//   excelData.forEach( record => {
//     let columnIndex = 1;
//     Object.keys(record ).forEach(columnName =>{
//         ws.cell(rowIndex,columnIndex++)
//             .string(record [columnName])
//     });
//     rowIndex++;
// }); 
// wb.write('prod-data.xlsx');
// wb.SheetNames.push("Test Sheet");


var wb = XLSX.utils.book_new();
wb.Props = {
        Title: "SheetJS Tutorial",
        Subject: "Test",
        Author: "Red Stapler",
        CreatedDate: new Date(2017,12,19)
};
var ws = XLSX.utils.aoa_to_sheet(excelData);
wb.Sheets["Test Sheet"] = ws;
var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});
  saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'test.xlsx');
function s2ab(s) { 
  var buf = new ArrayBuffer(s.length); //convert s to arrayBuffer
  var view = new Uint8Array(buf);  //create uint8array as viewer
  for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; //convert to octet
  return buf;    
}
    function task(i,totalIssues) { 
      setTimeout(function() { 
        let email=document.getElementsByClassName("aui-avatar aui-avatar-small")[1]
        .getElementsByTagName("img")[0].getAttribute("alt");
        console.log(`email ${email}`);
        let status=document.getElementsByClassName("wrap")[1]
        .getElementsByTagName("span")[0].getElementsByTagName("span")[0].innerText;
        console.log(`status ${status}`);
        let totalSalesForce=document.getElementsByClassName("wrap")[9]
        .getElementsByTagName("a");
        let totalSalesForces="";
        for(let i=0;i<totalSalesForce.length;i++)
        {
          console.log(`Sales force ${totalSalesForce[i].getAttribute("href")}`);
          totalSalesForces+="\n\n"+totalSalesForce[i].getAttribute("href");
        }
        let description=document.getElementsByClassName("user-content-block")[0]
        .getElementsByTagName("p")[0].innerText;
        console.log(`description ${description}`);
        let allComments=document.getElementsByClassName("issuePanelContainer")[0]
        .getElementsByClassName("action-body");
        let comments="";
        for(let i=0;i<allComments.length;i++)
        {
          let comment="";
          let paragraphs=allComments[i].getElementsByTagName("p");
          for(let j=0;j<paragraphs.length;j++)
          {
            comment+=" "+paragraphs[j].innerText;
          }
          console.log(`comment ${i+1} ${comment}`);
          comments+="\n\n"+comment;
        }
        if(i+1!=totalIssues)
        {
        document.getElementsByClassName("splitview-issue-link")[i+1].click();
        }
        console.log(`-----------------------------------------------------------------`); 
        excelData.push({"email":email,"status":status,"totalSalesForces":totalSalesForces,
          "description":description,"comments":comments});
      }, 9000 * i); 
    } 
    return 'satz';
}

init();
