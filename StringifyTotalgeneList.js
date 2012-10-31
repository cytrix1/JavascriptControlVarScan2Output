var excApp = new ActiveXObject("Excel.Application"); 
excApp.visible = false;
var excBook = excApp.Workbooks.open("C:\\Users\\nam\\Dropbox\\Research projects\\Small Cell Lung Cancer ºÐ¼®\\Smallceltemp.xlsx");
var ourGene = new Array();
for (var i=2; excBook.WorkSheets(1).Cells(i, 1).value != undefined ; ++i) {
  ourGene[i-2] = excBook.WorkSheets(1).Cells(i, 1).value; 
}
console.log("(\"" + ourGene.join("\", \"") + "\")");