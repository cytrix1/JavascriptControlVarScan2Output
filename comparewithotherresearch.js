Array.prototype.contains = function (element) {
    for (var i = 0; i < this.length; i++) {
        if (this[i] == element) {
            return true;
        }
    }
    return false;
}

	if(!Array.indexOf){
	    Array.prototype.indexOf = function(obj){
	        for(var i=0; i<this.length; i++){
	            if(this[i]==obj){
	                return i;
	            }
	        }
	        return -1;
	    }
	}

var excApp = new ActiveXObject("Excel.Application"); 
excApp.visible = false;
var excBook = excApp.Workbooks.open("C:\\test\\Smallceltemp.xlsx");
//alert(excBook.WorkSheets(1).Cells(2000,1).value);
var ourGene = new Array();
for (var i=2; excBook.WorkSheets(1).Cells(i, 1).value != undefined ; ++i) {
    ourGene[i - 2] = excBook.WorkSheets(1).Cells(i, 1).value;
}

//excBook.WorkSheets(1).Cells(1, 1).value = "Geeene";
sh3 = excBook.WorkSheets(2);
for (var i=1; i<=3 ; ++i)  
  for (var j=2; sh3.Cells(j,2*i -1).value != undefined; ++j) 
	  sh3.Cells(j, 2*i).value = ourGene.indexOf(sh3.Cells(j, 2*i-1).value) + 1;
excBook.SaveAs("C:\\test\\smallCellcomother.xlsx");    
excBook.Close(true);
excApp.Application.Quit();
excApp.quit();
excApp = null;
//alert("Completed");
