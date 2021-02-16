var express = require("express");
const { file } = require("googleapis/build/src/apis/file");
var router = express.Router();
const app = express();



router.get("/", function(req, res, next) {
   //res.send("API is working properly");

    //Reference the FileUpload element.
    //var fileUpload = document.getElementById("fileUpload");
    // Requiring the module 
    const reader1 = require('xlsx') 
    
    // Reading our test file 
    var fileUpload = reader1.readFile('./SAMPLE DATA DEMO.xlsx') 
    const XLSX = require('xlsx'); 
    const data = XLSX.readFile('./SAMPLE DATA DEMO.xlsx'); 
    ProcessExcel(data)
   
    //let data = [] 
  
    //const sheets = file.SheetNames 


           
function ec(r, c){
    return XLSX.utils.encode_cell({r:r,c:c});
};

function delete_row(ws, row_index){
    var variable = XLSX.utils.decode_range(ws["!ref"])
    for(var R = row_index; R <= variable.e.r; ++R){
        for(var C = variable.s.c; C <= variable.e.c; ++C){
            ws[ec(R,C)] = ws[ec(R+1,C)];
}
}
variable.e.r--
ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
};

function ProcessExcel(workbook) {
    
    var firstSheet = workbook.SheetNames[0];
    var sheet = workbook.Sheets[firstSheet]; // get the first worksheet
    console.log("heeeeeee", sheet)
    var comment = "";
    var comarray = [];
    var range = XLSX.utils.decode_range(sheet['!ref']); // get the range
    for(var R = range.s.r; R <= range.e.r; ++R) {
      for(var C = range.s.c; C <= range.e.c; ++C) {
        /* find the cell object */
        console.log('Row : ' + R);
        console.log('Column : ' + C);
        
        var cellref = XLSX.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
        if(!sheet[cellref]) continue; // if cell doesn't exist, move on
        var cell = sheet[cellref];
        //console.log("cellllll", cellref);
        //console.log("typeof", typeof cell);

        var str = cell.v;
        for (var i = 0; i < str.length; i++) {
            if (str[i]== '*') {
                if(str[i+1] == '*'){
                    comment = str;
                    comarray.push(R);
                    break;
                };
            
            };
        }
       
        
      };
    };
    for (var i = 0; i < comarray.length; i++) {
        console.log("commmmmmmn " ,comarray[i]);
        delete_row(sheet, comarray[i]); 
        delete_row(sheet, 0)    ;              
        };
    var ranges = XLSX.utils.decode_range(sheet['!ref']); // get the range
    for(var R = ranges.s.r; R <= ranges.e.r; ++R) {
      for(var C = ranges.s.c; C <= ranges.e.c; ++C) {
        /* find the cell object */
        //console.log('Row : ' + R);
        //console.log('Column : ' + C);
        var cellref = XLSX.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
        if(!sheet[cellref]) continue; // if cell doesn't exist, move on
        var cell = sheet[cellref];
        //console.log(cell.v);
        //console.log(typeof cell.v);
      };
    };
    //Read all rows from First Sheet into an JSON array.
    //console.log(workbook.Sheets[firstSheet])
    var excelRows = XLSX.utils.sheet_to_row_object_array(sheet, {blankrows: false});
    console.log(typeof excelRows);
    //Create a HTML Table element.
    type = typeof excelRows;
    res.setHeader('Content-Type', 'application/json');
    res.json(excelRows);
    /*res.render('index', {
        exceltable : JSON.stringify(excelRows),
    });

    if (excelRows.length > 0) {
            BindTable(excelRows, '#exceltable');
    }
   // document.getElementById("comment").innerHTML = comment;
   //$('#exceltable').show();
   */
};
});



module.exports = router;