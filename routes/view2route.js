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
    debugger;
    var fileUpload = reader1.readFile('./SAMPLE DATA DEMO.xlsx') 
    const XLSX = require('xlsx'); 
    const data = XLSX.readFile('./SAMPLE DATA DEMO.xlsx'); 
    ProcessExcel(data)
   
    //let data = [] 
  
    //const sheets = file.SheetNames 


function clamp_range(range) {
        if(range.e.r >= (1<<20)) range.e.r = (1<<20)-1;
        if(range.e.c >= (1<<14)) range.e.c = (1<<14)-1;
        return range;
    }
    
    
    function delete_cols(ws, start_col, ncols) {
        if(!ws) throw new Error("operation expects a worksheet");
        var dense = Array.isArray(ws);
        if(!ncols) ncols = 1;
        if(!start_col) start_col = 0;
    
        /* extract original range */
        var range = XLSX.utils.decode_range(ws["!ref"]);
        var R = 0, C = 0;
    
        var formula_cb = function($0, $1, $2, $3, $4, $5) {
            var _R = XLSX.utils.decode_row($5), _C = XLSX.utils.decode_col($3);
            if(_C >= start_col) {
                _C -= ncols;
                if(_C < start_col) return "#REF!";
            }
            return $1+($2=="$" ? $2+$3 : XLSX.utils.encode_col(_C))+($4=="$" ? $4+$5 : XLSX.utils.encode_row(_R));
        };
    
        var addr, naddr;
        for(C = start_col + ncols; C <= range.e.c; ++C) {
            for(R = range.s.r; R <= range.e.r; ++R) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                naddr = XLSX.utils.encode_cell({r:R, c:C - ncols});
                if(!ws[addr]) { delete ws[naddr]; continue; }
                if(ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
                ws[naddr] = ws[addr];
            }
        }
        for(C = range.e.c; C > range.e.c - ncols; --C) {
            for(R = range.s.r; R <= range.e.r; ++R) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                delete ws[addr];
            }
        }
        for(C = 0; C < start_col; ++C) {
            for(R = range.s.r; R <= range.e.r; ++R) {
                addr = XLSX.utils.encode_cell({r:R, c:C});
                if(ws[addr] && ws[addr].f) ws[addr].f = ws[addr].f.replace(crefregex, formula_cb);
            }
        }
    
        /* write new range */
        range.e.c -= ncols;
        if(range.e.c < range.s.c) range.e.c = range.s.c;
        ws["!ref"] = XLSX.utils.encode_range(clamp_range(range));
    
        /* merge cells */
        if(ws["!merges"]) ws["!merges"].forEach(function(merge, idx) {
            var mergerange;
            switch(typeof merge) {
                case 'string': mergerange = XLSX.utils.decode_range(merge); break;
                case 'object': mergerange = merge; break;
                default: throw new Error("Unexpected merge ref " + merge);
            }
            if(mergerange.s.c >= start_col) {
                mergerange.s.c = Math.max(mergerange.s.c - ncols, start_col);
                if(mergerange.e.c < start_col + ncols) { delete ws["!merges"][idx]; return; }
                mergerange.e.c -= ncols;
                if(mergerange.e.c < mergerange.s.c) { delete ws["!merges"][idx]; return; }
            } else if(mergerange.e.c >= start_col) mergerange.e.c = Math.max(mergerange.e.c - ncols, start_col);
            clamp_range(mergerange);
            ws["!merges"][idx] = mergerange;
        });
        if(ws["!merges"]) ws["!merges"] = ws["!merges"].filter(function(x) { return !!x; });
    
        /* cols */
        if(ws["!cols"]) ws["!cols"].splice(start_col, ncols);
    }         
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
function add_cell_to_sheet(worksheet, address, value) {
    /* cell object */
    var cell = {t:'?', v:value};

    /* assign type */
    if(typeof value == "string") cell.t = 's'; // string
    else if(typeof value == "number") cell.t = 'n'; // number
    else if(value === true || value === false) cell.t = 'b'; // boolean
    else if(value instanceof Date) cell.t = 'd';
    else throw new Error("cannot store value");

    /* add to worksheet, overwriting a cell if it exists */
    worksheet[address] = cell;

    /* find the cell range */
    var range = XLSX.utils.decode_range(worksheet['!ref']);
    var addr = XLSX.utils.decode_cell(address);

    /* extend the range to include the new cell */
    if(range.s.c > addr.c) range.s.c = addr.c;
    if(range.s.r > addr.r) range.s.r = addr.r;
    if(range.e.c < addr.c) range.e.c = addr.c;
    if(range.e.r < addr.r) range.e.r = addr.r;

    /* update range */
    worksheet['!ref'] = XLSX.utils.encode_range(range);
}
function ProcessExcel(workbook) {
    
    var firstSheet = workbook.SheetNames[0];
    var sheet = workbook.Sheets[firstSheet]; // get the first worksheet
    console.log("heeeeeee")
    var comment = "";
    var comarray = [];
    var lastcell = "";
    var viewcell = "";
    debugger;
    var range = XLSX.utils.decode_range(sheet['!ref']); // get the range
    for(var R = range.s.r; R <= range.e.r; ++R) {
      for(var C = range.s.c; C <= range.e.c; ++C) {
        /* find the cell object */
        //console.log('Row : ' + R);
        //console.log('Column : ' + C);
        
        var cellref = XLSX.utils.encode_cell({c:C, r:R}); // construct A1 reference for cell
        if(!sheet[cellref]) continue; // if cell doesn't exist, move on
        var cell = sheet[cellref];
        if (C == range.e.c && cellref[0] > lastcell) {
            lastcell = cellref[0];
        }
        viewcell = lastcell.substring(0,lastcell.length-1) + String.fromCharCode(lastcell.charCodeAt(lastcell.length-1)+1);
    
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
       // console.log("commmmmmmn " ,comarray[i]);
        delete_row(sheet, comarray[i]); 
        delete_row(sheet, 0)    ;              
        };
    var ranges = XLSX.utils.decode_range(sheet['!ref']); // get the range
    count = 0;
    for(var R = ranges.s.r+1; R <= ranges.e.r; ++R) {
       // console.log(" ref" , lastcell+R);
       // console.log("sheeet with ref" , sheet[lastcell+R]);
        if (count == 0) {
            add_cell_to_sheet(sheet, viewcell+R, "TRACKING");
            count = -1;
            continue;
            
        }
        checkcell = lastcell.substring(0,lastcell.length-1) + String.fromCharCode(lastcell.charCodeAt(lastcell.length-1)-3);
        
        if (sheet[checkcell+R] != undefined) {
            
            if (sheet[lastcell+R].v <=3.0) {
                add_cell_to_sheet(sheet, viewcell+R, "ON TRACK")
                
            }
            else if (sheet[lastcell+R].v >3.0 && sheet[lastcell+R].v<=10.0) {
                
                add_cell_to_sheet(sheet, viewcell+R, "DELAYED")
            } else {
                add_cell_to_sheet(sheet, viewcell+R, "OVERDUE")
            }   
    }
        
    };
    //Read all rows from First Sheet into an JSON array.
    //console.log(workbook.Sheets[firstSheet])
    delete_cols(sheet,2, 0);
    var excelRows = XLSX.utils.sheet_to_row_object_array(sheet, {blankrows: false});
    
    type = typeof excelRows;
    res.setHeader('Content-Type', 'application/json');
    //console.log(excelRows);
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