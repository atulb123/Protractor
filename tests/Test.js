var Excel = require('exceljs');
function test() {
    
var workbook = new Excel.Workbook();
workbook.xlsx.readFile("C:/Users/acer/Desktop/Book1.xlsx")
    .then(function() {
        let rowcount=workbook.getWorksheet("Sheet1").rowCount;
        let colcount=workbook.getWorksheet("Sheet1").columnCount;
        for(let i=1;i<=rowcount;i++)
        {
            for(let j=1;j<=colcount;j++)
            {
                if(workbook.getWorksheet("Sheet1").getRow(i).getCell(j).value!=null)
                console.log(workbook.getWorksheet("Sheet1").getRow(i).getCell(j).value);
            }
        }
    });
    
}
test();
