import * as path from 'path';
import * as Excel from 'exceljs'; // node package

const filePath = process.argv[2]; //'testmerge.xlsx'; // path.resolve(__dirname, 'testmerge.xlsx');
const main = async () => {
    handleWorkbook(filePath);
};
  
const handleWorkbook = async (path: string) => {
    const workbook = new Excel.Workbook();
    const content = await workbook.xlsx.readFile(path);
  
    content.worksheets.forEach( (sheet, index) => {
        handleWorksheet(sheet, index);
    });
}

function toColumnName(num: number) {
    for (var ret = '', a = 1, b = 26; (num -= a) >= 0; a = b, b *= 26) {
      ret = String.fromCharCode(( (num % b) / a) + 65 ) + ret;
    }
    return ret;
  }

const handleWorksheet = async (worksheet: Excel.Worksheet, index: number) => {
    const rowCount: number = worksheet.rowCount;
    const colCount: number = worksheet.columnCount;
    
    if(!worksheet.hasMerges) {
        console.log("------\nWorksheet number " + (index+1) + " has no merged cells.\n------");
        return;
    }
    console.log("------\nWorksheet number " + (index+1) + "\n------");

    const mergeMap = new Map<string, string>();

    let currCell: Excel.Cell;
    let parentCell: Excel.Cell;
    for( let i = 1; i <= rowCount; i++ ) {
        for( let j = 1; j <= colCount; j++ ) {
            currCell = worksheet.getCell(i, j);
            if(currCell.isMerged) {
                parentCell = currCell.master;
                mergeMap.set(
                    toColumnName(parseInt(parentCell.col)) + parentCell.row, // parent location
                    toColumnName(j) + i // current location, the last time this code runs will by definition be the end of the merge range
                );
            }
        }
    }

    mergeMap.forEach( (end, beginning) => {
        console.log(beginning + ":" + end);
    })
}

main().then();