
const ExcelJs=require('exceljs');

async function excelTest()
{
    
    let output={row:-1};

const workbook=new ExcelJs.Workbook();

await workbook.xlsx.readFile('C:/Users/KeshojuKavya/OneDrive - IBM/Personal/Career/IBM/Playwright/download.xlsx') 

    const worksheet=workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row,rowNumber)=>
    {
        row.eachCell((cell,colNumber)=>
        {
            if(cell.value==='Apple')
            {
                output.row=rowNumber;
                output.column=colNumber;
            }
        })
    })
    const cell=worksheet.getCell(output.row,output.column);
    cell.value="Republic";
    await workbook.xlsx.writeFile('C:/Users/KeshojuKavya/OneDrive - IBM/Personal/Career/IBM/Playwright/download.xlsx');


}
excelTest();