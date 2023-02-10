const Excel = require('exceljs');
const fs = require('fs');

const getDOCTInvLevel = async (file, sectors) => {
    const workbook = new Excel.Workbook();
    const wb = await workbook.xlsx.readFile(file);
    const result = [];

    sectors.map(name => {
        let skuText;

        name !== 'Prime' ? skuText = 'FBA No. of  SKU\'S' : skuText = 'No. of SKUs';
        wb.eachSheet(function(worksheet) {
            if (worksheet.name === name) {
                const rows = worksheet.getSheetValues();
                let skuId = 0, date;
    
                for (let x in rows) {
                    const row = rows[x];
                    if (x < 2) {
                        for (let i in row) {
                            if (row[i].toString().includes(skuText)) {
                                skuId = i;
                                date = row[parseInt(skuId) - 5];
                            };
                        };
                    } else {
                        const resObj = {
                            date: date,
                            category: worksheet.name,
                            supName: row[2],
                            skuNo: typeof row[skuId] === "object" ? row[skuId].result : row[skuId],
                            unitsNo: typeof row[parseInt(skuId) + 1] === "object" ? row[parseInt(skuId) + 1].result : row[parseInt(skuId) + 1],
                            totalCost: typeof row[parseInt(skuId) + 2] === "object" ? row[parseInt(skuId) + 2].result : row[parseInt(skuId) + 2]
                        };
    
                        result.push(resObj);
                    };

                };
            };
        });
    });

    console.log(file);
    return result;
};

const master = async () => {
    const Workbook = new Excel.Workbook();
    const Worksheet = Workbook.addWorksheet("Results");

    Worksheet.columns = [
        { key: 'date', header: 'Date'},
        { key: 'category', header: 'Category'},
        { key: 'supName', header: 'Supplier Name'},
        { key: 'skuNo', header: 'No. SKU'},
        { key: 'unitsNo', header: 'No. Units'},
        { key: 'totalCost', header: 'Total Cost'},
    ];

    let body = [];
    const _dirDOCT = "D:/_Jomael/Maintenance/DOCT/Inventory Levels Updating/DOCT (sum) Files/August/";
    const categories = ['FBA', 'Prime'];
    const mmDDList = ['08.29', '08.30', '08.31'];

    await Promise.all(mmDDList.map(async (day) => {
        let _fileDOCT = `DOCT (sum) ${day}.2022.xlsx`;
        let fileName = `${_dirDOCT}${_fileDOCT}`;
        
        let data = await getDOCTInvLevel(fileName, categories);
        body.push(data);
    }));

    for (let i = 0; i < body.length; i++) {
        await body[i].map(data => {
            Worksheet.addRow(data);
        });
    };

    await Workbook.xlsx.writeFile('D:/_Jomael/Maintenance/DOCT/Inventory Levels Updating/Results/H-Exports.xlsx');
    console.log('Complete!');
};

master();