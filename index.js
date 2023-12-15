const express = require('express');
const ExcelJS = require('exceljs');
const { Readable } = require('stream');

const app = express();

function generateSampleData(numRecords) {
    const data = [];
    for (let i = 1; i <= numRecords; i++) {
        data.push({
            id: i,
            name: `User ${i}`,
            email: `user${i}@example.com`
        });
    }
    return data;
}

async function exportToExcelWithMultipleSheets(data, batchSize) {
    const workbook = new ExcelJS.Workbook();

    const numBatches = Math.ceil(data.length / batchSize);

    for (let batchIndex = 0; batchIndex < numBatches; batchIndex++) {
        const batchData = data.slice(batchIndex * batchSize, (batchIndex + 1) * batchSize);

        const worksheet = workbook.addWorksheet(`Sheet ${batchIndex + 1}`);

        worksheet.columns = [
            { header: 'ID', key: 'id', width: 10 },
            { header: 'Name', key: 'name', width: 30 },
            { header: 'Email', key: 'email', width: 40 }
        ];

        batchData.forEach(record => {
            worksheet.addRow(record);
        });
    }

    return await workbook.xlsx.writeBuffer();
}

app.get('/download-excel', async (req, res) => {
    const sampleData = generateSampleData(400000);
    const batchSize = 100000;

    try {
        const workbookBuffer = await exportToExcelWithMultipleSheets(sampleData, batchSize);
        const fileName = 'large_data.xlsx';

        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        const bufferStream = new Readable();
        bufferStream.push(workbookBuffer);
        bufferStream.push(null);

        bufferStream.pipe(res);
    } catch (err) {
        console.error('Error exporting data:', err);
        res.status(500).send('Error exporting data');
    }
});

const port = 3000;
app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
