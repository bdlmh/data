function handleFile() {
    const input = document.getElementById('originalFile');
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'binary'});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, {header:1});

        let transformedRecords = [];
        let recordCount = 1; // 用于生成编号，从A1开始

        // 跳过标题行，从第二行开始处理
        for (let i = 1; i < json.length; i++) {
            const row = json[i];
            const name = row[0]; // “名称”在第一列
            const quantityPiecesStr = row[1]; // “数量*件数”在第二列

            if (quantityPiecesStr) {
                const quantityPieces = quantityPiecesStr.split(/；|;/);
                quantityPieces.forEach(qp => {
                    const [quantity, pieces] = qp.split('*');
                    for (let j = 0; j < parseInt(pieces, 10); j++) {
                        // 为每条记录生成编号，格式为 'A' + 当前记录的序号
                        const id = `A${recordCount}`;
                        transformedRecords.push({ '编号': id, '名称': name, '数量': quantity });
                        recordCount++; // 递增记录计数
                    }
                });
            }
        }

        // 将转换后的记录数组转换回工作表
        //const newSheet = XLSX.utils.json_to_sheet(transformedRecords, {skipHeader: true});
        const newSheet = XLSX.utils.json_to_sheet(transformedRecords, {origin: 0});
        // 添加表头
        XLSX.utils.sheet_add_aoa(newSheet, [['编号', '名称', '数量']], {origin: 'A1'});

        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "TransformedData");

        // 保存新工作簿为 Excel 文件
        XLSX.writeFile(newWorkbook, "transformed.xlsx");
    };

    // 读取上传的文件
    const file = input.files[0];
    reader.readAsBinaryString(file);
}

// 假设的固定表头，您需要根据 path7.xlsx 文件的实际表头进行调整
const fixedHeaders = ['Header1', 'Header2', 'Header3'];

// 从 handleFile 函数中提取出的转换逻辑
function handleAndTranslateFile() {
    const originalFile = document.getElementById('originalFile').files[0];
    const translationFile = document.getElementById('translationFile').files[0];

    // 确保文件已上传
    if (!originalFile || !translationFile) {
        alert("请确保已上传原始文件和中间文件！");
        return;
    }

    // 读取并构建中英文对照表

    const translationReader = new FileReader();
    translationReader.onload = function(e) {
        const workbook = XLSX.read(e.target.result, {type: 'binary'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const translationJson = XLSX.utils.sheet_to_json(sheet);
        const translationMap = {};
        translationJson.forEach(row => {
            translationMap[row['部件名称']] = row['Name'];  // 调整以匹配翻译文件的实际列名
        });
    

        // 读取原始文件
        const originalReader = new FileReader();
        originalReader.onload = function(e) {
            const data = e.target.result;
            const workbook = XLSX.read(data, {type: 'binary'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, {header:1});

            // 应用转换逻辑
            let transformedRecords = [];
            let recordCount = 0;  // 用于生成编号，从A1开始

            // 跳过标题行，从第二行开始处理
            for (let i = 1; i < json.length; i++) {
                const row = json[i];
                const name = row[0];  // “名称”在第一列
                const quantityPiecesStr = row[1];  // “数量*件数”在第二列

                if (quantityPiecesStr) {
                    const quantityPieces = quantityPiecesStr.split(/；|;/);
                    quantityPieces.forEach(qp => {
                        const [quantity, pieces] = qp.split('*');
                        for (let j = 0; j < parseInt(pieces, 10); j++) {
                            const id = `A${++recordCount}`;
                            transformedRecords.push({ '编号': id, '名称': name, '数量': quantity });
                        }
                    });
                }
            }

            // 使用中英文对照表转换转换后的数据
            const translatedRecords = transformedRecords.map(record => {
                const translatedRecord = {};
                Object.keys(record).forEach(key => {
                    translatedRecord[key] = translationMap[record[key]] || record[key];
                });
                return translatedRecord;
            });

            // 生成并下载新的英文版 Excel 文件
            const newSheet = XLSX.utils.json_to_sheet(translatedRecords, {origin: 0});
            //const newSheet = XLSX.utils.json_to_sheet(transformedRecords, {origin: 0});
            XLSX.utils.sheet_add_aoa(newSheet, [['CARTON NO', 'CONTENTS', 'QUANTITY']], {origin: 'A1'});
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, "EnglishVersion");
            XLSX.writeFile(newWorkbook, "EnglishVersion.xlsx");
        };
        originalReader.readAsBinaryString(originalFile);
    };
    translationReader.readAsBinaryString(translationFile);
}

function handlefinalFile() {
    const originalFile = document.getElementById('originalFile').files[0];
    const translationFile = document.getElementById('translationFile').files[0];

    const reader = new FileReader();

    reader.onload = function(e) {
        const workbook1 = XLSX.read(e.target.result, { type: 'binary' });
        const originalData = XLSX.utils.sheet_to_json(workbook1.Sheets[workbook1.SheetNames[0]], {header:1});

        const reader2 = new FileReader();

        reader2.onload = function(e) {
            const workbook2 = XLSX.read(e.target.result, { type: 'binary' });
            const translationData = XLSX.utils.sheet_to_json(workbook2.Sheets[workbook2.SheetNames[0]], {header:1});

            const finalData = processData(originalData, translationData);
            downloadFinalFile(finalData);
        };

        reader2.readAsBinaryString(translationFile);
    };

    reader.readAsBinaryString(originalFile);
}

function processData(originalData, translationData) {
    const translationMap = {};
    translationData.forEach(row => {
        translationMap[row[0]] = row[1];
    });

    const finalData = [['中文名品名', '英文品名', '件数', '数量/箱']];
    originalData.slice(1).forEach(row => {
        const parts = row[1].split(/；|;/);
        parts.forEach((part, partIndex) => {
            const [quantity, count] = part.split('*');
            finalData.push([
                row[0],
                translationMap[row[0]] || 'No translation found',
                parseInt(count, 10),
                parseInt(quantity, 10),
                partIndex > 0 // 标记是否为新起的行
            ]);
        });
    });

    return finalData;
}

function downloadFinalFile(data) {
    const ws = XLSX.utils.aoa_to_sheet(data.map(row => row.slice(0, 4))); // 只取前4列数据，不包括标记列

    // 为新起的行设置样式
    data.forEach((row, rowIndex) => {
        if (row[4]) { // 使用之前添加的标记检查是否为新起的行
            for (let col = 0; col < 4; col++) { // 为每个单元格设置样式
                const cellRef = XLSX.utils.encode_cell({r: rowIndex, c: col});
                ws[cellRef].s = {
                    fill: { fgColor: { rgb: "FFFF00" }} // 设置背景色为黄色
                };
            }
        }
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Final Data');
    XLSX.writeFile(wb, 'finalFile.xlsx');
}





