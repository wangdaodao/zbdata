const axios = require('axios');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');

// 定义异步函数，用于获取网页内容
async function fetchPageContent(url) {
    try {
        const response = await axios.get(url);
        return response.data;
    } catch (error) {
        console.error(`请求 ${url} 内容时出错:`, error.message);
        throw error;
    }
}

// 清理数据，移除特殊字符和所有空格
function cleanData(data) {
    if (typeof data === 'string') {
        // 移除不可见字符、特殊字符和所有空格
        return data.replace(/[\x00-\x1F\x7F-\x9F]/g, '').replace(/\s/g, '');
    }
    return data;
}

// 定义函数，用于提取 class 为 table-box 和 table-line 的 div 中的数据
function extractDataFromTables(html) {
    const $ = cheerio.load(html);
    const allTableData = [];

    // 提取 class 为 table-box 的表格数据
    $('.table-box').each((boxIndex, box) => {
        const tableData = [];
        $(box).find('li.tr').each((rowIndex, row) => {
            const rowData = [];
            $(row).find('div.td').each((cellIndex, cell) => {
                const cellText = $(cell).text();
                rowData.push(cleanData(cellText));
            });
            tableData.push(rowData);
        });
        allTableData.push(tableData);
    });

    // 提取 class 为 table-line 的表格数据
    $('.table-line').each((lineIndex, lineTable) => {
        const tableData = [];
        $(lineTable).find('tr').each((rowIndex, row) => {
            const rowData = [];
            $(row).find('td, th').each((cellIndex, cell) => {
                const cellText = $(cell).text();
                rowData.push(cleanData(cellText));
            });
            tableData.push(rowData);
        });
        allTableData.push(tableData);
    });

    return allTableData;
}

// 定义异步函数，用于生成单个 Excel 文件
async function generateSingleExcel(regionName, regionData) {
    const workbook = new ExcelJS.Workbook();

    regionData.forEach((table, tableIndex) => {
        // 限制工作表名称长度为 31 个字符，避免 Excel 不支持长名称
        const sheetName = `Table_${tableIndex + 1}`.substring(0, 31);
        const worksheet = workbook.addWorksheet(sheetName);
        table.forEach((row, rowIndex) => {
            row.forEach((cell, cellIndex) => {
                worksheet.getRow(rowIndex + 1).getCell(cellIndex + 1).value = cell;
            });
        });
    });

    const fileName = `${regionName.replace(/[/\\?%*:|"<>]/g, '')}.xlsx`;
    try {
        await workbook.xlsx.writeFile(fileName);
        console.log(`Excel 文件 ${fileName} 生成成功`);
    } catch (error) {
        console.error(`生成 ${fileName} 时出错:`, error);
    }
}

// 主函数
async function main() {
    const regions = [
        { name: '全国', url: 'https://hq.zhaobiao.cn/data_0_0.html' },
        { name: '北京', url: 'https://hq.zhaobiao.cn/data_110000_0.html' },
        { name: '河北', url: 'https://hq.zhaobiao.cn/data_130000_0.html' },
        { name: '上海', url: 'https://hq.zhaobiao.cn/data_310000_0.html' },
        { name: '江苏', url: 'https://hq.zhaobiao.cn/data_320000_0.html' },
        { name: '浙江', url: 'https://hq.zhaobiao.cn/data_330000_0.html' },
        { name: '山东', url: 'https://hq.zhaobiao.cn/data_370000_0.html' },
        { name: '湖南', url: 'https://hq.zhaobiao.cn/data_430000_0.html' },
        { name: '广东', url: 'https://hq.zhaobiao.cn/data_440000_0.html' },
        { name: '重庆', url: 'https://hq.zhaobiao.cn/data_500000_0.html' },
        { name: '四川', url: 'https://hq.zhaobiao.cn/data_510000_0.html' },
    ];

    for (const region of regions) {
        try {
            console.log(`开始处理 ${region.name} (${region.url})...`);
            const html = await fetchPageContent(region.url);
            const extractedData = extractDataFromTables(html);
            await generateSingleExcel(region.name, extractedData);
            console.log(`${region.name} 数据处理完成。`);
        } catch (error) {
            console.error(`处理 ${region.name} 数据时出错:`, error);
        }
    }
}

main();