const fs = require('fs');

enum Command {
    XLS_TO_JSON_REPORT_SERVICE = 'xlstojsonreportservice',
    XLS_TO_JSON_MAIL_SERVICE = 'xlstojsonmailservice',
}

const xlsToJsonReportService = async (filePath: string) => {
    const Excel = require('exceljs');
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(2);
    const getCell = (i: number, j: number): string => worksheet.getRow(i).getCell(j).text;

    const langs: string[] = [];

    for (let j = 2; ; j++) {
        const lang = getCell(1, j);
        if (lang) {
            langs.push(lang);
            continue;
        }

        break;
    }

    const json = {};
    for (let langIndex = 0; langIndex < langs.length; langIndex++) {
        const lang = {};

        for (let i = 2; ; i++) {
            const key = getCell(i, 1);
            if (key) {
                const val = getCell(i, 2 + langIndex);
                if (val) {
                    lang[key] = val;
                }
                continue;
            }

            break;
        }
        json[langs[langIndex]] = lang
    }
    fs.writeFileSync(`${__dirname}/../output_json/report_translations.json`, JSON.stringify(json, null, 2) + '\n');
};

const xlsToJsonMailService = async (filePath: string) => {
    const Excel = require('exceljs');
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);

    const getCell = (i: number, j: number): string => worksheet.getRow(i).getCell(j).text;

    const langs: string[] = [];
    for (let j = 2; ; j++) {
        const lang = getCell(1, j);
        if (lang) {
            langs.push(lang);
            continue;
        }

        break;
    }
    const json = {}
    for (let key = 2; ; key++) {
        const jsonKey = getCell(key, 1)
        if (!jsonKey) {
            break;
        }
        const entry = {}
        for (let langIndex = 0; langIndex < langs.length; langIndex++) {
            const language = langs[langIndex]
            const translation = getCell(key, langIndex + 2)
            if (translation) {
                entry[language] = translation
            }
        }
        json[jsonKey] = entry
    }
    fs.writeFileSync(`${__dirname}/../output_json/mail_translations.json`, JSON.stringify(json, null, 2) + '\n');
};

const command = process.argv[2];
const arg = process.argv[3];

switch (command) {
    case Command.XLS_TO_JSON_MAIL_SERVICE:
        xlsToJsonMailService(arg);
        break;
    case Command.XLS_TO_JSON_REPORT_SERVICE:
        xlsToJsonReportService(arg);
        break;
    default:
        console.log('unknown command.');
}
