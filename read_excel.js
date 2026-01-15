const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const files = [
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES SEPTEMBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES OCTOBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES NOVEMBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES DECEMBRE 2025.xlsx'
];

files.forEach(file => {
    if (fs.existsSync(file)) {
        console.log('\n========================================');
        console.log('FILE:', path.basename(file));
        console.log('========================================\n');
        
        try {
            const workbook = XLSX.readFile(file);
            const sheetNames = workbook.SheetNames;
            
            sheetNames.forEach(sheetName => {
                console.log('--- Sheet:', sheetName, '---');
                const sheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
                
                console.log('Rows:', data.length);
                console.log('');
                
                // Show first 60 rows
                data.slice(0, 60).forEach((row, rowIndex) => {
                    const joined = row.join(' | ').trim();
                    if (joined) {
                        console.log(joined);
                    }
                });
                
                console.log('');
            });
        } catch (err) {
            console.error('Error reading file:', err.message);
        }
    } else {
        console.log('\nFile not found:', file);
    }
});
