const fs = require('fs');

const files = ['extracted_word_data.txt', 'extracted_excel_data.txt'];

files.forEach(file => {
    try {
        if (fs.existsSync(file)) {
            const content = fs.readFileSync(file, 'utf16le');
            fs.writeFileSync(file, content, 'utf8');
            console.log(`Converted ${file} to UTF-8`);
        }
    } catch (e) {
        console.error(`Error converting ${file}:`, e.message);
    }
});
