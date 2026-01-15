const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const WordExtractor = require('word-extractor');

const directoryPath = 'd:/RAPPORT MENSUEL HACC 2025';

async function readFiles() {
    const files = fs.readdirSync(directoryPath);
    const wordExtractor = new WordExtractor();

    for (const file of files) {
        const fullPath = path.join(directoryPath, file);
        if (fs.statSync(fullPath).isDirectory()) continue;

        const ext = path.extname(file).toLowerCase();
        
        if (ext === '.docx') {
            console.log(`\n\n=== FILE: ${file} ===\n`);
            try {
                const result = await mammoth.extractRawText({ path: fullPath });
                console.log(result.value);
            } catch (err) {
                console.error(`Error reading .docx file ${file}:`, err.message);
            }
        } else if (ext === '.doc') {
            console.log(`\n\n=== FILE: ${file} ===\n`);
            try {
                const extracted = await wordExtractor.extract(fullPath);
                console.log(extracted.getBody());
            } catch (err) {
                console.error(`Error reading .doc file ${file}:`, err.message);
            }
        }
    }
}

readFiles().then(() => {
    console.log('\nDone reading all word files.');
}).catch(err => {
    console.error('Fatal error:', err);
});
