const mammoth = require('mammoth');
const fs = require('fs');
const path = require('path');

const files = [
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE YAOUNDE FEVRIER 2025.doc.docx',
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE YAOUNDE MARS 2025 AGENT ADMINISTRATIF.doc',
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE YAOUNDE AVRIL 2025 AGENT ADMINISTRATIF.doc',
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE YAOUNDE MAI 2025 AGENT ADMINISTRATIF.doc',
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE COMMERCIAL YAOUNDE JUIN 2025.docx',
    'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ACTIVITE COMMERCIAL YAOUNDE JUILLET 2025.docx'
];

async function readWordDocument(filePath) {
    try {
        const result = await mammoth.extractRawText({path: filePath});
        return result.value;
    } catch (err) {
        if (filePath.endsWith('.doc')) {
            return `[Old format .doc - cannot read directly]: ${filePath}`;
        }
        return `Error reading file: ${filePath} - ${err.message}`;
    }
}

async function processAllFiles() {
    for (const file of files) {
        if (fs.existsSync(file)) {
            console.log('\n========================================');
            console.log('FILE:', path.basename(file));
            console.log('========================================\n');
            
            const content = await readWordDocument(file);
            console.log(content);
        } else {
            console.log('\n========================================');
            console.log('FILE NOT FOUND:', path.basename(file));
            console.log('========================================');
        }
    }
}

processAllFiles().catch(console.error);
