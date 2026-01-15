const XLSX = require('xlsx');
const fs = require('fs');

const files = [
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES AOUT 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES SEPTEMBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES OCTOBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES NOVEMBRE 2025.xlsx',
    'D:\\RAPPORT MENSUEL HACC 2025\\REPORTING ADMINISTRATIF-FINANCIER-VENTES DECEMBRE 2025.xlsx'
];

const monthNames = {
    'AOUT': 'Août',
    'SEPTEMBRE': 'Septembre',
    'OCTOBRE': 'Octobre',
    'NOVEMBRE': 'Novembre',
    'DECEMBRE': 'Décembre'
};

function getMonthName(filePath) {
    const upperPath = filePath.toUpperCase();
    for (const [key, value] of Object.entries(monthNames)) {
        if (upperPath.includes(key)) {
            return value;
        }
    }
    return 'Inconnu';
}

function extractSynthesisData(workbook, month) {
    const synthesisSheet = workbook.Sheets['SYNTHÈSE'];
    if (!synthesisSheet) return null;

    const data = XLSX.utils.sheet_to_json(synthesisSheet, {header: 1, defval: ''});
    
    let totalVentes = 0;
    let totalDepenses = 0;
    let soldeNet = 0;
    let totalVersements = 0;

    for (const row of data) {
        if (row.length >= 2) {
            const indicator = row[0] ? row[0].toString().trim() : '';
            const montant = row[1] ? parseFloat(row[1].toString().replace(/[\s]/g, '')) : 0;

            if (indicator.includes('Total Ventes') || indicator.includes('VENTES')) {
                totalVentes = montant;
            } else if (indicator.includes('Total Dépenses') || indicator.includes('DEPENSES')) {
                totalDepenses = montant;
            } else if (indicator.includes('Solde Net')) {
                soldeNet = montant;
            } else if (indicator.includes('Total Versements') || indicator.includes('VERSEMENTS')) {
                totalVersements = montant;
            }
        }
    }

    return { month, totalVentes, totalDepenses, soldeNet, totalVersements };
}

function analyzeSales(workbook, month) {
    const salesSheet = workbook.Sheets['VENTES'];
    if (!salesSheet) return { month, productSales: {}, clientSales: {} };

    const data = XLSX.utils.sheet_to_json(salesSheet, {header: 1, defval: ''});
    const productSales = {};
    const clientSales = {};

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length >= 7) {
            const client = row[2] ? row[2].toString().trim() : '';
            const product = row[3] ? row[3].toString().trim() : '';
            const amount = row[6] ? parseFloat(row[6].toString().replace(/[\s]/g, '')) || 0 : 0;

            if (client && amount > 0) {
                clientSales[client] = (clientSales[client] || 0) + amount;
            }

            if (product && amount > 0) {
                productSales[product] = (productSales[product] || 0) + amount;
            }
        }
    }

    return { month, productSales, clientSales };
}

function createAnnualReport() {
    const monthlyData = [];
    const annualProductSales = {};
    const annualClientSales = {};

    files.forEach(file => {
        if (fs.existsSync(file)) {
            const month = getMonthName(file);
            const workbook = XLSX.readFile(file);

            const synthesisData = extractSynthesisData(workbook, month);
            const salesData = analyzeSales(workbook, month);

            if (synthesisData) {
                monthlyData.push(synthesisData);
            }

            if (salesData.productSales) {
                for (const [product, amount] of Object.entries(salesData.productSales)) {
                    annualProductSales[product] = (annualProductSales[product] || 0) + amount;
                }
            }

            if (salesData.clientSales) {
                for (const [client, amount] of Object.entries(salesData.clientSales)) {
                    annualClientSales[client] = (annualClientSales[client] || 0) + amount;
                }
            }
        }
    });

    const annualTotals = monthlyData.reduce((acc, month) => ({
        totalVentes: acc.totalVentes + month.totalVentes,
        totalDepenses: acc.totalDepenses + month.totalDepenses,
        soldeNet: acc.soldeNet + month.soldeNet,
        totalVersements: acc.totalVersements + month.totalVersements
    }), { totalVentes: 0, totalDepenses: 0, soldeNet: 0, totalVersements: 0 });

    const avgMonthlySales = annualTotals.totalVentes / monthlyData.length;
    const ratioVersements = ((annualTotals.totalVersements / annualTotals.totalVentes) * 100).toFixed(2);
    const ratioDepenses = ((annualTotals.totalDepenses / annualTotals.totalVentes) * 100).toFixed(2);

    const wb = XLSX.utils.book_new();

    const summaryData = [
        ['', 'RAPPORT ANNUEL D\'ACTIVITÉ 2025 - HACC YAOUNDÉ', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['INDICATEURS FINANCIERS ANNUELS', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['Total des Ventes', formatNumber(annualTotals.totalVentes), 'FCFA', '', '', ''],
        ['Total des Dépenses', formatNumber(annualTotals.totalDepenses), 'FCFA', '', '', ''],
        ['Solde Net Annuel', formatNumber(annualTotals.soldeNet), 'FCFA', '', '', ''],
        ['Versements Bancaires', formatNumber(annualTotals.totalVersements), 'FCFA', '', '', ''],
        ['', '', '', '', '', ''],
        ['INDICATEURS CLES DE PERFORMANCE', '', '', '', '', ''],
        ['', '', '', '', '', ''],
        ['Moyenne Mensuelle Ventes', formatNumber(avgMonthlySales), 'FCFA', '', '', ''],
        ['Ratio Versements/Ventes', ratioVersements, '%', '', '', ''],
        ['Ratio Dépenses/Ventes', ratioDepenses, '%', '', '', ''],
        ['Marge Bénéficiaire', ((1 - (annualTotals.totalDepenses / annualTotals.totalVentes)) * 100).toFixed(2), '%', '', '', '']
    ];
    
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary['!cols'] = [{wch: 30}, {wch: 25}, {wch: 10}, {wch: 15}, {wch: 15}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'SYNTHÈSE');

    const monthlyAnalysis = [
        ['Mois', 'Total Ventes', 'Total Dépenses', 'Solde Net']
    ];

    monthlyData.forEach(month => {
        monthlyAnalysis.push([
            month.month,
            month.totalVentes,
            month.totalDepenses,
            month.soldeNet
        ]);
    });

    monthlyAnalysis.push([
        'TOTAL ANNUEL',
        annualTotals.totalVentes,
        annualTotals.totalDepenses,
        annualTotals.soldeNet
    ]);

    const wsMonthly = XLSX.utils.aoa_to_sheet(monthlyAnalysis);
    wsMonthly['!cols'] = [{wch: 20}, {wch: 20}, {wch: 20}, {wch: 20}];
    XLSX.utils.book_append_sheet(wb, wsMonthly, 'ANALYSE MENSUELLE');

    const sortedProducts = Object.entries(annualProductSales)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const productData = [
        ['Rang', 'Produit', 'Chiffre d\'Affaires', '% du Total']
    ];

    sortedProducts.forEach((item, index) => {
        const [product, amount] = item;
        const percentage = ((amount / annualTotals.totalVentes) * 100).toFixed(2);
        productData.push([
            index + 1,
            product,
            amount,
            percentage + '%'
        ]);
    });

    const wsProducts = XLSX.utils.aoa_to_sheet(productData);
    wsProducts['!cols'] = [{wch: 8}, {wch: 60}, {wch: 25}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, wsProducts, 'TOP PRODUITS');

    const sortedClients = Object.entries(annualClientSales)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);

    const clientData = [
        ['Rang', 'Client', 'Chiffre d\'Affaires', '% du Total']
    ];

    sortedClients.forEach((item, index) => {
        const [client, amount] = item;
        const percentage = ((amount / annualTotals.totalVentes) * 100).toFixed(2);
        clientData.push([
            index + 1,
            client,
            amount,
            percentage + '%'
        ]);
    });

    const wsClients = XLSX.utils.aoa_to_sheet(clientData);
    wsClients['!cols'] = [{wch: 8}, {wch: 40}, {wch: 25}, {wch: 15}];
    XLSX.utils.book_append_sheet(wb, wsClients, 'TOP CLIENTS');

    XLSX.writeFile(wb, 'D:\\RAPPORT MENSUEL HACC 2025\\RAPPORT ANNUEL 2025.xlsx');
    console.log('Rapport annuel Excel cree avec succes!');
}

function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

createAnnualReport();
