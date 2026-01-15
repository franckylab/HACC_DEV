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

function generateAnnualReport() {
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

    console.log('═'.repeat(78));
    console.log("RAPPORT ANNUEL D'ACTIVITÉ 2025 - HACC YAOUNDÉ");
    console.log('═'.repeat(78));
    console.log('');

    console.log('═'.repeat(78));
    console.log('I - SYNTHESE FINANCIERE ANNUELLE');
    console.log('═'.repeat(78));
    console.log('');

    console.log('INDICATEURS FINANCIERS ANNUELS');
    console.log('-'.repeat(78));
    console.log(`Total des Ventes       : ${formatNumber(annualTotals.totalVentes)} FCFA`);
    console.log(`Total des Depenses    : ${formatNumber(annualTotals.totalDepenses)} FCFA`);
    console.log(`Solde Net Annuel       : ${formatNumber(annualTotals.soldeNet)} FCFA`);
    console.log(`Versements Bancaires  : ${formatNumber(annualTotals.totalVersements)} FCFA`);
    console.log('');

    console.log('═'.repeat(78));
    console.log('II - ANALYSE MENSUELLE');
    console.log('═'.repeat(78));
    console.log('');

    console.log('Mois            | Total Ventes        | Total Depenses       | Solde Net           ');
    console.log('-'.repeat(78));

    monthlyData.forEach(month => {
        console.log(`${month.month.padEnd(15)} | ${formatNumber(month.totalVentes).padEnd(20)} | ${formatNumber(month.totalDepenses).padEnd(20)} | ${formatNumber(month.soldeNet)}`);
    });

    console.log('-'.repeat(78));
    console.log(`TOTAL ANNUEL     | ${formatNumber(annualTotals.totalVentes).padEnd(20)} | ${formatNumber(annualTotals.totalDepenses).padEnd(20)} | ${formatNumber(annualTotals.soldeNet)}`);
    console.log('');

    console.log('═'.repeat(78));
    console.log('III - TOP 10 PRODUITS PAR CHIFFRE D\'AFFAIRES');
    console.log('═'.repeat(78));
    console.log('');

    const sortedProducts = Object.entries(annualProductSales)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    console.log('#  | Produit                                        | Chiffre d\'A.        | %       ');
    console.log('-'.repeat(78));

    sortedProducts.forEach((item, index) => {
        const [product, amount] = item;
        const percentage = ((amount / annualTotals.totalVentes) * 100).toFixed(2);
        const productName = product.length > 46 ? product.substring(0, 46) : product;
        console.log(`${(index + 1).toString().padStart(2)} | ${productName.padEnd(46)} | ${formatNumber(amount).padEnd(20)} | ${percentage}%`);
    });

    console.log('');

    console.log('═'.repeat(78));
    console.log('IV - TOP 15 CLIENTS PAR CHIFFRE D\'AFFAIRES');
    console.log('═'.repeat(78));
    console.log('');

    const sortedClients = Object.entries(annualClientSales)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);

    console.log('#  | Client                                 | Chiffre d\'A.        | %       ');
    console.log('-'.repeat(78));

    sortedClients.forEach((item, index) => {
        const [client, amount] = item;
        const percentage = ((amount / annualTotals.totalVentes) * 100).toFixed(2);
        const clientName = client.length > 39 ? client.substring(0, 39) : client;
        console.log(`${(index + 1).toString().padStart(2)} | ${clientName.padEnd(39)} | ${formatNumber(amount).padEnd(20)} | ${percentage}%`);
    });

    console.log('');

    console.log('═'.repeat(78));
    console.log('V - INDICATEURS DE PERFORMANCE');
    console.log('═'.repeat(78));
    console.log('');

    const avgMonthlySales = annualTotals.totalVentes / monthlyData.length;
    const ratioVersements = ((annualTotals.totalVersements / annualTotals.totalVentes) * 100).toFixed(2);
    const ratioDepenses = ((annualTotals.totalDepenses / annualTotals.totalVentes) * 100).toFixed(2);

    console.log('INDICATEURS CLES DE PERFORMANCE');
    console.log('-'.repeat(78));
    console.log(`Moyenne Mensuelle Ventes : ${formatNumber(avgMonthlySales)} FCFA`);
    console.log(`Ratio Versements/Ventes  : ${ratioVersements}%`);
    console.log(`Ratio Depenses/Ventes   : ${ratioDepenses}%`);
    console.log(`Marge Beneficiaire      : ${((1 - (annualTotals.totalDepenses / annualTotals.totalVentes)) * 100).toFixed(2)}%`);
    console.log('');

    console.log('═'.repeat(78));
    console.log('VI - RECOMMANDATIONS');
    console.log('═'.repeat(78));
    console.log('');

    console.log('1. GESTION COMMERCIALE:');
    console.log('   - Maintenir la relation avec les principaux clients identifies');
    console.log('   - Developper des actions de fidelisation pour les gros clients');
    console.log('   - Explorer de nouveaux segments de marche pour diversifier le portefeuille');
    console.log('');

    console.log('2. OPTIMISATION DES DEPENSES:');
    console.log(`   - Le ratio de depenses par rapport aux ventes est de ${ratioDepenses}%`);
    console.log('   - Continuer le controle strict des depenses operationnelles');
    console.log('   - Identifier les opportunites d\'economies sur les couts fixes');
    console.log('');

    console.log('3. GESTION DE TRESORERIE:');
    console.log(`   - ${ratioVersements}% des ventes sont encaissees via versements bancaires`);
    console.log('   - Maintenir une politique de recouvrement efficace');
    console.log('   - Optimiser les delais de depot bancaire');
    console.log('');

    console.log('4. PERSPECTIVES 2026:');
    console.log('   - Poursuivre la croissance des ventes mensuelles');
    console.log('   - Renforcer la presence sur les marches actuels');
    console.log('   - Explorer de nouvelles zones geographiques d\'expansion');
    console.log('');

    console.log('═'.repeat(78));
    console.log('FIN DU RAPPORT ANNUEL');
    console.log('═'.repeat(78));
    console.log('');
}

function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

generateAnnualReport();
