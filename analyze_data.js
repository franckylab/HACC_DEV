const fs = require('fs');

const wordDataPath = 'extracted_word_data.txt';
const excelDataPath = 'extracted_excel_data.txt';

function parseExcelData(content) {
    const months = [];
    const salesDetails = [];

    let currentMonth = null;
    let isSalesSheet = false;
    let lines = content.split('\n');

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Detect Month File
        if (line.includes('FILE: REPORTING ADMINISTRATIF-FINANCIER-VENTES')) {
            const match = line.match(/VENTES ([A-Z]+) 2025/);
            if (match) {
                currentMonth = {
                    month: match[1],
                    sales: 0,
                    expenses: 0,
                    net: 0,
                    bankDeposits: 0
                };
                months.push(currentMonth);
            }
        }

        // Detect Sheet
        if (line.includes('--- Sheet:')) {
            isSalesSheet = line.includes('VENTES');
        }

        if (line.includes('|')) {
            const parts = line.split('|').map(s => s.trim());

            // Parse Summary Data (Synthesis)
            if (currentMonth && !isSalesSheet) {
                if (parts[0].includes('Total Ventes')) currentMonth.sales = parseNumber(parts[1]);
                if (parts[0].match(/Total D.*penses/)) currentMonth.expenses = parseNumber(parts[1]);
                if (parts[0].includes('Solde Net')) currentMonth.net = parseNumber(parts[1]);
                if (parts[0].includes('Total Versements Bancaires')) currentMonth.bankDeposits = parseNumber(parts[1]);
            }

            // Parse Detailed Sales Data
            // Header: Date | Réf | Client | Produit | Quantité | Prix unitaire | Montant total
            // Index: 0 | 1 | 2 | 3 | 4 | 5 | 6
            if (isSalesSheet && parts.length > 5) {
                const client = parts[2];
                const product = parts[3];
                const qty = parseNumber(parts[4]);
                const amount = parseNumber(parts[6]);

                // Basic validation to ensure it's a data row and not a header or empty
                if (client && product && amount > 0 && client !== 'Client') {
                    salesDetails.push({
                        month: currentMonth ? currentMonth.month : 'UNKNOWN',
                        client: client,
                        product: product,
                        quantity: qty,
                        amount: amount
                    });
                }
            }
        }
    }
    return { months, salesDetails };
}

function parseWordData(content) {
    const reports = [];
    const lines = content.split('\n');
    let currentReport = null;

    for (const line of lines) {
        if (line.includes('=== FILE:')) {
            if (currentReport) reports.push(currentReport);
            currentReport = {
                filename: line.replace('=== FILE: ', '').replace(' ===', ''),
                highlights: []
            };
        } else if (currentReport) {
            const lower = line.toLowerCase();
            if (lower.includes('conclusion') || lower.includes('difficult') || lower.includes('perspective') || lower.includes('objectif')) {
                if (line.length < 300 && line.length > 10) {
                    currentReport.highlights.push(line.trim());
                }
            }
        }
    }
    if (currentReport) reports.push(currentReport);
    return reports;
}

function parseNumber(str) {
    if (!str) return 0;
    return parseFloat(str.replace(/[^0-9.-]+/g, '')) || 0;
}

try {
    const wordContent = fs.readFileSync(wordDataPath, 'utf8');
    const excelContent = fs.readFileSync(excelDataPath, 'utf8');

    const { months, salesDetails } = parseExcelData(excelContent);
    const activityReports = parseWordData(wordContent);

    // Aggregations
    const productStats = {};
    const clientStats = {};
    // Simplify product names to aggregate similar items (e.g., removing packaging details if feasible, or keeping as is)
    // For now, keep as is but trim

    salesDetails.forEach(s => {
        // Product Stats
        const pName = s.product;
        if (!productStats[pName]) productStats[pName] = { amount: 0, quantity: 0 };
        productStats[pName].amount += s.amount;
        productStats[pName].quantity += s.quantity;

        // Client Stats
        if (!clientStats[s.client]) clientStats[s.client] = 0;
        clientStats[s.client] += s.amount;
    });

    const summary = {
        financials: months,
        sales_details: {
            total_transactions: salesDetails.length,
            top_products: Object.entries(productStats)
                .map(([name, stats]) => ({ name, value: stats.amount, quantity: stats.quantity }))
                .sort((a, b) => b.value - a.value)
                .slice(0, 10),
            top_clients: Object.entries(clientStats)
                .map(([name, value]) => ({ name, value }))
                .sort((a, b) => b.value - a.value)
                .slice(0, 10)
        },
        activity_reports_covered: activityReports,
        total_yearly_stats: {
            sales: months.reduce((acc, m) => acc + m.sales, 0),
            expenses: months.reduce((acc, m) => acc + m.expenses, 0),
            net: months.reduce((acc, m) => acc + m.net, 0)
        }
    };

    console.log(JSON.stringify(summary, null, 2));

} catch (err) {
    console.error(err);
}
