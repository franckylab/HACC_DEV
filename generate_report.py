
import json
import os

JSON_PATH = r"d:\HACC_DEV\sales_data.json"
OUTPUT_PATH = r"d:\HACC_DEV\Rapport_Annuel_2025.html"

def generate_html():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    json_data_string = json.dumps(data)

    html_content = f"""
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rapport Annuel 2025</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/js/all.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: 'Inter', sans-serif; -webkit-print-color-adjust: exact; }}
        @media print {{
            .no-print {{ display: none; }}
            .page-break {{ page-break-before: always; }}
            body {{ background: white; }}
        }}
        .glass {{ background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(10px); }}
    </style>
</head>
<body class="bg-slate-50 text-slate-800">

    <!-- Header -->
    <div class="bg-slate-900 text-white p-8 shadow-lg print:bg-slate-900 print:text-white">
        <div class="max-w-7xl mx-auto flex justify-between items-center">
            <div>
                <h1 class="text-3xl font-bold tracking-tight">RAPPORT ANNUEL <span class="text-blue-400">2025</span></h1>
                <p class="text-slate-400 mt-1">Analyse Commerciale et Financière</p>
            </div>
            <div class="text-right hidden sm:block">
                <p class="text-sm opacity-70">Généré le {os.popen('date /t').read().strip()}</p>
                <p class="font-semibold">HACC DEV</p>
            </div>
        </div>
    </div>

    <div class="max-w-7xl mx-auto p-8 space-y-8">
        
        <!-- Introduction -->
        <section class="bg-white rounded-xl shadow-sm p-6 border-l-4 border-blue-500">
            <h2 class="text-xl font-bold mb-4 flex items-center"><i class="fas fa-info-circle mr-3 text-blue-500"></i> Introduction</h2>
            <p class="text-slate-600 leading-relaxed">
                Ce rapport présente une analyse détaillée de l'activité commerciale pour l'année 2025. 
                Il consolide les données issues des rapports mensuels (ventes, quantités, portefeuille client) 
                afin d'offrir une vision globale de la performance de l'entreprise. 
                Les analyses suivantes mettent en lumière les tendances de revenus, la dynamique des produits et la fidélité des clients.
            </p>
        </section>

        <!-- KPI Cards -->
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <!-- CA Total -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <div class="flex justify-between items-start">
                    <div>
                        <p class="text-sm text-slate-500 font-medium">Chiffre d'Affaires Total</p>
                        <h3 class="text-3xl font-bold text-slate-900 mt-2" id="kpi-revenue">---</h3>
                    </div>
                    <div class="p-3 bg-blue-50 rounded-lg text-blue-600">
                        <i class="fas fa-wallet text-xl"></i>
                    </div>
                </div>
                <p class="text-xs text-green-600 mt-4 flex items-center">
                    <i class="fas fa-arrow-up mr-1"></i> Performance Annuelle
                </p>
            </div>

            <!-- Volume Ventes -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <div class="flex justify-between items-start">
                    <div>
                        <p class="text-sm text-slate-500 font-medium">Volume Total (Unités)</p>
                        <h3 class="text-3xl font-bold text-slate-900 mt-2" id="kpi-volume">---</h3>
                    </div>
                    <div class="p-3 bg-purple-50 rounded-lg text-purple-600">
                        <i class="fas fa-cubes text-xl"></i>
                    </div>
                </div>
                <p class="text-xs text-slate-400 mt-4">Cumul annuel</p>
            </div>

            <!-- Meilleur Mois -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <div class="flex justify-between items-start">
                    <div>
                        <p class="text-sm text-slate-500 font-medium">Meilleur Mois</p>
                        <h3 class="text-2xl font-bold text-slate-900 mt-2" id="kpi-best-month">---</h3>
                        <p class="text-sm text-slate-500" id="kpi-best-month-val">---</p>
                    </div>
                    <div class="p-3 bg-green-50 rounded-lg text-green-600">
                        <i class="fas fa-calendar-check text-xl"></i>
                    </div>
                </div>
            </div>

            <!-- Top Client -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <div class="flex justify-between items-start">
                    <div class="overflow-hidden">
                        <p class="text-sm text-slate-500 font-medium">Top Client</p>
                        <h3 class="text-xl font-bold text-slate-900 mt-2 truncate" id="kpi-top-client">---</h3>
                        <p class="text-sm text-slate-500" id="kpi-top-client-val">---</p>
                    </div>
                    <div class="p-3 bg-orange-50 rounded-lg text-orange-600">
                        <i class="fas fa-crown text-xl"></i>
                    </div>
                </div>
            </div>
        </div>

        <!-- Charts Row 1 -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <!-- Main Chart: Evolution -->
            <div class="lg:col-span-2 bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <h3 class="text-lg font-bold mb-6">Évolution Mensuelle du Chiffre d'Affaires</h3>
                <canvas id="chart-revenue"></canvas>
                <p class="text-sm text-slate-500 mt-4 italic bg-slate-50 p-3 rounded">
                    <i class="fas fa-comment-alt mr-2"></i>
                    <span id="comment-revenue">Analyse en cours...</span>
                </p>
            </div>

            <!-- Pie Chart: Products -->
            <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100">
                <h3 class="text-lg font-bold mb-6">Répartition par Produit</h3>
                <div class="relative h-64">
                    <canvas id="chart-products"></canvas>
                </div>
            </div>
        </div>

        <!-- Charts Row 2 -->
        <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-100 page-break">
            <h3 class="text-lg font-bold mb-6">Top 10 Clients (Chiffre d'Affaires)</h3>
            <canvas id="chart-clients" height="100"></canvas>
        </div>

        <!-- Data Table -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden">
            <div class="p-6 border-b border-slate-100 flex justify-between items-center">
                <h3 class="text-lg font-bold">Détail Mensuel (Synthèse)</h3>
                <button onclick="window.print()" class="bg-slate-800 text-white px-4 py-2 rounded-lg text-sm hover:bg-slate-700 transition no-print">
                    <i class="fas fa-print mr-2"></i> Imprimer
                </button>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-sm text-left">
                    <thead class="bg-slate-50 text-slate-600 uppercase">
                        <tr>
                            <th class="px-6 py-3">Mois</th>
                            <th class="px-6 py-3 text-right">CA Total</th>
                            <th class="px-6 py-3 text-right">Quantité</th>
                            <th class="px-6 py-3">Top Client du Mois</th>
                        </tr>
                    </thead>
                    <tbody id="table-body" class="divide-y divide-slate-100">
                        <!-- JS generated -->
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Conclusion -->
        <section class="bg-slate-900 text-slate-300 rounded-xl shadow-sm p-8 mt-8">
            <h2 class="text-2xl font-bold text-white mb-4">Conclusion & Perspectives</h2>
            <p class="leading-relaxed mb-4">
                L'année 2025 se caractérise par une dynamique commerciale variée. 
                L'analyse des données montre une forte concentration du chiffre d'affaires sur les produits phares et une fidélité notable de certains grands comptes.
            </p>
            <p class="leading-relaxed">
                Pour l'exercice à venir, il est recommandé de focaliser les efforts sur la diversification du portefeuille client et le maintien de la croissance observée lors des mois les plus performants.
            </p>
        </section>

    </div>

    <script>
        const rawData = {json_data_string};

        // --- Data Processing ---
        const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        const monthNamesFr = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

        // Aggregate by Month
        const monthlyData = {{}};
        monthNames.forEach(m => monthlyData[m] = {{ revenue: 0, qty: 0, clients: {{}} }});

        let totalRevenue = 0;
        let totalQty = 0;
        const clientTotals = {{}};
        const productTotals = {{}};

        rawData.forEach(row => {{
            const m = row.month; // Assuming full English month name from Python
            if (monthlyData[m]) {{
                monthlyData[m].revenue += row.revenue;
                monthlyData[m].qty += row.quantity;
                
                // Track client per month for "Top Client of Month"
                if (!monthlyData[m].clients[row.client]) monthlyData[m].clients[row.client] = 0;
                monthlyData[m].clients[row.client] += row.revenue;
            }}

            totalRevenue += row.revenue;
            totalQty += row.quantity;

            // Global Client Totals
            if (!clientTotals[row.client]) clientTotals[row.client] = 0;
            clientTotals[row.client] += row.revenue;

            // Global Product Totals
            if (!productTotals[row.product]) productTotals[row.product] = 0;
            productTotals[row.product] += row.revenue;
        }});

        // Prepare Chart Data arrays aligned with monthNames
        const revenueArray = monthNames.map(m => monthlyData[m].revenue);
        const qtyArray = monthNames.map(m => monthlyData[m].qty);

        // Find Best Month
        let maxRev = -1;
        let bestMonth = "";
        revenueArray.forEach((rev, idx) => {{
            if (rev > maxRev) {{
                maxRev = rev;
                bestMonth = monthNamesFr[idx];
            }}
        }});

        // Find Top Client
        let topClientName = "";
        let topClientRev = -1;
        Object.entries(clientTotals).forEach(([name, rev]) => {{
            if (rev > topClientRev && name !== "Unknown") {{
                topClientRev = rev;
                topClientName = name;
            }}
        }});

        // --- Update UI ---
        const formatMoney = (amount) => {{
            return new Intl.NumberFormat('fr-FR', {{ style: 'currency', currency: 'XAF' }}).format(amount);
        }};

        document.getElementById('kpi-revenue').textContent = formatMoney(totalRevenue);
        document.getElementById('kpi-volume').textContent = new Intl.NumberFormat('fr-FR').format(totalQty);
        document.getElementById('kpi-best-month').textContent = bestMonth;
        document.getElementById('kpi-best-month-val').textContent = formatMoney(maxRev);
        document.getElementById('kpi-top-client').textContent = topClientName;
        document.getElementById('kpi-top-client-val').textContent = formatMoney(topClientRev);

        // --- Charts ---
        
        // 1. Revenue Evolution (Line)
        new Chart(document.getElementById('chart-revenue'), {{
            type: 'line',
            data: {{
                labels: monthNamesFr,
                datasets: [{{
                    label: "Chiffre d'Affaires (XAF)",
                    data: revenueArray,
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    fill: true,
                    tension: 0.4
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{ legend: {{ display: false }} }},
                scales: {{ y: {{ beginAtZero: true }} }}
            }}
        }});

        // Commentary Logic
        const avgRev = totalRevenue / 12; // Simple avg
        document.getElementById('comment-revenue').textContent = 
            `Le chiffre d'affaires cumulé atteint ${{formatMoney(totalRevenue)}}. ` +
            `Le mois record est ${{bestMonth}}. ` + 
            (revenueArray[11] > revenueArray[0] ? "La tendance globale est à la hausse vers la fin de l'année." : "L'année montre des fluctuations saisonnières importantes.");


        // 2. Products (Doughnut)
        // Sort products
        const sortedProducts = Object.entries(productTotals).sort((a,b) => b[1] - a[1]);
        const topProducts = sortedProducts.slice(0, 5);
        const otherProductsRev = sortedProducts.slice(5).reduce((acc, curr) => acc + curr[1], 0);
        
        const productLabels = topProducts.map(p => p[0]);
        const productData = topProducts.map(p => p[1]);
        if (otherProductsRev > 0) {{
            productLabels.push("Autres");
            productData.push(otherProductsRev);
        }}

        new Chart(document.getElementById('chart-products'), {{
            type: 'doughnut',
            data: {{
                labels: productLabels,
                datasets: [{{
                    data: productData,
                    backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#cbd5e1']
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{ legend: {{ position: 'right' }} }}
            }}
        }});

        // 3. Top Clients (Bar)
        const sortedClients = Object.entries(clientTotals).sort((a,b) => b[1] - a[1]).slice(0, 10);
        new Chart(document.getElementById('chart-clients'), {{
            type: 'bar',
            data: {{
                labels: sortedClients.map(c => c[0]),
                datasets: [{{
                    label: "Chiffre d'Affaires",
                    data: sortedClients.map(c => c[1]),
                    backgroundColor: '#1e293b',
                    borderRadius: 4
                }}]
            }},
            options: {{
                responsive: true,
                indexAxis: 'y',
                plugins: {{ legend: {{ display: false }} }}
            }}
        }});

        // --- Populate Table ---
        const tbody = document.getElementById('table-body');
        monthNames.forEach((m, idx) => {{
            if (monthlyData[m].revenue === 0) return; // Skip empty months if any

            // Find top client for this month
            let bestClientMonth = "-";
            let bestClientMonthRev = -1;
            Object.entries(monthlyData[m].clients).forEach(([c, r]) => {{
                if (r > bestClientMonthRev) {{ bestClientMonthRev = r; bestClientMonth = c; }}
            }});

            const tr = document.createElement('tr');
            tr.className = "hover:bg-slate-50";
            tr.innerHTML = `
                <td class="px-6 py-4 font-medium text-slate-900">${{monthNamesFr[idx]}}</td>
                <td class="px-6 py-4 text-right">${{formatMoney(monthlyData[m].revenue)}}</td>
                <td class="px-6 py-4 text-right">${{monthlyData[m].qty}}</td>
                <td class="px-6 py-4 text-slate-500 text-xs truncate max-w-xs" title="${{bestClientMonth}}">${{bestClientMonth}}</td>
            `;
            tbody.appendChild(tr);
        }});

    </script>
</body>
</html>
    """
    
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"Report generated at: {OUTPUT_PATH}")

if __name__ == "__main__":
    generate_html()
