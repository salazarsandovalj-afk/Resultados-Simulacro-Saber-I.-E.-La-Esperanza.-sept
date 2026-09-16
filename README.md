<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Institución Educativa La Esperanza — Dashboard Saber</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1e3a8a;
            --secondary: #2563eb;
            --accent: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --bg-gray: #f8fafc;
            --card-bg: #ffffff;
            --text-main: #1e293b;
            --text-muted: #64748b;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Inter', sans-serif;
        }

        body {
            background-color: var(--bg-gray);
            color: var(--text-main);
            padding: 24px;
        }

        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 24px;
            background: linear-gradient(135deg, #1e3a8a, #2563eb);
            color: white;
            padding: 22px 30px;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        }

        header h1 { font-size: 1.6rem; font-weight: 700; }
        header p { font-size: 0.95rem; opacity: 0.9; margin-top: 4px; }

        .controls {
            display: flex;
            gap: 12px;
            margin-bottom: 24px;
        }

        select, button {
            padding: 10px 16px;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            background: white;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
        }

        button.btn-primary {
            background: var(--accent);
            color: white;
            border: none;
        }

        button.btn-primary:hover {
            background: #059669;
        }

        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.04);
            border: 1px solid #e2e8f0;
        }

        .metric-title { font-size: 0.85rem; color: var(--text-muted); font-weight: 600; text-transform: uppercase; }
        .metric-value { font-size: 1.8rem; font-weight: 700; color: var(--primary); margin-top: 6px; }
        .metric-sub { font-size: 0.8rem; color: var(--text-muted); margin-top: 4px; }

        .dafo-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }

        .dafo-card {
            padding: 16px;
            border-radius: 10px;
            border-left: 5px solid;
            background: white;
        }

        .dafo-f { border-color: var(--accent); background: #f0fdf4; }
        .dafo-d { border-color: var(--danger); background: #fef2f2; }
        .dafo-o { border-color: var(--secondary); background: #eff6ff; }
        .dafo-a { border-color: var(--warning); background: #fffbeb; }

        .dafo-card h3 { font-size: 1rem; margin-bottom: 8px; }
        .dafo-card ul { font-size: 0.85rem; padding-left: 18px; line-height: 1.5; }

        .charts-grid {
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 24px;
            margin-bottom: 24px;
        }

        @media (max-width: 900px) {
            .charts-grid { grid-template-columns: 1fr; }
        }

        .chart-card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 20px;
            border: 1px solid #e2e8f0;
        }

        .chart-card h2 { font-size: 1.1rem; margin-bottom: 16px; color: var(--primary); }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.85rem;
        }

        th, td {
            padding: 12px 14px;
            text-align: left;
            border-bottom: 1px solid #e2e8f0;
        }

        th { background-color: #f1f5f9; color: var(--text-muted); font-weight: 600; }

        .badge {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 700;
            display: inline-block;
        }

        .badge-alto { background: #dcfce7; color: #15803d; }
        .badge-medio { background: #fef3c7; color: #b45309; }
        .badge-bajo { background: #fee2e2; color: #b91c1c; }
    </style>
</head>
<body>

    <header>
        <div>
            <h1>Institución Educativa La Esperanza</h1>
            <p>Dashboard de Resultados Saber — Simulacro #2 (Grados 09° y 10°)</p>
        </div>
        <button class="btn-primary" onclick="exportExcel()">Exportar Excel</button>
    </header>

    <div class="controls">
        <select id="filterGrado" onchange="renderDashboard()">
            <option value="ALL">Todos los Grados (09° y 10°)</option>
            <option value="10°">Grado 10°</option>
            <option value="09°">Grado 09°</option>
        </select>
    </div>

    <div class="metrics-grid">
        <div class="card">
            <div class="metric-title">Promedio Global</div>
            <div class="metric-value" id="valPromedio">179.9 pts</div>
            <div class="metric-sub" id="valDE">Desviación Estándar: 35.7 pts</div>
        </div>
        <div class="card">
            <div class="metric-title">Mejor Área</div>
            <div class="metric-value" style="color: #059669;" id="valMejorArea">C. Naturales</div>
            <div class="metric-sub" id="valMejorScore">41.1 pts prom.</div>
        </div>
        <div class="card">
            <div class="metric-title">Área a Reforzar</div>
            <div class="metric-value" style="color: var(--danger);" id="valPeorArea">Sociales</div>
            <div class="metric-sub" id="valPeorScore">25.7 pts prom.</div>
        </div>
        <div class="card">
            <div class="metric-title">Total Estudiantes</div>
            <div class="metric-value" id="valTotalEst">23</div>
            <div class="metric-sub">13 de 9° | 10 de 10°</div>
        </div>
    </div>

    <!-- Sección DAFO -->
    <h2 style="margin-bottom: 12px; font-size: 1.2rem; color: var(--primary);">Matriz de Diagnóstico DAFO</h2>
    <div class="dafo-grid">
        <div class="dafo-card dafo-f">
            <h3 style="color: #15803d;">Fortalezas</h3>
            <ul>
                <li>Buen desempeño en Naturales (41.1 pts) y Lectura (41.0 pts).</li>
                <li>Estudiante top pertenece a 9° (Samuel Pinedo - 244.6 pts).</li>
                <li>40% de Grado 10° se ubica en Nivel Medio/Alto.</li>
            </ul>
        </div>
        <div class="dafo-card dafo-d">
            <h3 style="color: #b91c1c;">Debilidades</h3>
            <ul>
                <li>Sociales y Ciudadanas muy rezagada (25.7 pts prom.).</li>
                <li>Inglés en 9° presenta bajo desempeño (28.1 pts).</li>
                <li>Mayoría de la población por debajo de los 200 pts.</li>
            </ul>
        </div>
        <div class="dafo-card dafo-o">
            <h3 style="color: #1d4ed8;">Oportunidades</h3>
            <ul>
                <li>3 estudiantes de 9° y 4 de 10° listos para superar los 250 pts.</li>
                <li>Estrategias de control de fatiga para la Sesión 2.</li>
                <li>Refuerzo cruzado con monitores destacados.</li>
            </ul>
        </div>
        <div class="dafo-card dafo-a">
            <h3 style="color: #b45309;">Amenazas</h3>
            <ul>
                <li>Estudiantes en Nivel Bajo (< 200 pts) requieren plan urgente.</li>
                <li>Riesgo de brecha acumulada al pasar de 9° a 10°.</li>
            </ul>
        </div>
    </div>

    <div class="charts-grid">
        <div class="chart-card">
            <h2>Promedios por Área</h2>
            <canvas id="barChart" height="140"></canvas>
        </div>
        <div class="chart-card">
            <h2>Distribución por Niveles de Desempeño</h2>
            <canvas id="pieChart"></canvas>
        </div>
    </div>

    <div class="chart-card">
        <h2>Ranking y Clasificación de Estudiantes</h2>
        <table id="studentsTable">
            <thead>
                <tr>
                    <th>Pos.</th>
                    <th>Grado</th>
                    <th>Nombre y Apellidos</th>
                    <th>Lectura</th>
                    <th>Matemáticas</th>
                    <th>Sociales</th>
                    <th>Naturales</th>
                    <th>Inglés</th>
                    <th>Puntaje Global</th>
                    <th>Nivel</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>

    <script>
        const rawData = [
            {grado: '10°', id: 1, nombre: 'CASTRILLO MARRIAGA SHARIK JHOANA', lec: 53.7, mat: 46, soc: 36, nat: 57.85, ing: 40.0, global: 238.71},
            {grado: '10°', id: 2, nombre: 'CASTRILLO MARRIAGA SHANI PATRICIA', lec: 63.4, mat: 42, soc: 30, nat: 45.60, ing: 41.8, global: 224.92},
            {grado: '10°', id: 3, nombre: 'FLORES GUERRERO MARIA FERNANDA', lec: 73.2, mat: 26, soc: 40, nat: 38.45, ing: 50.9, global: 224.56},
            {grado: '10°', id: 4, nombre: 'TORRES RODRIGUEZ JHON ANDERSON', lec: 46.3, mat: 52, soc: 26, nat: 61.30, ing: 25.5, global: 223.96},
            {grado: '10°', id: 5, nombre: 'ORTIZ TALAIGUA MARIA ALEJANDRA', lec: 41.5, mat: 60, soc: 24, nat: 28.10, ing: 54.5, global: 198.19},
            {grado: '10°', id: 6, nombre: 'VEGA CONTRERAS JAIDER DAVID', lec: 39.0, mat: 46, soc: 32, nat: 40.25, ing: 30.9, global: 193.33},
            {grado: '10°', id: 7, nombre: 'GONZALEZ SOLIS JUAN CAMILO', lec: 26.8, mat: 46, soc: 22, nat: 43.90, ing: 47.3, global: 178.23},
            {grado: '10°', id: 8, nombre: 'GASPAR SIERRA EDUARDO EMILIO', lec: 39.0, mat: 24, soc: 24, nat: 52.50, ing: 29.1, global: 172.15},
            {grado: '10°', id: 9, nombre: 'PAYARES HERNANDEZ LEISI CAROLINA', lec: 39.0, mat: 28, soc: 16, nat: 29.90, ing: 49.1, global: 149.15},
            {grado: '10°', id: 10, nombre: 'GUERRERO VELASQUEZ ANDRES', lec: 31.7, mat: 22, soc: 18, nat: 43.90, ing: 25.5, global: 143.19},
            {grado: '09°', id: 1, nombre: 'PINEDO BARRIOS SAMUEL', lec: 61.0, mat: 54, soc: 22, nat: 64.65, ing: 30.9, global: 244.56},
            {grado: '09°', id: 2, nombre: 'BENITEZ URANGO DINA LUZ', lec: 48.8, mat: 48, soc: 38, nat: 36.55, ing: 25.5, global: 207.52},
            {grado: '09°', id: 3, nombre: 'SANCHEZ FUENTES VALENTINA', lec: 58.5, mat: 36, soc: 20, nat: 50.80, ing: 30.9, global: 202.62},
            {grado: '09°', id: 4, nombre: 'ARIAS MONTALVO MARIA JOSE', lec: 34.1, mat: 50, soc: 22, nat: 36.65, ing: 36.4, global: 178.71},
            {grado: '09°', id: 5, nombre: 'PEREZ VERGARA MARIA ANGEL', lec: 39.0, mat: 30, soc: 34, nat: 36.95, ing: 40.0, global: 176.87},
            {grado: '09°', id: 6, nombre: 'VARGAS APARICIO ORLANDO MIGUEL', lec: 41.5, mat: 34, soc: 26, nat: 38.45, ing: 20.0, global: 169.17},
            {grado: '09°', id: 7, nombre: 'FABRA MACEA ISAAC ELIAS', lec: 39.0, mat: 34, soc: 26, nat: 29.65, ing: 30.9, global: 160.33},
            {grado: '09°', id: 8, nombre: 'HENAO SOLAR HORIANA ISABEL', lec: 29.3, mat: 38, soc: 18, nat: 40.25, ing: 30.9, global: 156.75},
            {grado: '09°', id: 9, nombre: 'PINEDA OVIEDO NEIDIS YAMILES', lec: 43.9, mat: 30, soc: 16, nat: 33.30, ing: 34.5, global: 155.42},
            {grado: '09°', id: 10, nombre: 'MANCHEGO CESPEDES DARWIN ANDRES', lec: 31.7, mat: 22, soc: 22, nat: 45.35, ing: 25.5, global: 149.48},
            {grado: '09°', id: 11, nombre: 'ACOSTA ZUÑIGA LUIS ENRIQUE', lec: 24.4, mat: 30, soc: 30, nat: 27.85, ing: 27.3, global: 140.02},
            {grado: '09°', id: 12, nombre: 'TRUJILLO BERROCAL SAMUEL', lec: 31.7, mat: 28, soc: 24, nat: 27.95, ing: 14.5, global: 134.40},
            {grado: '09°', id: 13, nombre: 'SUAREZ TALAIGUA MIGUEL ANDRES', lec: 7.3, mat: 28, soc: 24, nat: 35.00, ing: 18.2, global: 115.81}
        ];

        let barChartInstance = null;
        let pieChartInstance = null;

        function renderDashboard() {
            const filter = document.getElementById('filterGrado').value;
            const filteredData = filter === 'ALL' ? rawData : rawData.filter(d => d.grado === filter);
            
            // Ordenar por puntaje global real en Excel (descendente)
            filteredData.sort((a,b) => b.global - a.global);

            const total = filteredData.length;
            const sumGlobal = filteredData.reduce((acc, d) => acc + d.global, 0);
            const avgGlobal = (sumGlobal / total).toFixed(2);
            
            // Desviación Estándar en décimas
            const stdDev = Math.sqrt(filteredData.reduce((acc, d) => acc + Math.pow(d.global - parseFloat(avgGlobal), 2), 0) / total).toFixed(2);

            document.getElementById('valPromedio').innerText = avgGlobal + ' pts';
            document.getElementById('valDE').innerText = 'Desviación Estándar: ' + stdDev + ' pts';
            document.getElementById('valTotalEst').innerText = total;

            // Promedios por Área mantenidos con sus décimas exactas del Excel
            const areaAvg = {
                'Lectura Crítica': (filteredData.reduce((a, b) => a + b.lec, 0) / total).toFixed(1),
                'Matemáticas': (filteredData.reduce((a, b) => a + b.mat, 0) / total).toFixed(1),
                'Sociales': (filteredData.reduce((a, b) => a + b.soc, 0) / total).toFixed(1),
                'Naturales': (filteredData.reduce((a, b) => a + b.nat, 0) / total).toFixed(1),
                'Inglés': (filteredData.reduce((a, b) => a + b.ing, 0) / total).toFixed(1)
            };

            const sortedAreas = Object.entries(areaAvg).sort((a,b) => b[1] - a[1]);
            document.getElementById('valMejorArea').innerText = sortedAreas[0][0];
            document.getElementById('valMejorScore').innerText = sortedAreas[0][1] + ' pts prom.';
            document.getElementById('valPeorArea').innerText = sortedAreas[sortedAreas.length - 1][0];
            document.getElementById('valPeorScore').innerText = sortedAreas[sortedAreas.length - 1][1] + ' pts prom.';

            // Niveles según los rangos estipulados sobre el valor con décimas:
            // Bajo: < 200 | Medio: 200 - 250 | Alto: > 250
            let alto = 0, medio = 0, bajo = 0;
            filteredData.forEach(d => {
                const g = d.global;
                if (g > 250) alto++;
                else if (g >= 200) medio++;
                else bajo++;
            });

            // Actualizar Tabla de Estudiantes manteniendo valores exactos
            const tbody = document.querySelector('#studentsTable tbody');
            tbody.innerHTML = '';
            filteredData.forEach((st, idx) => {
                const globalVal = st.global.toFixed(2);
                let badgeClass = 'badge-bajo';
                let levelText = 'Bajo (< 200)';

                if (st.global > 250) {
                    badgeClass = 'badge-alto';
                    levelText = 'Alto (> 250)';
                } else if (st.global >= 200) {
                    badgeClass = 'badge-medio';
                    levelText = 'Medio (200 - 250)';
                }

                tbody.innerHTML += `<tr>
                    <td><b>${idx + 1}</b></td>
                    <td>${st.grado}</td>
                    <td>${st.nombre}</td>
                    <td>${st.lec}</td>
                    <td>${st.mat}</td>
                    <td>${st.soc}</td>
                    <td>${st.nat}</td>
                    <td>${st.ing}</td>
                    <td><b>${globalVal}</b></td>
                    <td><span class="badge ${badgeClass}">${levelText}</span></td>
                </tr>`;
            });

            // Actualizar Gráficos
            renderBarChart(Object.keys(areaAvg), Object.values(areaAvg));
            renderPieChart([alto, medio, bajo]);
        }

        function renderBarChart(labels, values) {
            const ctx = document.getElementById('barChart').getContext('2d');
            if (barChartInstance) barChartInstance.destroy();

            barChartInstance = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'Promedio de Puntaje',
                        data: values,
                        backgroundColor: [
                            'rgba(54, 162, 235, 0.85)',  // Azul - Lectura Crítica
                            'rgba(255, 159, 64, 0.85)',   // Naranja - Matemáticas
                            'rgba(255, 99, 132, 0.85)',   // Rojo/Rosa - Sociales
                            'rgba(75, 192, 192, 0.85)',   // Verde - Naturales
                            'rgba(153, 102, 255, 0.85)'   // Morado - Inglés
                        ],
                        borderColor: [
                            '#1d4ed8', '#c2410c', '#be123c', '#047857', '#6d28d9'
                        ],
                        borderWidth: 1.5,
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    scales: { y: { beginAtZero: true, max: 100 } },
                    plugins: { legend: { display: false } }
                }
            });
        }

        function renderPieChart(dataCounts) {
            const ctx = document.getElementById('pieChart').getContext('2d');
            if (pieChartInstance) pieChartInstance.destroy();

            pieChartInstance = new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: ['Alto (> 250)', 'Medio (200 - 250)', 'Bajo (< 200)'],
                    datasets: [{
                        data: dataCounts,
                        backgroundColor: ['#10b981', '#f59e0b', '#ef4444']
                    }]
                },
                options: { responsive: true }
            });
        }

        function exportExcel() {
            const filter = document.getElementById('filterGrado').value;
            const filteredData = filter === 'ALL' ? rawData : rawData.filter(d => d.grado === filter);
            
            const processedData = filteredData.map(d => ({
                'Grado': d.grado,
                'Nombre y Apellidos': d.nombre,
                'Lectura Crítica': d.lec,
                'Matemáticas': d.mat,
                'Sociales y Ciudadanas': d.soc,
                'Ciencias Naturales': d.nat,
                'Inglés': d.ing,
                'Puntaje Global': Number(d.global.toFixed(2)),
                'Nivel': d.global > 250 ? 'Alto' : (d.global >= 200 ? 'Medio' : 'Bajo')
            }));

            const ws = XLSX.utils.json_to_sheet(processedData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Resultados");
            XLSX.writeFile(wb, "Resultados_IE_La_Esperanza.xlsx");
        }

        // Render inicial
        renderDashboard();
    </script>
</body>
</html>
