document.addEventListener('DOMContentLoaded', () => {
    // 1. Verificar sesión
    const currentUser = AUTH.checkSession();
    if (!currentUser) return;

    // UI Elements
    const userNameEl = document.getElementById('userName');
    const greetingEl = document.getElementById('greeting');
    const logoutBtn = document.getElementById('logoutBtn');
    const navItems = document.querySelectorAll('.nav-item');
    const sections = document.querySelectorAll('.content-section');
    
    // Sidebar Mobile
    const sidebar = document.getElementById('sidebar');
    const mobileToggle = document.getElementById('mobileToggle');
    const mobileClose = document.getElementById('mobileClose');

    // --- Advanced Excel State ---
    let currentWorkbook = null;
    let currentSheetData = [];
    let currentFileName = "";

    // Use window.excelContexto so it is globally accessible everywhere
    window.excelContexto = JSON.parse(localStorage.getItem(`context_${currentUser.email}`) || 'null');
    
    // UI Elements (Advanced)
    const dropZoneOverlay = document.getElementById('dropZoneOverlay');
    const excelStats = document.getElementById('excelStats');
    const sheetTabs = document.getElementById('sheetTabs');
    const tableSearch = document.getElementById('tableSearch');
    const tableCounter = document.getElementById('tableCounter');
    const exportBtn = document.getElementById('exportBtn');
    const askAssistantBtn = document.getElementById('askAssistantBtn');
    const historyList = document.getElementById('historyList');
    const fullHistoryGrid = document.getElementById('fullHistoryGrid');
    const currentFileNameEl = document.getElementById('currentFileName');
    const currentFileMetaEl = document.getElementById('currentFileMeta');
    const contextBanner = document.getElementById('contextBanner');
    const contextFileName = document.getElementById('contextFileName');
    const clearContextBtn = document.getElementById('clearContextBtn');

    // Initialize UI
    userNameEl.textContent = currentUser.name;
    greetingEl.textContent = `Hola, ${currentUser.name.split(' ')[0]} 👋`;
    renderHistory();
    renderContextBanner();
    renderSuggestions();
    lucide.createIcons();

    // --- Navigation Logic ---
    navItems.forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const target = item.getAttribute('data-target');
            navItems.forEach(nav => nav.classList.remove('active'));
            item.classList.add('active');
            sections.forEach(sec => {
                sec.classList.remove('active');
                if (sec.id === target) sec.classList.add('active');
            });
            if (window.innerWidth <= 1024) sidebar.classList.remove('open');
            // Refresh history grid if needed
            if (target === 'archivos') renderFullHistory();
        });
    });

    mobileToggle.addEventListener('click', () => sidebar.classList.add('open'));
    mobileClose.addEventListener('click', () => sidebar.classList.remove('open'));
    logoutBtn.addEventListener('click', () => AUTH.logout());

    // --- Elite Drag & Drop ---
    dropZone.addEventListener('click', () => fileInput.click());

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
        dropZoneOverlay.style.display = 'flex';
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
        dropZoneOverlay.style.display = 'none';
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        dropZoneOverlay.style.display = 'none';
        const file = e.dataTransfer.files[0];
        if (file) handleExcelFile(file);
    });

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });

    function handleExcelFile(file) {
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
        
        if (!validExtensions.includes(ext)) {
            alert('Formato no soportado. Solo se aceptan archivos .xlsx, .xls o .csv');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { type: 'array' });
                processWorkbook(workbook, file.name, file.size);
            } catch (err) {
                alert('Error al leer el archivo. Asegúrate de que no esté dañado.');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function processWorkbook(workbook, fileName, fileSize) {
        currentWorkbook = workbook;
        currentFileName = fileName;

        const firstSheet = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { header: 1 });
        const headers = data[0] || [];
        const rows = data.slice(1);

        // Save COMPLETE data to history (columnas + filas)
        saveToHistory({
            name: fileName,
            size: (fileSize / 1024).toFixed(1) + ' KB',
            date: new Date().toLocaleDateString('es-ES'),
            hoja: firstSheet,
            sheets: workbook.SheetNames.length,
            rows: rows.length,
            cols: headers.length,
            columnas: headers,
            filas: rows
        });

        renderSheetTabs(workbook.SheetNames);
        loadAndRender(firstSheet, data);
        renderHistory();
    }

    function loadAndRender(sheetName, data) {
        const loadingOverlay = document.getElementById('loadingOverlay');
        // STEP 1: Show spinner
        loadingOverlay.style.display = 'block';
        excelStats.style.display = 'none';
        previewContainer.style.display = 'none';
        loadingOverlay.scrollIntoView({ behavior: 'smooth' });

        // STEP 2-4: After 800ms, show stats + table
        setTimeout(() => {
            loadingOverlay.style.display = 'none';

            currentSheetData = data;
            currentFileNameEl.textContent = currentFileName;
            currentFileMetaEl.textContent = `Hoja: ${sheetName} · ${data.length} filas detectadas`;

            updateStats(data);
            renderProTable(data);

            excelStats.style.display = 'grid';
            previewContainer.style.display = 'block';
            previewContainer.scrollIntoView({ behavior: 'smooth' });
        }, 800);
    }

    function renderSheetTabs(sheetNames) {
        sheetTabs.innerHTML = '';
        if (sheetNames.length <= 1) return;
        
        sheetNames.forEach((name, index) => {
            const tab = document.createElement('div');
            tab.className = `sheet-tab ${index === 0 ? 'active' : ''}`;
            tab.textContent = name;
            tab.onclick = () => {
                document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');
                loadSheet(name);
            };
            sheetTabs.appendChild(tab);
        });
    }

    function loadSheet(sheetName) {
        if (!currentWorkbook) return;
        const worksheet = currentWorkbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        loadAndRender(sheetName, data);
    }

    function updateStats(data) {
        const rows = data.length;
        const cols = data[0] ? data[0].length : 0;
        let empty = 0;
        let numeric = 0;

        // Simple analysis of first 50 rows for stats to avoid lag
        const sample = data.slice(0, 50);
        sample.forEach(row => {
            row.forEach(cell => {
                if (cell === null || cell === undefined || cell === "") empty++;
                if (typeof cell === 'number') numeric++;
            });
        });

        document.getElementById('totalRows').textContent = rows;
        document.getElementById('totalCols').textContent = cols;
        document.getElementById('emptyCells').textContent = empty + (rows > 50 ? '+' : '');
        document.getElementById('numericCols').textContent = numeric > 0 ? 'Detectadas' : '0';
    }

    function renderProTable(data) {
        const rowsToRender = data.slice(0, 101); // Header + 100 rows
        let html = '';
        
        rowsToRender.forEach((row, i) => {
            html += '<tr>';
            row.forEach(cell => {
                const content = cell !== undefined ? cell : '';
                html += i === 0 ? `<th>${content}</th>` : `<td>${content}</td>`;
            });
            html += '</tr>';
        });

        excelTable.innerHTML = html;
        tableCounter.textContent = `Mostrando ${Math.min(data.length, 100)} de ${data.length} filas`;
    }

    // --- Table Features ---
    tableSearch.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase();
        if (!term) {
            renderProTable(currentSheetData);
            return;
        }

        const filtered = currentSheetData.filter((row, i) => {
            if (i === 0) return true; // Keep header
            return row.some(cell => String(cell).toLowerCase().includes(term));
        });

        renderProTable(filtered);
        tableCounter.textContent = `${filtered.length - 1} coincidencias`;
    });

    closePreviewBtn.addEventListener('click', () => {
        previewContainer.style.display = 'none';
        excelStats.style.display = 'none';
        excelTable.innerHTML = '';
    });

    exportBtn.addEventListener('click', () => {
        const term = tableSearch.value.toLowerCase();
        let dataToExport = currentSheetData;
        if (term) {
            dataToExport = currentSheetData.filter((row, i) => {
                if (i === 0) return true;
                return row.some(cell => String(cell).toLowerCase().includes(term));
            });
        }
        
        const ws = XLSX.utils.aoa_to_sheet(dataToExport);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Exportado");
        XLSX.writeFile(wb, `ExcelEZ_Export_${Date.now()}.xlsx`);
    });

    // ─────────────────────────────────────────
    // ASSISTANT BRIDGE — "Consultar IA" button
    // ─────────────────────────────────────────
    askAssistantBtn.addEventListener('click', () => {
        const headers = (currentSheetData[0] || []).map(h => String(h));
        const rawRows = currentSheetData.slice(1);
        if (headers.length === 0) {
            alert('No hay datos cargados. Sube un archivo primero.');
            return;
        }

        // Convert rows to keyed objects: { ACTIVOS: 123, GASTOS: 456 }
        const filas = rawRows.map(row =>
            Object.fromEntries(headers.map((h, i) => [h, row[i] !== undefined ? row[i] : '']))
        );

        const sheetLabel = currentFileMetaEl.textContent.replace('Hoja: ', '').split(' ·')[0] || 'Hoja1';

        window.excelContexto = {
            activo: true,
            nombreArchivo: currentFileName,
            hoja: sheetLabel,
            columnas: headers,
            filas: filas,
            totalFilas: filas.length,
            totalColumnas: headers.length
        };

        // Persist (without filas to avoid quota issues on large files)
        const toSave = { ...window.excelContexto, filas: filas.slice(0, 200) };
        try { localStorage.setItem(`context_${currentUser.email}`, JSON.stringify(toSave)); } catch(e) {}

        // Navigate to chat
        document.querySelector('[data-target="consultas"]').click();
        renderContextBanner();
        renderSuggestions();

        // Welcome message
        const col1 = headers[0] || 'datos';
        const col2 = headers[1] || col1;
        const lastRow = filas.length + 1;
        const welcomeMsg =
            `📊 He cargado tu archivo **${currentFileName}**\n\n` +
            `**Hoja:** ${sheetLabel} | **${filas.length}** filas | **${headers.length}** columnas\n` +
            `**Columnas:** ${headers.join(', ')}\n\n` +
            `Puedo ayudarte con:\n` +
            `• Calcular totales, promedios, máximos y mínimos **con tus datos reales**\n` +
            `• Darte fórmulas exactas (ej: =SUMA(A2:A${lastRow}))\n` +
            `• Comparar columnas, calcular diferencias y porcentajes\n` +
            `• Responder cualquier duda general de Excel\n\n` +
            `¿Qué quieres saber?`;

        chatMessages.innerHTML = '';
        addMessage('ai', formatAIResponse(welcomeMsg));
        chatMessages.scrollTop = 0;
    });

    clearContextBtn.addEventListener('click', () => {
        window.excelContexto = null;
        localStorage.removeItem(`context_${currentUser.email}`);
        renderContextBanner();
        renderSuggestions();
        addMessage('ai', 'Contexto quitado. Ahora responderé consultas generales de Excel 📊');
    });

    function renderContextBanner() {
        const ctx = window.excelContexto;
        if (ctx && ctx.activo) {
            contextBanner.style.display = 'flex';
            contextFileName.textContent =
                `📎 ${ctx.nombreArchivo} · ${ctx.hoja} · ${ctx.totalFilas} filas · ${ctx.totalColumnas} columnas`;
        } else {
            contextBanner.style.display = 'none';
        }
    }

    // --- History Management ---
    function saveToHistory(meta) {
        const key = `history_pro_${currentUser.email}`;
        let history = JSON.parse(localStorage.getItem(key) || '[]');
        history = history.filter(h => h.name !== meta.name);
        history.unshift(meta);
        if (history.length > 10) history.pop();
        localStorage.setItem(key, JSON.stringify(history));
    }

    // Centralized: opens a saved item from history into the viewer
    function openFromHistory(item) {
        if (!item.columnas || !item.filas) {
            alert('Los datos de este archivo no están disponibles en caché. Por favor, vuelve a subirlo.');
            return;
        }
        currentFileName = item.name;
        // Build a full data array: [headers, ...rows]
        const fullData = [item.columnas, ...item.filas];
        currentSheetData = fullData;

        // Ensure we are on the Mi Excel section
        document.querySelector('[data-target="mi-excel"]').click();

        setTimeout(() => {
            sheetTabs.innerHTML = '';
            loadAndRender(item.hoja || 'Hoja1', fullData);
        }, 50);
    }

    function renderHistory() {
        const history = JSON.parse(localStorage.getItem(`history_pro_${currentUser.email}`) || '[]');
        historyList.innerHTML = '';

        if (history.length === 0) {
            historyList.innerHTML = '<p style="color: var(--text-muted); font-size: 0.8rem; text-align: center; padding: 1rem;">No hay archivos recientes.</p>';
            return;
        }

        history.slice(0, 5).forEach(item => {
            const div = document.createElement('div');
            div.className = 'history-item';
            div.style.cursor = 'pointer';
            div.title = 'Haz clic para abrir este archivo';
            div.innerHTML = `
                <i data-lucide="file-spreadsheet"></i>
                <div class="history-info">
                    <span class="history-name">${item.name}</span>
                    <span class="history-meta">${item.date} · ${item.size || '?'} · ${item.rows || '?'} filas</span>
                </div>
                <i data-lucide="chevron-right" style="width:16px; color: var(--text-muted); flex-shrink:0;"></i>
            `;
            div.addEventListener('click', () => openFromHistory(item));
            historyList.appendChild(div);
        });
        lucide.createIcons();
    }

    function renderFullHistory() {
        const history = JSON.parse(localStorage.getItem(`history_pro_${currentUser.email}`) || '[]');
        fullHistoryGrid.innerHTML = '';

        if (history.length === 0) {
            fullHistoryGrid.innerHTML = `
                <div class="empty-state" style="grid-column: 1 / -1;">
                    <i data-lucide="folder-open"></i>
                    <h3>Tu biblioteca está vacía</h3>
                    <p>Sube archivos en la sección "Mi Excel" para verlos aquí.</p>
                </div>
            `;
            lucide.createIcons();
            return;
        }

        history.forEach(item => {
            const card = document.createElement('div');
            card.className = 'card history-item-full';
            card.style.padding = '1.5rem';
            card.innerHTML = `
                <div style="display: flex; align-items: flex-start; gap: 1rem; margin-bottom: 1rem;">
                    <div style="background: rgba(0, 200, 150, 0.1); padding: 10px; border-radius: 12px; color: var(--primary);">
                        <i data-lucide="file-text"></i>
                    </div>
                    <div style="flex: 1; overflow: hidden;">
                        <h4 style="margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${item.name}">${item.name}</h4>
                        <p style="font-size: 0.8rem; color: var(--text-muted); margin: 4px 0;">${item.date} · ${item.size || '?'}</p>
                    </div>
                </div>
                <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; margin-bottom: 1.5rem;">
                    <div style="font-size: 0.75rem;"><strong style="color: var(--primary);">${item.rows || '?'}</strong> filas</div>
                    <div style="font-size: 0.75rem;"><strong style="color: var(--primary);">${item.cols || '?'}</strong> cols</div>
                    <div style="font-size: 0.75rem;"><strong style="color: var(--primary);">${item.sheets || 1}</strong> hojas</div>
                </div>
                <button class="btn-primary" style="padding: 8px; font-size: 0.85rem; width: 100%;">
                    <i data-lucide="eye" style="width: 14px;"></i> Abrir en visor
                </button>
            `;
            card.querySelector('.btn-primary').addEventListener('click', () => openFromHistory(item));
            fullHistoryGrid.appendChild(card);
        });
        lucide.createIcons();
    }

    // --- Local Intelligent Assistant ---
    const EXCEL_KNOWLEDGE_BASE = [
        {
            keywords: ['buscarv', 'vlookup', 'buscar valor', 'encontrar en tabla'],
            title: 'Función BUSCARV',
            formula: '=BUSCARV(valor_buscado, matriz_tabla, indicador_columnas, [ordenado])',
            explanation: 'Busca un valor en la primera columna de la izquierda de una tabla y luego devuelve un valor en la misma fila desde una columna especificada.',
            example: '`=BUSCARV("Juan", A2:C10, 2, FALSO)` busca "Juan" y devuelve lo que hay en la segunda columna (ej. su edad).',
            errors: 'Asegúrate de que el valor buscado esté en la PRIMERA columna del rango seleccionado. Usa FALSO para coincidencia exacta.'
        },
        {
            keywords: ['suma', 'sumar', 'total', 'matematicas'],
            title: 'Función SUMA',
            formula: '=SUMA(rango_o_celdas)',
            explanation: 'Suma todos los números en un rango de celdas.',
            example: '`=SUMA(A1:A100)` devuelve el total de la columna A.',
            errors: 'Evita espacios en blanco o texto dentro del rango de números.'
        },
        {
            keywords: ['sumarsi', 'sumar si', 'suma condicion', 'sumar con condicion'],
            title: 'Función SUMAR.SI',
            formula: '=SUMAR.SI(rango, criterio, [rango_suma])',
            explanation: 'Suma las celdas que cumplen un determinado criterio (número, texto o expresión).',
            example: '`=SUMAR.SI(A1:A10, ">100")` suma solo los valores mayores a 100.',
            errors: 'Si el rango a evaluar es distinto al de suma, especifica ambos.'
        },
        {
            keywords: ['si ', 'entonces', 'condicional', 'if', 'si anidado', 'si conjunto'],
            title: 'Funciones LÓGICAS (SI)',
            formula: '=SI(prueba_logica, valor_si_verdadero, valor_si_falso)',
            explanation: 'Permite realizar una comparación lógica entre un valor y el resultado que espera. Úsalo para categorizar datos.',
            example: '`=SI(B2>=5, "Aprobado", "Suspenso")` verifica si la nota es suficiente.',
            errors: 'Recuerda poner el texto entre comillas dobles. Máximo 64 niveles de anidamiento.'
        },
        {
            keywords: ['contar', 'contarsi', 'contar si', 'contar celdas', 'cuantos'],
            title: 'Función CONTAR.SI',
            formula: '=CONTAR.SI(rango, criterio)',
            explanation: 'Cuenta cuántas celdas en un rango cumplen una condición específica.',
            example: '`=CONTAR.SI(C2:C20, "Vendido")` cuenta cuántas ventas se han realizado.',
            errors: 'Para contar texto exacto, asegúrate de escribirlo igual (sensible a tildes).'
        },
        {
            keywords: ['buscarx', 'xlookup', 'buscar moderno'],
            title: 'Función BUSCARX',
            formula: '=BUSCARX(valor_buscado, matriz_busqueda, matriz_devuelta)',
            explanation: 'La versión moderna y superior de BUSCARV. No requiere que el valor esté en la primera columna.',
            example: '`=BUSCARX("A-01", A:A, C:C)` busca ID en columna A y devuelve nombre en columna C.',
            errors: 'Solo disponible en Office 365 y Excel 2021+.'
        },
        {
            keywords: ['concatenar', 'unir texto', 'unir celdas', '&', 'unir'],
            title: 'CONCATENAR / Unir Texto',
            formula: '=CONCATENAR(A1, " ", B1)  o  =A1 & " " & B1',
            explanation: 'Une dos o más cadenas de texto en una sola celda. El símbolo & es el método más rápido.',
            example: '`=A2 & ", " & B2` uniría "Apellido" con "Nombre".',
            errors: 'No olvides añadir el espacio en blanco manual `" "` si lo necesitas.'
        },
        {
            keywords: ['hoy', 'ahora', 'fecha actual', 'fecha de hoy', 'hora'],
            title: 'Funciones de FECHA (HOY)',
            formula: '=HOY()  o  =AHORA()',
            explanation: 'Devuelven la fecha actual o la fecha y hora exacta del sistema.',
            example: '`=DIAS(HOY(), A1)` calcula los días transcurridos desde la fecha en A1.',
            errors: 'Se actualizan cada vez que la hoja se recalcula.'
        },
        {
            keywords: ['diferencia fechas', 'dias entre', 'restar fechas'],
            title: 'Cálculo de fechas',
            formula: '=DIAS(fecha_fin, fecha_inicio)  o  =fecha2 - fecha1',
            explanation: 'Calcula el número de días entre dos fechas.',
            example: '`=HOY() - A2` devuelve cuántos días han pasado desde la fecha en A2.',
            errors: 'Asegúrate de que las celdas tengan formato de Fecha, no de Texto.'
        },
        {
            keywords: ['promedio', 'media', 'average'],
            title: 'Función PROMEDIO',
            formula: '=PROMEDIO(A1:A50)',
            explanation: 'Calcula la media aritmética de un grupo de números.',
            example: '`=PROMEDIO(B2:B10)` para saber la nota media de un grupo.',
            errors: 'Ignora celdas vacías, pero incluye celdas con valor cero (0).'
        }
    ];

    sendBtn.addEventListener('click', () => handleChatMessage());
    chatInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleChatMessage();
        }
    });

    async function handleChatMessage() {
        const query = chatInput.value.trim();
        if (!query) return;

        addMessage('user', query);
        chatInput.value = '';

        const typingId = addMessage('ai', 'Pensando en una respuesta...');
        
        // Simulación de procesamiento local
        setTimeout(() => {
            const response = getAssistantResponse(query);
            updateMessage(typingId, formatAIResponse(response));
            saveChatHistory(query, response);
        }, 800);
    }

    // ─────────────────────────────────────────
    // SMART RESPONSE ENGINE
    // ─────────────────────────────────────────
    function getAssistantResponse(query) {
        const q = query.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        const ctx = window.excelContexto;

        if (ctx && ctx.activo && ctx.filas && ctx.filas.length > 0) {

            // Helper: get numeric values from a column
            function colVals(colName) {
                return ctx.filas.map(r => parseFloat(r[colName])).filter(v => !isNaN(v));
            }
            // Helper: Excel column letter A, B, C...
            function colLetter(colName) {
                const idx = ctx.columnas.indexOf(colName);
                return idx >= 0 ? getExcelColumnName(idx) : '?';
            }
            function excelRange(colName) {
                const L = colLetter(colName);
                return `${L}2:${L}${ctx.totalFilas + 1}`;
            }
            function fmt(n) { return typeof n === 'number' ? n.toLocaleString('es-ES') : n; }

            // Detect mentioned column (case-insensitive)
            const mentionedCol = ctx.columnas.find(c =>
                q.includes(c.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''))
            );

            // ── SUMA / TOTAL ──────────────────────────
            if (/suma|total|cuanto suman|cuantos suman|sumar/.test(q)) {
                const col = mentionedCol;
                if (col) {
                    const vals = colVals(col);
                    const total = vals.reduce((a, b) => a + b, 0);
                    const formula = `=SUMA(${excelRange(col)})`;
                    return `### 📊 Resultado: ${col}\n` +
                           `El total de **${col}** es **${fmt(total)}**\n\n` +
                           `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                           `💡 Cómo usarla: Escríbela en cualquier celda vacía de tu hoja.`;
                } else {
                    // Sum all numeric columns
                    const numCols = ctx.columnas.filter(c => colVals(c).length > 0);
                    if (numCols.length > 0) {
                        const lines = numCols.map(c => {
                            const s = colVals(c).reduce((a, b) => a + b, 0);
                            return `• **${c}:** ${fmt(s)}`;
                        }).join('\n');
                        return `### 📊 Totales por columna\n${lines}\n\n¿Quieres el total de alguna columna específica?`;
                    }
                }
            }

            // ── PROMEDIO / MEDIA ──────────────────────
            if (/promedio|media|average|promedi/.test(q)) {
                const col = mentionedCol;
                if (col) {
                    const vals = colVals(col);
                    if (vals.length === 0) return `No hay valores numéricos en la columna **${col}**.`;
                    const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
                    const formula = `=PROMEDIO(${excelRange(col)})`;
                    return `### 📈 Resultado: ${col}\n` +
                           `El promedio de **${col}** es **${fmt(Math.round(avg * 100) / 100)}**\n\n` +
                           `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                           `💡 Cómo usarla: Escríbela en cualquier celda vacía.`;
                }
            }

            // ── MÁXIMO ───────────────────────────────
            if (/maximo|mayor|mas alto|mas grande|el mas/.test(q)) {
                const col = mentionedCol;
                if (col) {
                    const vals = colVals(col);
                    const max = Math.max(...vals);
                    const formula = `=MAX(${excelRange(col)})`;
                    return `### 🏆 Valor Máximo: ${col}\n` +
                           `El valor más alto de **${col}** es **${fmt(max)}**\n\n` +
                           `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                           `💡 Cómo usarla: Devuelve el número máximo del rango.`;
                }
            }

            // ── MÍNIMO ───────────────────────────────
            if (/minimo|menor|mas bajo|mas pequeno/.test(q)) {
                const col = mentionedCol;
                if (col) {
                    const vals = colVals(col);
                    const min = Math.min(...vals);
                    const formula = `=MIN(${excelRange(col)})`;
                    return `### 📉 Valor Mínimo: ${col}\n` +
                           `El valor más bajo de **${col}** es **${fmt(min)}**\n\n` +
                           `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\``;
                }
            }

            // ── DIFERENCIA / MENOS ───────────────────
            if (/diferencia|menos|resta|restar|fila por fila/.test(q) && ctx.columnas.length >= 2) {
                const numCols = ctx.columnas.filter(c => colVals(c).length > 0);
                if (numCols.length >= 2) {
                    const c1 = mentionedCol || numCols[0];
                    const c2 = numCols.find(c => c !== c1) || numCols[1];
                    const rows = ctx.filas.map((r, i) => {
                        const v1 = parseFloat(r[c1]) || 0;
                        const v2 = parseFloat(r[c2]) || 0;
                        return `Fila ${i + 2}: ${fmt(v1)} - ${fmt(v2)} = **${fmt(v1 - v2)}**`;
                    });
                    const L1 = colLetter(c1); const L2 = colLetter(c2);
                    return `### ➗ Diferencia ${c1} - ${c2}\n` +
                           rows.join('\n') + '\n\n' +
                           `📋 Fórmula Excel para cada fila:\n\`\`\`excel\n=${L1}2-${L2}2\n\`\`\`\n` +
                           `💡 Arrástrala hacia abajo para calcular todas las filas.`;
                }
            }

            // ── PORCENTAJE ────────────────────────────
            if (/porcent|porcentaje|que porcentaje|representa/.test(q)) {
                const col = mentionedCol || ctx.columnas.find(c => colVals(c).length > 0);
                if (col) {
                    const vals = colVals(col);
                    const total = vals.reduce((a, b) => a + b, 0);
                    if (total === 0) return `El total de **${col}** es cero, no se pueden calcular porcentajes.`;
                    const L = colLetter(col);
                    const rows = ctx.filas.map((r, i) => {
                        const v = parseFloat(r[col]) || 0;
                        return `Fila ${i + 2}: ${fmt(v)} → **${((v / total) * 100).toFixed(1)}%**`;
                    });
                    const totalRow = ctx.totalFilas + 2;
                    return `### 📊 Porcentajes de ${col}\nTotal: ${fmt(total)}\n\n` +
                           rows.join('\n') + '\n\n' +
                           `📋 Fórmula Excel:\n\`\`\`excel\n=${L}2/SUMA(${L}2:${L}${ctx.totalFilas + 1})\n\`\`\`\n` +
                           `💡 Formatea la celda como % para ver el resultado directamente.`;
                }
            }

            // ── FILTRAR ───────────────────────────────
            if (/filtr|filtrar/.test(q)) {
                const col = mentionedCol || ctx.columnas[0];
                const L = colLetter(col);
                const formula = `=FILTRAR(${L}2:${L}${ctx.totalFilas + 1},${L}2:${L}${ctx.totalFilas + 1}>0)`;
                return `### 🔎 Filtrar ${col}\n` +
                       `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                       `💡 Cambia "0" por el valor que quieres como límite de filtro. Solo disponible en Excel 365/2021+.`;
            }

            // ── ORDENAR ───────────────────────────────
            if (/orden|ordenar|de mayor a menor|de menor a mayor/.test(q)) {
                const desc = q.includes('mayor') || q.includes('desc');
                const firstL = colLetter(ctx.columnas[0]);
                const lastL = colLetter(ctx.columnas[ctx.columnas.length - 1]);
                const formula = `=ORDENAR(${firstL}2:${lastL}${ctx.totalFilas + 1},1,${desc ? -1 : 1})`;
                return `### 🔃 Ordenar datos ${desc ? 'de mayor a menor' : 'de menor a mayor'}\n` +
                       `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                       `💡 El tercer argumento \`-1\` = descendente, \`1\` = ascendente. Solo Excel 365/2021+.`;
            }

            // ── SUMAR.SI ─────────────────────────────
            if (/sumar.si|suma.*condicion|sumar.*mayor|sumar.*solo/.test(q)) {
                const col = mentionedCol || ctx.columnas.find(c => colVals(c).length > 0);
                if (col) {
                    const L = colLetter(col);
                    const formula = `=SUMAR.SI(${L}2:${L}${ctx.totalFilas + 1},">500000")`;
                    return `### 🧮 SUMAR.SI para ${col}\n` +
                           `📋 Fórmula Excel:\n\`\`\`excel\n${formula}\n\`\`\`\n` +
                           `💡 Sustituye \`>500000\` por el criterio que necesitas (ej: >1000000, "Vendido", etc.).`;
                }
            }

            // ── RÉSUMEN / ANALIZA ─────────────────────
            if (/resum|resumen|analiz|estadistic|todo.*datos|informe/.test(q)) {
                const numCols = ctx.columnas.filter(c => colVals(c).length > 0);
                let report = `### 📋 Resumen de **"${ctx.nombreArchivo}"**\n`;
                report += `${ctx.totalFilas} filas · ${ctx.totalColumnas} columnas\n\n`;
                numCols.forEach(col => {
                    const vals = colVals(col);
                    const sum = vals.reduce((a, b) => a + b, 0);
                    const avg = sum / vals.length;
                    const max = Math.max(...vals);
                    const min = Math.min(...vals);
                    report += `**${col}** (${colLetter(col)}):\n`;
                    report += `• Total: ${fmt(sum)} | Promedio: ${fmt(Math.round(avg))} | Máx: ${fmt(max)} | Mín: ${fmt(min)}\n`;
                });
                const emptyCount = ctx.filas.reduce((acc, row) =>
                    acc + Object.values(row).filter(v => v === null || v === undefined || v === '').length, 0
                );
                report += `\n⚠️ Celdas vacías detectadas: **${emptyCount}**`;
                return report;
            }

            // ── ERRORES / VACIAS ─────────────────────
            if (/vacia|vacias|hueco|huec|error|incompleto/.test(q)) {
                const emptyCount = ctx.filas.reduce((acc, row) =>
                    acc + Object.values(row).filter(v => v === null || v === undefined || v === '').length, 0
                );
                const perCol = ctx.columnas.map(c => {
                    const empty = ctx.filas.filter(r => r[c] === null || r[c] === undefined || r[c] === '').length;
                    return empty > 0 ? `• **${c}**: ${empty} celda(s) vacía(s)` : null;
                }).filter(Boolean);
                if (emptyCount === 0) {
                    return `✅ ¡Perfecto! No hay celdas vacías en el archivo. Los datos están completos.`;
                }
                return `⚠️ Se detectaron **${emptyCount}** celdas vacías:\n${perCol.join('\n')}\n\n💡 Para reemplazarlas con cero:\n\`\`\`excel\n=SI(ESBLANCO(A2),0,A2)\n\`\`\``;
            }

            // ── QUÉ COLUMNA TIENE VALORES MÁS ALTOS ─
            if (/columna.*mayor|columna.*alta|mas alta|cual.*mayor/.test(q)) {
                const numCols = ctx.columnas.filter(c => colVals(c).length > 0);
                const sums = numCols.map(c => ({ col: c, sum: colVals(c).reduce((a, b) => a + b, 0) }));
                sums.sort((a, b) => b.sum - a.sum);
                const lines = sums.map((s, i) => `${i + 1}. **${s.col}**: Total ${fmt(s.sum)}`).join('\n');
                return `### 📊 Comparación de columnas numéricas\n${lines}\n\n` +
                       `La columna con valores más altos es **${sums[0].col}**.`;
            }

            // If context active but no specific match, give suggestions
            const c1 = ctx.columnas[0] || 'datos';
            const c2 = ctx.columnas[1] || c1;
            return `Tengo cargado **${ctx.nombreArchivo}** y estoy listo para analizar tus datos 📊\n\n` +
                   `Prueba preguntarme:\n` +
                   `• "¿Cuánto suman los **${c1}**?"\n` +
                   `• "¿Cuál es el promedio de **${c2}**?"\n` +
                   `• "Resume todos los datos"\n` +
                   `• "¿Hay celdas vacías?"\n` +
                   `• "¿Qué columna tiene valores más altos?"`;
        }

        // ── GENERAL EXCEL KNOWLEDGE BASE ─────────────
        const match = EXCEL_KNOWLEDGE_BASE.find(item =>
            item.keywords.some(kw => q.includes(kw))
        );
        if (match) {
            return `### 📋 ${match.title}\n` +
                   `**Fórmula:**\n\`\`\`excel\n${match.formula}\n\`\`\`\n` +
                   `**📖 Explicación:**\n${match.explanation}\n\n` +
                   `**💡 Ejemplo:**\n${match.example}\n\n` +
                   `**⚠️ Errores comunes:**\n${match.errors}`;
        }

        const isExcelRelated = ['excel','tabla','fila','columna','celda','hoja','datos','formula','calculo','funcion'].some(w => q.includes(w));
        if (isExcelRelated) {
            return `No encontré una fórmula exacta para esa consulta. Prueba con:\n\n` +
                   EXCEL_KNOWLEDGE_BASE.slice(0, 3).map(m => `• **${m.title}**`).join('\n') +
                   `\n\n¿O tienes un archivo cargado del que quieras analizar datos?`;
        }

        const ctx2 = window.excelContexto;
        return `Soy especialista en Excel 📊 ${ctx2 ? `Tengo **${ctx2.nombreArchivo}** cargado y listo.` : ''}\n\n¿Tienes alguna duda sobre tu archivo o sobre fórmulas?\n• **"Resume los datos"**\n• **"¿Cómo uso BUSCARV?"**`;
    }

    function formatAIResponse(text) {
        // Formatear bloques de código excel: ```excel ... ```
        const codeBlockRegex = /```excel\n([\s\S]*?)```/g;
        let formattedText = text.replace(codeBlockRegex, (match, formula) => {
            const cleanFormula = formula.trim();
            const safeFormula = cleanFormula.replace(/'/g, "\\'");
            return `
                <div class="code-container" style="background: rgba(0, 0, 0, 0.3); padding: 1rem; border-radius: 8px; margin: 10px 0; border: 1px solid var(--border); position: relative;">
                    <pre style="margin: 0; overflow-x: auto;"><code class="language-excel" style="color: var(--primary); font-family: 'Courier New', monospace;">${cleanFormula}</code></pre>
                    <button class="copy-btn" onclick="copyToClipboard('${safeFormula}', this)" style="margin-top: 10px; background: var(--primary); color: #000; border: none; padding: 5px 12px; border-radius: 4px; cursor: pointer; font-size: 0.8rem; font-weight: 600; display: flex; align-items: center; gap: 5px;">
                        <i data-lucide="copy" style="width: 14px;"></i> Copiar fórmula
                    </button>
                </div>
            `;
        });

        // Formatear negritas y saltos de línea
        formattedText = formattedText.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        formattedText = formattedText.replace(/### (.*?)\n/g, '<h3>$1</h3>');
        formattedText = formattedText.replace(/\n/g, '<br>');

        return formattedText;
    }

    window.copyToClipboard = function(text, btn) {
        navigator.clipboard.writeText(text).then(() => {
            const originalHTML = btn.innerHTML;
            btn.innerHTML = '<i data-lucide="check" style="width: 14px;"></i> ¡Copiado!';
            btn.classList.add('copied');
            btn.style.background = '#00C896';
            lucide.createIcons();
            
            setTimeout(() => {
                btn.innerHTML = originalHTML;
                btn.classList.remove('copied');
                btn.style.background = 'var(--primary)';
                lucide.createIcons();
            }, 2000);
        });
    };

    function addMessage(type, text) {
        const messageDiv = document.createElement('div');
        messageDiv.className = `message ${type}`;
        const id = 'msg-' + Date.now();
        messageDiv.id = id;
        
        if (type === 'ai') {
            messageDiv.innerHTML = text;
        } else {
            const p = document.createElement('p');
            p.textContent = text;
            messageDiv.appendChild(p);
        }
        
        chatMessages.appendChild(messageDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
        lucide.createIcons();
        return id;
    }

    function updateMessage(id, text) {
        const msg = document.getElementById(id);
        if (msg) {
            msg.innerHTML = text;
            chatMessages.scrollTop = chatMessages.scrollHeight;
            lucide.createIcons();
        }
    }

    function saveChatHistory(query, response) {
        const history = JSON.parse(localStorage.getItem(`chat_${currentUser.email}`) || '[]');
        history.push({ query, response, date: new Date().toISOString() });
        localStorage.setItem(`chat_${currentUser.email}`, JSON.stringify(history));
    }

    function loadChatHistory() {
        // Limpiar para evitar duplicidad si se llama varias veces
        chatMessages.innerHTML = '';
        const history = JSON.parse(localStorage.getItem(`chat_${currentUser.email}`) || '[]');
        if (history.length === 0) {
            addMessage('ai', '¡Hola! Soy tu asistente experto en Excel 📊 ¿En qué puedo ayudarte hoy?');
        } else {
            history.forEach(item => {
                addMessage('user', item.query);
                addMessage('ai', formatAIResponse(item.response));
            });
        }
    }

    function renderSuggestions() {
        const existing = document.querySelector('.quick-suggestions');
        if (existing) existing.remove();

        const ctx = window.excelContexto;
        const suggestionBox = document.createElement('div');
        suggestionBox.className = 'quick-suggestions';
        suggestionBox.style.cssText = 'padding:1rem;display:flex;flex-wrap:wrap;gap:10px;border-top:1px solid var(--border)';

        let topics = [];
        if (ctx && ctx.activo) {
            const col1 = ctx.columnas[0] || 'datos';
            const col2 = ctx.columnas[1] || col1;
            topics = [
                { text: `¿Cuánto suman los ${col1}?`,   query: `cuanto suman los ${col1}` },
                { text: `Promedio de ${col2}`,             query: `promedio de ${col2}` },
                { text: '¿Hay celdas vacías?',             query: 'hay celdas vacias' },
                { text: 'Resume todos los datos',          query: 'resume todos los datos' },
                { text: `Fórmulas para ${col1}`,           query: `sumar ${col1} con condicion` }
            ];
        } else {
            topics = [
                { text: '¿Cómo uso BUSCARV?',          query: 'como usar buscarv' },
                { text: '¿Cómo sumo con condición?',   query: 'sumar con condicion' },
                { text: 'SI anidado',                  query: 'si anidado' },
                { text: 'Restar fechas',               query: 'restar fechas' }
            ];
        }

        topics.forEach(topic => {
            const btn = document.createElement('button');
            btn.className = 'suggestion-pill';
            btn.style.cssText = 'background:rgba(0,200,150,0.1);border:1px solid var(--primary);color:var(--primary);padding:8px 15px;border-radius:20px;font-size:0.8rem;cursor:pointer;transition:all 0.3s ease';
            btn.textContent = topic.text;
            btn.onmouseover = () => { btn.style.background = 'var(--primary)'; btn.style.color = '#000'; };
            btn.onmouseout  = () => { btn.style.background = 'rgba(0,200,150,0.1)'; btn.style.color = 'var(--primary)'; };
            btn.onclick = () => { chatInput.value = topic.text; handleChatMessage(); };
            suggestionBox.appendChild(btn);
        });

        chatMessages.parentNode.insertBefore(suggestionBox, chatInput.closest('.chat-input-wrapper'));
    }

    // Inicializar historial
    loadChatHistory();
});
