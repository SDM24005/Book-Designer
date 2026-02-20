// Book Designer - Script principale

document.addEventListener('DOMContentLoaded', function() {
    const nomeProgettoInput = document.getElementById('nome-progetto');
    const formatoSelect = document.getElementById('formato');
    const orientamentoSelect = document.getElementById('orientamento');
    const tecnicaSelect = document.getElementById('tecnica-stampa');
    const foglioSelect = document.getElementById('foglio-macchina');
    const formatoApertoDisplay = document.getElementById('formato-aperto');
    const foglioWarning = document.getElementById('foglio-warning');
    const numeroPagineInput = document.getElementById('numero-pagine');
    const pagineWarning = document.getElementById('pagine-warning');
    const customInputs = document.getElementById('custom-inputs');
    const customWidth = document.getElementById('custom-width');
    const customHeight = document.getElementById('custom-height');
    const swapButton = document.getElementById('swap-dimensions');
    const widthError = document.getElementById('width-error');
    const heightError = document.getElementById('height-error');
    const segnatureList = document.getElementById('segnature-list');
    const spineDisplay = document.getElementById('spine-display');
    const platesDisplay = document.getElementById('plates-display');
    const platesGroup = document.getElementById('plates-group');
    const ottimizzazioneContainer = document.getElementById('ottimizzazione-container');
    const suggerimentoOttimizzato = document.getElementById('suggerimento-ottimizzato');
    const signaturesConfigList = document.getElementById('signatures-config-list');
    const addSignatureBtn = document.getElementById('add-signature-btn');
    const autoSignaturesBtn = document.getElementById('auto-signatures-btn');
    const resetSignaturesBtn = document.getElementById('reset-signatures-btn');
    const pagineResidueDisplay = document.getElementById('pagine-residue');
    const cartaTipoSelect = document.getElementById('carta-tipo');
    const cartaGrammaturaSelect = document.getElementById('carta-grammatura');
    const resetProjectBtn = document.getElementById('reset-project');
    const toggleViewBtn = document.getElementById('toggle-view-btn');
    const mobileTabDataBtn = document.getElementById('mobile-tab-data');
    const mobileTabImpositionBtn = document.getElementById('mobile-tab-imposition');
    const mobileTab3DBtn = document.getElementById('mobile-tab-3d');
    const previewSection = document.querySelector('.preview-section');
    const dataSection = document.querySelector('.data-section');
    const preview3D = document.getElementById('preview3D');
    const impositionView = document.getElementById('imposition-view');
    const tiraturaInput = document.getElementById('tiratura');
    const fogliMacchinaDisplay = document.getElementById('fogli-macchina-display');
    let currentView = 'imposition'; // '3d' o 'imposition'

    let segnatureConfigurate = [];
    function updateImpositionView() {
        if (!impositionView) return;
        impositionView.innerHTML = '';

        if (segnatureConfigurate.length === 0) {
            impositionView.innerHTML = '<div class="imposition-title">Nessuna segnatura inserita</div>';
            return;
        }

        const container = document.createElement('div');
        container.className = 'imposition-sheet';

        let pageStart = 1;

        segnatureConfigurate.forEach((sig, index) => {
            const type = sig.pagine;
            const variante = sig.variante || null;
            const pattern = getImpositionPattern(type, variante);

            if (!pattern) {
                const section = document.createElement('div');
                section.className = 'imposition-section';
                section.dataset.type = type;

                const title = document.createElement('div');
                title.className = 'imposition-title';
                title.textContent = `Segnatura ${index + 1} (${type}°) - Schema non disponibile`;
                section.appendChild(title);
                container.appendChild(section);
                pageStart += type || 0;
                return;
            }

            const { cols, rows, bianca, volta, label } = pattern;
            const offset = pageStart - 1;
            const rangeEnd = pageStart + type - 1;

            const biancaGrid = applyPageOffset(bianca, offset);
            const voltaGrid = applyPageOffset(volta, offset);

            const section = document.createElement('div');
            section.className = 'imposition-section';
            section.dataset.type = type;

            const title = document.createElement('div');
            title.className = 'imposition-title';
            title.textContent = `${label} — Segnatura ${index + 1} (${pageStart}–${rangeEnd})`;
            section.appendChild(title);

            const pair = document.createElement('div');
            pair.className = 'imposition-pair';

            const biancaSide = createImpositionSide('Bianca', biancaGrid, cols, rows, type, pattern.cssClass);
            pair.appendChild(biancaSide);

            const voltaSide = createImpositionSide('Volta', voltaGrid, cols, rows, type, pattern.cssClass);
            pair.appendChild(voltaSide);

            section.appendChild(pair);
            container.appendChild(section);

            pageStart += type;
        });

        impositionView.appendChild(container);
    }

    function getImpositionPattern(type, variante) {
        if (type === 4) {
            return {
                label: 'Quartino (4°)',
                cols: 2,
                rows: 1,
                bianca: [{ n: 4, r: false }, { n: 1, r: false }],
                volta: [{ n: 2, r: false }, { n: 3, r: false }]
            };
        }
        if (type === 8) {
            if (variante === 'finestra') {
                return {
                    label: 'Ottavo a finestra (8°)',
                    cols: 4,
                    rows: 1,
                    cssClass: 'imposition-grid-8-finestra',
                    bianca: [{ n: 7, r: false }, { n: 8, r: false }, { n: 1, r: false }, { n: 2, r: false }],
                    volta: [{ n: 3, r: false }, { n: 4, r: false }, { n: 5, r: false }, { n: 6, r: false }]
                };
            }
            return {
                label: 'Ottavo (8°)',
                cols: 2,
                rows: 2,
                bianca: [
                    { n: 8, r: false }, { n: 5, r: false },
                    { n: 1, r: false }, { n: 4, r: false }
                ],
                volta: [
                    { n: 6, r: false }, { n: 7, r: false },
                    { n: 3, r: false }, { n: 2, r: false }
                ]
            };
        }
        if (type === 12) {
            return {
                label: 'Dodicesimo (12°)',
                cols: 3,
                rows: 2,
                bianca: [
                    { n: 11, r: false, b: 'right' }, { n: 10, r: false, b: 'left' }, { n: 7, r: false, b: 'right' },
                    { n: 2, r: false, b: 'right' }, { n: 3, r: false, b: 'left' }, { n: 9, r: false, b: 'right' }
                ],
                volta: [
                    { n: 8, r: false, b: 'left' }, { n: 6, r: false, b: 'right' }, { n: 12, r: false, b: 'left' },
                    { n: 5, r: false, b: 'left' }, { n: 4, r: false, b: 'right' }, { n: 1, r: false, b: 'left' }
                ]
            };
        }
        if (type === 16) {
            return {
                label: 'Sedicesimo (16°)',
                cols: 4,
                rows: 2,
                bianca: [
                    { n: 5, r: true }, { n: 12, r: true }, { n: 9, r: true }, { n: 8, r: true },
                    { n: 4, r: false }, { n: 13, r: false }, { n: 16, r: false }, { n: 1, r: false }
                ],
                volta: [
                    { n: 7, r: true }, { n: 10, r: true }, { n: 11, r: true }, { n: 6, r: true },
                    { n: 2, r: false }, { n: 15, r: false }, { n: 14, r: false }, { n: 3, r: false }
                ]
            };
        }
        if (type === 32) {
            return {
                label: 'Trentaduesimo (32°)',
                cols: 4,
                rows: 4,
                bianca: [
                    { n: 12, r: true }, { n: 5, r: true }, { n: 8, r: true }, { n: 9, r: true },
                    { n: 21, r: false }, { n: 28, r: false }, { n: 25, r: false }, { n: 24, r: false },
                    { n: 20, r: false }, { n: 29, r: false }, { n: 32, r: false }, { n: 17, r: false },
                    { n: 13, r: false }, { n: 4, r: false }, { n: 1, r: false }, { n: 16, r: false }
                ],
                volta: [
                    { n: 10, r: true }, { n: 7, r: true }, { n: 6, r: true }, { n: 11, r: true },
                    { n: 23, r: false }, { n: 22, r: false }, { n: 27, r: false }, { n: 26, r: false },
                    { n: 18, r: false }, { n: 31, r: false }, { n: 30, r: false }, { n: 19, r: false },
                    { n: 15, r: false }, { n: 2, r: false }, { n: 3, r: false }, { n: 14, r: false }
                ]
            };
        }
        return null;
    }

    function applyPageOffset(grid, offset) {
        return grid.map(cell => {
            if (cell.n === '?' || cell.n == null) {
                return { ...cell };
            }
            return { ...cell, n: cell.n + offset };
        });
    }

    function createImpositionSide(label, grid, cols, rows, type, extraGridClass) {
        const side = document.createElement('div');
        side.className = 'imposition-side';

        const sideLabel = document.createElement('div');
        sideLabel.className = 'side-label';
        sideLabel.textContent = label;
        side.appendChild(sideLabel);

        const gridEl = document.createElement('div');
        gridEl.className = 'imposition-grid';
        if (extraGridClass) {
            gridEl.classList.add(extraGridClass);
        }
        if (type === 4) {
            gridEl.classList.add('imposition-grid-4');
        }
        if (type === 8) {
            gridEl.classList.add('imposition-grid-8');
        }
        if (type === 12) {
            gridEl.classList.add('imposition-grid-12');
        }
        if (type === 16) {
            gridEl.classList.add('imposition-grid-16');
        }
        if (type === 32) {
            gridEl.classList.add('imposition-grid-32');
        }
        gridEl.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
        gridEl.style.gridTemplateRows = `repeat(${rows}, 1fr)`;

        grid.forEach((cell, index) => {
            const cellEl = document.createElement('div');
            cellEl.className = `imposition-cell ${cell.r ? 'rotated' : ''}`;
            if (type === 8 && extraGridClass !== 'imposition-grid-8-finestra') {
                const colIndex = index % cols;
                cellEl.classList.add(colIndex % 2 === 0 ? 'base-left' : 'base-right');
            }
            if (type === 12) {
                const base = cell.b === 'left' || cell.b === 'right'
                    ? cell.b
                    : ((index % cols) % 2 === 0 ? 'left' : 'right');
                cellEl.classList.add(base === 'left' ? 'base-left' : 'base-right');
            }
            if (type === 32) {
                const colIndex = index % cols;
                const shouldBeLeft = colIndex % 2 === 0;
                const useLeftClass = cell.r ? !shouldBeLeft : shouldBeLeft;
                cellEl.classList.add(useLeftClass ? 'base-left' : 'base-right');
            }
            const num = document.createElement('span');
            num.className = 'page-num';
            num.textContent = cell.n;
            cellEl.appendChild(num);
            gridEl.appendChild(cellEl);
        });

        side.appendChild(gridEl);
        return side;
    }

    function isMobileView() {
        return window.matchMedia('(max-width: 768px)').matches;
    }

    function setPreviewView(view) {
        currentView = view;
        if (currentView === 'imposition') {
            if (preview3D) {
                preview3D.style.display = 'none';
            }
            if (impositionView) {
                impositionView.style.display = 'flex';
            }
            if (toggleViewBtn) {
                toggleViewBtn.textContent = 'Visualizza 3D';
            }
            updateImpositionView();
        } else {
            if (preview3D) {
                preview3D.style.display = 'block';
            }
            if (impositionView) {
                impositionView.style.display = 'none';
            }
            if (toggleViewBtn) {
                toggleViewBtn.textContent = 'Visualizza Imposizione';
            }
        }
    }

    function setMobileTab(tab) {
        document.body.classList.remove('mobile-view-data', 'mobile-view-imposition', 'mobile-view-3d');
        document.body.classList.add(`mobile-view-${tab}`);

        if (mobileTabDataBtn) {
            mobileTabDataBtn.classList.toggle('active', tab === 'data');
        }
        if (mobileTabImpositionBtn) {
            mobileTabImpositionBtn.classList.toggle('active', tab === 'imposition');
        }
        if (mobileTab3DBtn) {
            mobileTab3DBtn.classList.toggle('active', tab === '3d');
        }

        if (tab === 'imposition') {
            setPreviewView('imposition');
        } else if (tab === '3d') {
            setPreviewView('3d');
        }

        if (tab === 'data' && dataSection && isMobileView()) {
            dataSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        if ((tab === 'imposition' || tab === '3d') && previewSection && isMobileView()) {
            previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    if (toggleViewBtn) {
        toggleViewBtn.addEventListener('click', function() {
            setPreviewView(currentView === '3d' ? 'imposition' : '3d');
        });
    }

    if (mobileTabDataBtn) {
        mobileTabDataBtn.addEventListener('click', function() {
            setMobileTab('data');
        });
    }

    if (mobileTabImpositionBtn) {
        mobileTabImpositionBtn.addEventListener('click', function() {
            setMobileTab('imposition');
        });
    }
    if (mobileTab3DBtn) {
        mobileTab3DBtn.addEventListener('click', function() {
            setMobileTab('3d');
        });
    }

    if (isMobileView()) {
        setMobileTab('data');
    }

    window.addEventListener('resize', () => {
        if (isMobileView()) {
            if (
                !document.body.classList.contains('mobile-view-data') &&
                !document.body.classList.contains('mobile-view-imposition') &&
                !document.body.classList.contains('mobile-view-3d')
            ) {
                setMobileTab('data');
            }
            return;
        }
        document.body.classList.remove('mobile-view-data', 'mobile-view-imposition', 'mobile-view-3d');
    });
    let availableSignatureTypes = new Set();
    const PAPER_THICKNESS = {
        patinata_lucida: {
            80: 68, 90: 76, 100: 85, 115: 97, 130: 110, 150: 127, 170: 144, 200: 170, 250: 212, 300: 255
        },
        patinata_opaca: {
            80: 80, 90: 90, 100: 100, 115: 115, 130: 130, 150: 150, 170: 170, 200: 200, 250: 250, 300: 300
        },
        usomano: {
            80: 100, 90: 112, 100: 125, 115: 143, 130: 162, 150: 187, 170: 212, 200: 250, 250: 312, 300: 375
        },
        edizioni: {
            80: 140, 90: 157.5, 100: 175, 115: 190, 130: 215
        }
    };
    const COLLAPSE_FACTOR = {
        usomano: 0.95,
        edizioni: 0.95
    };
    const GLUE_MM = 1;
    const PAPER_TYPE_OPTIONS = [
        { value: 'patinata_lucida', label: 'Patinata Lucida' },
        { value: 'patinata_opaca', label: 'Patinata Opaca' },
        { value: 'usomano', label: 'Usomano' },
        { value: 'edizioni', label: 'Edizioni / Avoriata' }
    ];
    const GRAMMATURA_OPTIONS = [80, 90, 100, 115, 130, 150, 170, 200, 250, 300];
    
    const previewContainer = document.getElementById('preview3D');
    const parallelepiped = document.querySelector('.parallelepiped');
    
    const widthLabel = document.querySelector('.dimension-label.width');
    const heightLabel = document.querySelector('.dimension-label.height');
    const depthLabel = document.querySelector('.dimension-label.depth');

    let rotationX = -20;
    let rotationY = 30;
    let zoom = 1;

    function applyPreviewTransform() {
        if (!parallelepiped) {
            return;
        }
        parallelepiped.style.transform = `translate(-50%, -50%) rotateX(${rotationX}deg) rotateY(${rotationY}deg) scale(${zoom})`;
    }

    const dimensioniPreset = {
        'A4': { width: 210, height: 297, depth: 30 },
        'A5': { width: 148, height: 210, depth: 25 },
        'B5': { width: 176, height: 250, depth: 25 },
        '17x24': { width: 170, height: 240, depth: 25 },
        '15x21': { width: 150, height: 210, depth: 20 },
        '16x23': { width: 160, height: 230, depth: 20 },
        '28x28': { width: 280, height: 280, depth: 25 },
        '30x30': { width: 300, height: 300, depth: 25 }
    };

    const LIMITS = {
        maxWidth: 420,
        maxHeight: 420,
        maxDepth: 100
    };

    // Dimensioni fogli macchina (in mm)
    const FOGLI_MACCHINA = {
        '70x100': { width: 700, height: 1000 },
        '64x88': { width: 640, height: 880 },
        '33x48_3': { width: 330, height: 483 },
        'a4': { width: 210, height: 297 },
        'a3': { width: 297, height: 420 },
        'super_a3': { width: 320, height: 450 },
        'banner_33x100': { width: 330, height: 1000 }
    };

    const OTTIMIZZAZIONI = {
        '70x100': [
            { width: 240, height: 330, label: '24x33 (16°)', type: 16 },
            { width: 170, height: 240, label: '17x24 (32°)', type: 32 }
        ],
        '64x88': [
            { width: 210, height: 297, label: '21x29.7 (16°)', type: 16 },
            { width: 150, height: 210, label: '15x21 (32°)', type: 32 }
        ],
        '33x48_3': [
            { width: 210, height: 297, label: '21x29.7 (4°)', type: 4 },
            { width: 150, height: 210, label: '15x21 (8°)', type: 8 }
        ]
    };

    function updateSegnature(openWidth, openHeight, sheet) {
        const isQuadrotto = orientamentoSelect && orientamentoSelect.value === 'quadrotto';
        const types = isQuadrotto ? [12, 24] : [4, 8, 16, 32];
        segnatureList.innerHTML = '';
        availableSignatureTypes = new Set();
        
        if (!sheet) return;

        types.forEach(type => {
            let possible = false;
            
            // Verifica semplificata di imposizione
            // 4 pag = 1 foglio aperto (fronte/retro) -> servono 1 aperti per lato
            // 8 pag = 2 fogli aperti -> servono 2 aperti per lato
            // 12 pag = 3 fogli aperti -> servono 3 aperti per lato
            // 16 pag = 4 fogli aperti -> servono 4 aperti per lato
            // 24 pag = 6 fogli aperti -> servono 6 aperti per lato
            // 32 pag = 8 fogli aperti -> servono 8 aperti per lato
            const numApertiPerLato = type / 4;

            // Possibili griglie per numApertiPerLato
            const grids = {
                1: [[1, 1]],
                2: [[2, 1], [1, 2]],
                3: [[3, 1], [1, 3]],
                4: [[2, 2], [4, 1], [1, 4]],
                6: [[3, 2], [2, 3], [6, 1], [1, 6]],
                8: [[4, 2], [2, 4], [8, 1], [1, 8]]
            };

            const possibleGrids = grids[numApertiPerLato];
            if (possibleGrids) {
                for (const [cols, rows] of possibleGrids) {
                    const totalW = openWidth * cols;
                    const totalH = openHeight * rows;
                    const totalW_rot = openWidth * rows;
                    const totalH_rot = openHeight * cols;

                    if ((totalW <= sheet.width && totalH <= sheet.height) ||
                        (totalW <= sheet.height && totalH <= sheet.width) ||
                        (totalW_rot <= sheet.width && totalH_rot <= sheet.height) ||
                        (totalW_rot <= sheet.height && totalH_rot <= sheet.width)) {
                        possible = true;
                        break;
                    }
                }
            }

            if (possible) {
                availableSignatureTypes.add(type);
            }

            const badge = document.createElement('div');
            badge.className = `signature-badge ${possible ? 'available' : 'unavailable'}`;
            badge.textContent = `${type}°`;
            segnatureList.appendChild(badge);
        });
    }

    function updateSuggerimenti(sheetKey) {
        ottimizzazioneContainer.style.display = 'none';
        if (!OTTIMIZZAZIONI[sheetKey]) return;

        const suggerimenti = OTTIMIZZAZIONI[sheetKey];
        suggerimentoOttimizzato.innerHTML = '';
        
        suggerimenti.forEach(sug => {
            const div = document.createElement('div');
            div.style.marginBottom = '5px';
            div.textContent = `Usa formato ${sug.label}`;
            suggerimentoOttimizzato.appendChild(div);
        });
        
        ottimizzazioneContainer.style.display = 'block';
    }

    let isUpdatingFromPreset = false; // Flag per evitare loop quando cambio preset

    function checkIfMatchesPreset() {
        const w = parseInt(customWidth.value) || 0;
        const h = parseInt(customHeight.value) || 0;
        const minVal = Math.min(w, h);
        const maxVal = Math.max(w, h);
        
        for (const [key, preset] of Object.entries(dimensioniPreset)) {
            const pMin = Math.min(preset.width, preset.height);
            const pMax = Math.max(preset.width, preset.height);
            if (pMin === minVal && pMax === maxVal) {
                return key;
            }
        }
        return null;
    }

    function updateFormaFromDimensions() {
        if (!orientamentoSelect) return;
        
        const w = parseInt(customWidth.value) || 0;
        const h = parseInt(customHeight.value) || 0;
        
        if (w === 0 || h === 0) return;
        
        let newForma = '';
        if (w === h) {
            newForma = 'quadrotto';
        } else {
            newForma = 'rettangolare';
        }

        if (orientamentoSelect.value !== newForma) {
            orientamentoSelect.value = newForma;
            updateFormatoOptions();
        }
    }

    function swapDimensions() {
        const w = customWidth.value;
        const h = customHeight.value;
        
        isUpdatingFromPreset = true;
        customWidth.value = h;
        customHeight.value = w;
        
        // Forza l'aggiornamento della forma in base alle nuove dimensioni
        updateFormaFromDimensions();
        
        // Dopo l'inversione, controlla se corrisponde ancora a un preset
        const matchingPreset = checkIfMatchesPreset();
        if (!matchingPreset) {
            formatoSelect.value = 'custom';
        } else {
            formatoSelect.value = matchingPreset;
        }
        
        isUpdatingFromPreset = false;
        updatePreview();
    }

    function validatePagine() {
        const raw = parseInt(numeroPagineInput.value, 10);

        pagineWarning.textContent = '';
        pagineWarning.classList.remove('warning-red');
        numeroPagineInput.style.borderColor = '#000000';

        if (!numeroPagineInput.value || numeroPagineInput.value.trim() === '') {
            return false;
        }

        if (!raw || raw < 4) {
            pagineWarning.textContent = 'Il numero minimo di pagine è 4.';
            pagineWarning.classList.add('warning-red');
            numeroPagineInput.style.borderColor = '#d8000c';
            return false;
        }

        if (raw > 1000) {
            pagineWarning.textContent = 'Il numero massimo di pagine è 1000.';
            pagineWarning.classList.add('warning-red');
            numeroPagineInput.style.borderColor = '#d8000c';
            return false;
        }

        if (raw % 4 !== 0) {
            pagineWarning.textContent = 'Il numero di pagine deve essere multiplo di 4.';
            pagineWarning.classList.add('warning-red');
            numeroPagineInput.style.borderColor = '#d8000c';
            return false;
        }

        return true;
    }

    function updateFoglioMacchina() {
        const tecnica = tecnicaSelect.value;
        let firstVisibleValue = '';

        Array.from(foglioSelect.options).forEach(option => {
            const tech = option.dataset.tecnica;
            // Se tecnica è vuoto, nascondi tutte le opzioni tranne quella vuota
            if (!tecnica) {
                option.hidden = option.value !== '';
            } else {
                const visible = tech === tecnica;
                option.hidden = !visible;

                if (visible && !firstVisibleValue) {
                    firstVisibleValue = option.value;
                }
            }
        });

        // Se l'opzione selezionata non è più valida, passa alla prima visibile o vuota
        const currentOption = foglioSelect.options[foglioSelect.selectedIndex];
        if (!currentOption || currentOption.hidden) {
            foglioSelect.value = firstVisibleValue || '';
        }
    }

    function updateFormatoOptions() {
        const orientamento = orientamentoSelect ? orientamentoSelect.value : '';
        const isQuadrotto = orientamento === 'quadrotto';
        
        Array.from(formatoSelect.options).forEach(option => {
            if (option.value === '') {
                // Mantieni sempre visibile l'opzione vuota
                option.hidden = false;
                return;
            }
            
            if (option.value === 'custom') {
                // Personalizzato sempre visibile
                option.hidden = false;
                return;
            }
            
            // Se è quadrotto, mostra solo 28x28 e 30x30
            if (isQuadrotto) {
                option.hidden = option.value !== '28x28' && option.value !== '30x30';
            } else {
                // Se non è quadrotto, mostra tutti i formati tranne quelli quadrotto
                option.hidden = option.value === '28x28' || option.value === '30x30';
            }
        });
        
        // Se l'opzione selezionata non è più valida, resetta
        const currentOption = formatoSelect.options[formatoSelect.selectedIndex];
        if (currentOption && currentOption.hidden) {
            formatoSelect.value = '';
        }
    }

    function getMicron(type, grammatura) {
        const table = PAPER_THICKNESS[type] || {};
        if (table[grammatura]) {
            return table[grammatura];
        }

        const grams = Object.keys(table).map(Number).sort((a, b) => a - b);
        if (grams.length > 0) {
            const nearest = grams.reduce((prev, curr) => (
                Math.abs(curr - grammatura) < Math.abs(prev - grammatura) ? curr : prev
            ), grams[0]);
            return table[nearest];
        }

        // fallback: se edizioni e non disponibile, usa usomano con moltiplicatore
        if (type === 'edizioni' && PAPER_THICKNESS.usomano[grammatura]) {
            return PAPER_THICKNESS.usomano[grammatura] * 1.6;
        }

        return PAPER_THICKNESS.patinata_opaca[100];
    }

    function buildPaperTypeOptions(selected) {
        return PAPER_TYPE_OPTIONS.map(opt => (
            `<option value="${opt.value}" ${opt.value === selected ? 'selected' : ''}>${opt.label}</option>`
        )).join('');
    }

    function buildGrammaturaOptions(selected) {
        return GRAMMATURA_OPTIONS.map(val => (
            `<option value="${val}" ${Number(selected) === val ? 'selected' : ''}>${val}</option>`
        )).join('');
    }

    function normalizePantoneEntry(entry) {
        if (typeof entry === 'string') {
            const name = entry.trim();
            if (!name) {
                return null;
            }
            return { name, color: '#000000' };
        }

        if (entry && typeof entry === 'object') {
            const rawName = typeof entry.name === 'string' ? entry.name : '';
            const name = rawName.trim();
            if (!name) {
                return null;
            }
            const color = typeof entry.color === 'string' && entry.color.trim()
                ? entry.color.trim()
                : '#000000';
            return { name, color };
        }

        return null;
    }

    function ensureSignatureColors(sig) {
        if (!sig.colors) {
            sig.colors = { c: true, m: true, y: true, k: true, pantone: [] };
        }
        if (!Array.isArray(sig.colors.pantone)) {
            sig.colors.pantone = [];
        }
        sig.colors.pantone = sig.colors.pantone
            .map(normalizePantoneEntry)
            .filter(Boolean);
    }

    function buildColorControls(sig, index) {
        if (!tecnicaSelect || tecnicaSelect.value !== 'offset') {
            return '';
        }

        ensureSignatureColors(sig);
        const colors = sig.colors;
        const pantoneList = colors.pantone.map((p, pantoneIndex) => (
            `<span class="sig-pantone-chip" data-index="${index}" data-pantone-index="${pantoneIndex}">
                <span class="sig-pantone-swatch" style="background-color: ${p.color};"></span>
                ${p.name}<button type="button" class="sig-pantone-remove" aria-label="Rimuovi ${p.name}">×</button>
            </span>`
        )).join('');

        return `
            <div class="sig-color-controls">
                <div class="sig-color-row">
                    <label class="sig-color-toggle">
                        <input type="checkbox" data-color="c" ${colors.c ? 'checked' : ''}>C
                    </label>
                    <label class="sig-color-toggle">
                        <input type="checkbox" data-color="m" ${colors.m ? 'checked' : ''}>M
                    </label>
                    <label class="sig-color-toggle">
                        <input type="checkbox" data-color="y" ${colors.y ? 'checked' : ''}>Y
                    </label>
                    <label class="sig-color-toggle">
                        <input type="checkbox" data-color="k" ${colors.k ? 'checked' : ''}>K
                    </label>
                </div>
                <div class="sig-pantone-row">
                    <input type="color" class="sig-pantone-color" value="#000000" aria-label="Colore Pantone">
                    <input type="text" class="sig-pantone-input" placeholder="Pantone (es. 485 C)">
                    <button type="button" class="sig-pantone-add">+ Pantone</button>
                </div>
                <div class="sig-pantone-list">${pantoneList}</div>
            </div>
        `;
    }

    function getSignatureThickness(sig, defaultType, defaultGram) {
        const pages = sig.pagine || 0;
        if (!pages) {
            return 0;
        }

        const type = sig.paperType || defaultType;
        const gram = sig.grammatura || defaultGram;
        const micron = getMicron(type, gram);
        const base = (pages / 2) * (micron / 1000);
        const collapse = COLLAPSE_FACTOR[type] || 1;
        return base * collapse;
    }

    function getSignatureColorCount(sig) {
        ensureSignatureColors(sig);
        const colors = sig.colors;
        const processCount = ['c', 'm', 'y', 'k'].filter(key => colors[key]).length;
        const pantoneCount = colors.pantone.length;
        return processCount + pantoneCount;
    }

    function buildSignatureStack(width, height, depths, scale) {
        if (!parallelepiped) {
            return;
        }

        parallelepiped.querySelectorAll('.signature-block').forEach(block => block.remove());

        const totalDepth = depths.reduce((sum, val) => sum + val, 0);
        const scaledTotalDepth = totalDepth * scale;
        let currentOffset = -scaledTotalDepth / 2;

        depths.forEach((depth, index) => {
            const scaledDepth = Math.max(depth * scale, 0.5);
            const halfD = scaledDepth / 2;

            // Alternanza tra grigio chiaro e grigio scuro basata sull'indice
            const isDark = index % 2 !== 0;
            const factor = isDark ? 0.88 : 1.0;
            
            const getColor = (r, g, b) => `rgb(${Math.round(r * factor)}, ${Math.round(g * factor)}, ${Math.round(b * factor)})`;

            const block = document.createElement('div');
            block.className = 'signature-block';
            block.style.width = `${width * scale}px`;
            block.style.height = `${height * scale}px`;
            block.style.transform = `translateZ(${currentOffset + halfD}px)`;

            const faces = {
                front: document.createElement('div'),
                back: document.createElement('div'),
                left: document.createElement('div'),
                right: document.createElement('div'),
                top: document.createElement('div'),
                bottom: document.createElement('div')
            };

            faces.front.className = 'face front';
            faces.back.className = 'face back';
            faces.left.className = 'face left';
            faces.right.className = 'face right';
            faces.top.className = 'face top';
            faces.bottom.className = 'face bottom';

            // Applichiamo i colori con le diverse tonalità per simulare la luce, ma scalati per blocco
            faces.front.style.backgroundColor = getColor(245, 245, 245);
            faces.back.style.backgroundColor = getColor(224, 224, 224);
            faces.left.style.backgroundColor = getColor(204, 204, 204);
            faces.right.style.backgroundColor = getColor(204, 204, 204);
            faces.top.style.backgroundColor = getColor(217, 217, 217);
            faces.bottom.style.backgroundColor = getColor(217, 217, 217);

            [faces.front, faces.back].forEach(face => {
                face.style.width = `${width * scale}px`;
                face.style.height = `${height * scale}px`;
            });
            faces.front.style.transform = `translateZ(${halfD}px)`;
            faces.back.style.transform = `rotateY(180deg) translateZ(${halfD}px)`;

            [faces.left, faces.right].forEach(face => {
                face.style.width = `${scaledDepth}px`;
                face.style.height = `${height * scale}px`;
                face.style.marginLeft = `-${halfD}px`;
            });
            faces.left.style.left = '0px';
            faces.left.style.transform = `rotateY(-90deg)`;
            faces.right.style.left = `${width * scale}px`;
            faces.right.style.transform = `rotateY(90deg)`;

            [faces.top, faces.bottom].forEach(face => {
                face.style.width = `${width * scale}px`;
                face.style.height = `${scaledDepth}px`;
                face.style.marginTop = `-${halfD}px`;
            });
            faces.top.style.top = '0px';
            faces.top.style.transform = `rotateX(90deg)`;
            faces.bottom.style.top = `${height * scale}px`;
            faces.bottom.style.transform = `rotateX(-90deg)`;

            block.appendChild(faces.front);
            block.appendChild(faces.back);
            block.appendChild(faces.left);
            block.appendChild(faces.right);
            block.appendChild(faces.top);
            block.appendChild(faces.bottom);

            parallelepiped.appendChild(block);
            currentOffset += scaledDepth;
        });
    }

    function calculateSpine() {
        let totalPagine = 0;
        const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
        const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
        const isOffset = tecnicaSelect && tecnicaSelect.value === 'offset';

        if (segnatureConfigurate.length > 0) {
            let totalThickness = 0;
            segnatureConfigurate.forEach(sig => {
                const pages = sig.pagine || 0;
                if (!pages) {
                    return;
                }

                totalPagine += pages;
                totalThickness += getSignatureThickness(sig, defaultType, defaultGram);
            });

            if (totalPagine === 0) {
                return 0;
            }

            return totalThickness + GLUE_MM;
        }

        totalPagine = parseInt(numeroPagineInput.value) || 0;
        if (totalPagine === 0) {
            return 0;
        }

        const micron = getMicron(defaultType, defaultGram);
        const base = (totalPagine / 2) * (micron / 1000);
        const collapse = COLLAPSE_FACTOR[defaultType] || 1;
        return (base * collapse) + GLUE_MM;
    }

    function checkGrammaturaForSignature(sigType, grammatura) {
        const gram = parseInt(grammatura, 10);
        if (sigType === 32) {
            return gram >= 70 && gram <= 115;
        } else if (sigType === 16) {
            return gram >= 130 && gram <= 150;
        } else if (sigType === 4 || sigType === 8) {
            return gram >= 170;
        }
        return true;
    }

    function getGrammaturaWarning(sigType, grammatura) {
        const gram = parseInt(grammatura, 10);
        if (sigType === 32) {
            if (gram < 70 || gram > 115) {
                return 'Grammatura consigliata per 32°: 70-115g';
            }
        } else if (sigType === 16) {
            if (gram < 130 || gram > 150) {
                return 'Grammatura consigliata per 16°: 130-150g';
            }
        } else if (sigType === 4 || sigType === 8) {
            if (gram < 170) {
                return 'Grammatura consigliata per 8°/4°: 170g+';
            }
        }
        return null;
    }

    function updateSignaturesUI() {
        signaturesConfigList.innerHTML = '';
        let totalPagineAssegnate = 0;
        const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
        const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
        const isOffset = tecnicaSelect && tecnicaSelect.value === 'offset';

        // Mostra/nascondi il pulsante reset in base al numero di segnature
        if (resetSignaturesBtn) {
            resetSignaturesBtn.style.display = segnatureConfigurate.length > 0 ? 'block' : 'none';
        }

        let hasInvalidSignature = false;
        let dropTargetIndex = null;

        const clearDragIndicator = () => {
            const indicator = signaturesConfigList.querySelector('.drag-indicator');
            if (indicator) {
                indicator.remove();
            }
        };

        const showDragIndicatorAt = (targetIndex) => {
            const items = Array.from(signaturesConfigList.querySelectorAll('.signature-item'));
            clearDragIndicator();

            const indicator = document.createElement('div');
            indicator.className = 'drag-indicator';

            if (items.length === 0 || targetIndex >= items.length) {
                signaturesConfigList.appendChild(indicator);
            } else {
                signaturesConfigList.insertBefore(indicator, items[targetIndex]);
            }
        };

        const getDropIndexFromPointer = (mouseY) => {
            const items = Array.from(signaturesConfigList.querySelectorAll('.signature-item'));
            for (let i = 0; i < items.length; i += 1) {
                const rect = items[i].getBoundingClientRect();
                const middle = rect.top + (rect.height / 2);
                if (mouseY < middle) {
                    return i;
                }
            }
            return items.length;
        };

        const moveSignature = (fromIndex, targetIndex) => {
            if (Number.isNaN(fromIndex) || fromIndex < 0 || fromIndex >= segnatureConfigurate.length) {
                return false;
            }

            let insertIndex = Math.max(0, Math.min(targetIndex, segnatureConfigurate.length));
            if (fromIndex < insertIndex) {
                insertIndex -= 1;
            }
            if (insertIndex === fromIndex) {
                return false;
            }

            const movedItem = segnatureConfigurate.splice(fromIndex, 1)[0];
            segnatureConfigurate.splice(insertIndex, 0, movedItem);
            return true;
        };

        segnatureConfigurate.forEach((sig, index) => {
            if (isOffset) {
                ensureSignatureColors(sig);
            }
            if (!sig.paperType) {
                sig.paperType = defaultType;
            }
            if (sig.grammatura == null || Number.isNaN(sig.grammatura)) {
                sig.grammatura = defaultGram;
            }

            const item = document.createElement('div');
            item.className = 'signature-item';
            item.setAttribute('data-type', sig.pagine || '');
            item.draggable = true;
            
            totalPagineAssegnate += sig.pagine;

            const gramWarning = getGrammaturaWarning(sig.pagine, sig.grammatura);
            const gramWarningHtml = gramWarning ? `<div class="sig-grammatura-warning">${gramWarning}</div>` : '';

            const isQuadrotto = orientamentoSelect && orientamentoSelect.value === 'quadrotto';
            const availableTypes = isQuadrotto ? [12, 24] : [4, 8, 16, 32];
            
            let typeOptions = '';
            availableTypes.forEach(type => {
                typeOptions += `<option value="${type}" ${sig.pagine === type ? 'selected' : ''}>${type}°</option>`;
            });
            
            const varianteSelectHtml = sig.pagine === 8 ? `
                <select class="sig-variante-select" data-index="${index}" style="font-size: 12px; padding: 5px; border: 1px solid #000000; font-family: 'Montreal Mono', monospace;">
                    <option value="" ${!sig.variante ? 'selected' : ''}>Standard</option>
                    <option value="finestra" ${sig.variante === 'finestra' ? 'selected' : ''}>A finestra</option>
                </select>
            ` : '';
            
            // Mostra il campo modalità di stampa solo se la tecnica è offset
            const stampaModeHtml = isOffset ? `
                <select class="sig-stampa-mode-select" data-index="${index}" style="font-size: 12px; padding: 5px; border: 1px solid #000000; font-family: 'Montreal Mono', monospace;">
                    <option value="normale" ${!sig.stampaMode || sig.stampaMode === 'normale' ? 'selected' : ''}>Bianca + Volta</option>
                    <option value="stesso-lato" ${sig.stampaMode === 'stesso-lato' ? 'selected' : ''}>Bianca + Volta su se stessa</option>
                </select>
            ` : '';
            
            item.innerHTML = `
                <div class="signature-header">
                    <span class="drag-handle">☰</span>
                    <span class="signature-label">Segnatura ${index + 1}</span>
                    <button class="remove-signature-btn">Elimina segnatura</button>
                </div>
                <div class="signature-content">
                    <select class="sig-type-select" data-index="${index}">
                        ${typeOptions}
                    </select>
                    ${varianteSelectHtml}
                    ${stampaModeHtml}
                    <select class="sig-paper-type" data-index="${index}">
                        ${buildPaperTypeOptions(sig.paperType)}
                    </select>
                    <input
                        type="number"
                        class="sig-grammatura"
                        data-index="${index}"
                        list="grammatura-options"
                        min="60"
                        max="400"
                        step="5"
                        value="${sig.grammatura || defaultGram}"
                    >
                    ${buildColorControls(sig, index)}
                    ${gramWarningHtml}
                </div>
            `;

            if (availableSignatureTypes.size > 0 && !availableSignatureTypes.has(sig.pagine)) {
                item.classList.add('invalid');
                hasInvalidSignature = true;
            }

            item.querySelector('.sig-type-select').onchange = (e) => {
                const newType = parseInt(e.target.value);
                segnatureConfigurate[index].pagine = newType;
                // Resetta la variante se non è più 8°
                if (newType !== 8) {
                    segnatureConfigurate[index].variante = null;
                }
                item.setAttribute('data-type', newType);
                updatePreview();
            };
            
            const varianteSelect = item.querySelector('.sig-variante-select');
            if (varianteSelect) {
                varianteSelect.onchange = (e) => {
                    segnatureConfigurate[index].variante = e.target.value || null;
                    updatePreview();
                };
            }

            const stampaModeSelect = item.querySelector('.sig-stampa-mode-select');
            if (stampaModeSelect) {
                stampaModeSelect.onchange = (e) => {
                    segnatureConfigurate[index].stampaMode = e.target.value || 'normale';
                    updatePreview();
                };
            }

            item.querySelector('.sig-paper-type').onchange = (e) => {
                segnatureConfigurate[index].paperType = e.target.value;
                updatePreview();
            };

            const grammaturaInput = item.querySelector('.sig-grammatura');
            grammaturaInput.addEventListener('input', (e) => {
                const raw = e.target.value.trim();
                segnatureConfigurate[index].grammatura = raw === '' ? null : parseInt(raw, 10);
                updatePreview({ skipSignatureRender: true });
            });
            grammaturaInput.addEventListener('blur', () => {
                updatePreview();
            });

            if (isOffset) {
                item.querySelectorAll('.sig-color-toggle input').forEach(input => {
                    input.addEventListener('change', (e) => {
                        const color = e.target.dataset.color;
                        sig.colors[color] = e.target.checked;
                        updatePreview();
                    });
                });

                const pantoneInput = item.querySelector('.sig-pantone-input');
                const pantoneColorInput = item.querySelector('.sig-pantone-color');
                const pantoneAdd = item.querySelector('.sig-pantone-add');
                if (pantoneAdd && pantoneInput && pantoneColorInput) {
                    pantoneAdd.addEventListener('click', () => {
                        const value = pantoneInput.value.trim();
                        const colorValue = pantoneColorInput.value || '#000000';
                        if (!value) {
                            return;
                        }

                        const exists = sig.colors.pantone.some(p => (
                            p.name.toLowerCase() === value.toLowerCase() &&
                            p.color.toLowerCase() === colorValue.toLowerCase()
                        ));
                        if (!exists) {
                            sig.colors.pantone.push({ name: value, color: colorValue });
                        }
                        pantoneInput.value = '';
                        pantoneColorInput.value = '#000000';
                        updatePreview();
                    });
                }

                item.querySelectorAll('.sig-pantone-remove').forEach(button => {
                    button.addEventListener('click', (e) => {
                        const chip = e.target.closest('.sig-pantone-chip');
                        const pantoneIndex = parseInt(chip?.dataset.pantoneIndex || '', 10);
                        if (Number.isNaN(pantoneIndex)) {
                            return;
                        }
                        sig.colors.pantone.splice(pantoneIndex, 1);
                        updatePreview();
                    });
                });
            }

            item.querySelector('.remove-signature-btn').onclick = () => {
                segnatureConfigurate.splice(index, 1);
                updatePreview();
            };

            item.ondragstart = (e) => {
                e.dataTransfer.setData('text/plain', index);
                e.dataTransfer.effectAllowed = 'move';
            };
            item.ondragend = () => {
                dropTargetIndex = null;
                clearDragIndicator();
            };
            item.ondragover = (e) => {
                e.preventDefault();
                dropTargetIndex = getDropIndexFromPointer(e.clientY);
                showDragIndicatorAt(dropTargetIndex);
            };
            item.ondrop = (e) => {
                e.preventDefault();
                e.stopPropagation();
                const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
                const targetIndex = dropTargetIndex == null ? index : dropTargetIndex;
                const moved = moveSignature(fromIndex, targetIndex);
                dropTargetIndex = null;
                clearDragIndicator();
                if (moved) {
                    updatePreview();
                }
            };

            signaturesConfigList.appendChild(item);
        });

        signaturesConfigList.ondragover = (e) => {
            e.preventDefault();
            dropTargetIndex = getDropIndexFromPointer(e.clientY);
            showDragIndicatorAt(dropTargetIndex);
        };
        signaturesConfigList.ondrop = (e) => {
            e.preventDefault();
            const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
            const targetIndex = dropTargetIndex == null
                ? signaturesConfigList.querySelectorAll('.signature-item').length
                : dropTargetIndex;
            const moved = moveSignature(fromIndex, targetIndex);
            dropTargetIndex = null;
            clearDragIndicator();
            if (moved) {
                updatePreview();
            }
        };
        signaturesConfigList.ondragleave = (e) => {
            if (!signaturesConfigList.contains(e.relatedTarget)) {
                dropTargetIndex = null;
                clearDragIndicator();
            }
        };

        const targetPagine = parseInt(numeroPagineInput.value) || 0;
        const residue = targetPagine - totalPagineAssegnate;
        
        pagineResidueDisplay.innerHTML = '';
        
        const statusDiv = document.createElement('div');
        statusDiv.className = `status-pagine ${residue === 0 ? 'success' : ''}`;
        statusDiv.textContent = `Pagine assegnate: ${totalPagineAssegnate} / ${targetPagine} (Residue: ${residue})`;
        pagineResidueDisplay.appendChild(statusDiv);

        if (hasInvalidSignature) {
            const warningDiv = document.createElement('div');
            warningDiv.className = 'warning-manuale';
            warningDiv.textContent = 'Alcune segnature non sono stampabili direttamente sul foglio selezionato: è necessario un accavallamento manuale.';
            pagineResidueDisplay.appendChild(warningDiv);
        }
    }

    function getMaxPrintableSignature() {
        if (!availableSignatureTypes || availableSignatureTypes.size === 0) {
            return null;
        }

        const printableTypes = Array.from(availableSignatureTypes)
            .filter(type => type > 0)
            .sort((a, b) => b - a);

        return printableTypes.length > 0 ? printableTypes[0] : null;
    }

    function calculateFogliMacchina() {
        const tiratura = parseInt(tiraturaInput?.value, 10) || 0;
        if (tiratura <= 0 || segnatureConfigurate.length === 0) {
            return {
                total: 0,
                base: 0,
                hasOffsetWaste: false
            };
        }

        const maxPrintableSignature = getMaxPrintableSignature();
        let fogliPerCopia = 0;

        segnatureConfigurate.forEach(sig => {
            const sigType = parseInt(sig.pagine, 10) || 0;
            if (sigType <= 0) {
                return;
            }

            // Regola base: una segnatura equivale a un foglio macchina.
            if (availableSignatureTypes.has(sigType)) {
                fogliPerCopia += 1;
                return;
            }

            // Se non stampabile, la segnatura viene suddivisa nelle segnature
            // massime stampabili: ogni blocco risultante vale un foglio macchina.
            if (maxPrintableSignature) {
                fogliPerCopia += Math.ceil(sigType / maxPrintableSignature);
                return;
            }

            // Senza riferimenti stampabili manteniamo una stima conservativa minima.
            fogliPerCopia += 1;
        });

        const base = fogliPerCopia * tiratura;
        const isOffset = tecnicaSelect && tecnicaSelect.value === 'offset';
        const total = isOffset ? Math.ceil(base * 1.15) : base;

        return {
            total,
            base,
            hasOffsetWaste: isOffset && base > 0
        };
    }

    function updateFogliMacchina() {
        if (!fogliMacchinaDisplay) return;

        const fogliData = calculateFogliMacchina();
        if (fogliData.hasOffsetWaste) {
            fogliMacchinaDisplay.textContent = `${fogliData.total.toLocaleString('it-IT')} (${fogliData.base.toLocaleString('it-IT')} + 15%)`;
        } else {
            fogliMacchinaDisplay.textContent = fogliData.total.toLocaleString('it-IT');
        }
    }

    function updatePreview(options = {}) {
        let dims;
        const selected = formatoSelect.value;

        // Prima di tutto, se non stiamo aggiornando da un preset, 
        // assicuriamoci che la forma sia sincronizzata con i valori attuali
        if (!isUpdatingFromPreset) {
            updateFormaFromDimensions();
        }

        // Mostra/nascondi il calcolo lastre in base alla tecnica
        if (platesGroup) {
            platesGroup.style.display = tecnicaSelect && tecnicaSelect.value === 'offset' ? 'block' : 'none';
        }

        // Calcola dorso dinamico
        const calculatedDepth = calculateSpine();
        if (spineDisplay) {
            spineDisplay.textContent = `${calculatedDepth.toFixed(2)} mm`;
        }
        if (platesDisplay) {
            platesDisplay.textContent = '';
        }

        // Se è un preset, popola i campi con i valori del preset in base alla forma
        if (selected !== 'custom' && dimensioniPreset[selected]) {
            const presetDims = dimensioniPreset[selected];
            const forma = orientamentoSelect ? orientamentoSelect.value : 'rettangolare';
            
            let targetW, targetH;
            
            if (forma === 'quadrotto') {
                targetW = Math.max(presetDims.width, presetDims.height);
                targetH = targetW;
            } else {
                // Per 'rettangolare', mantieni l'orientamento attuale dei campi input (o default verticale)
                const currentW = parseInt(customWidth.value) || 0;
                const currentH = parseInt(customHeight.value) || 0;
                
                if (currentW > currentH) {
                    // Orizzontale
                    targetW = Math.max(presetDims.width, presetDims.height);
                    targetH = Math.min(presetDims.width, presetDims.height);
                } else {
                    // Verticale (default)
                    targetW = Math.min(presetDims.width, presetDims.height);
                    targetH = Math.max(presetDims.width, presetDims.height);
                }
            }

            const currentW = parseInt(customWidth.value) || 0;
            const currentH = parseInt(customHeight.value) || 0;
            
            // Aggiorna solo se le dimensioni sono diverse o se sono vuote
            if (currentW !== targetW || currentH !== targetH || !currentW || !currentH) {
                const wasUpdating = isUpdatingFromPreset;
                isUpdatingFromPreset = true;
                customWidth.value = targetW;
                customHeight.value = targetH;
                customWidth.style.borderColor = '#000000';
                customHeight.style.borderColor = '#000000';
                widthError.style.display = 'none';
                heightError.style.display = 'none';
                isUpdatingFromPreset = wasUpdating;
            }
        }

        let w = parseInt(customWidth.value) || 0;
        let h = parseInt(customHeight.value) || 0;
        let d = calculatedDepth;

        // Se è quadrotto, forzare larghezza = altezza
        const isQuadrotto = orientamentoSelect && orientamentoSelect.value === 'quadrotto';
        if (isQuadrotto && w && h) {
            // Usa il valore più grande tra larghezza e altezza per entrambi
            const maxDim = Math.max(w, h);
            w = maxDim;
            h = maxDim;
            customWidth.value = maxDim;
            customHeight.value = maxDim;
        }

        if (w > LIMITS.maxWidth) {
            customWidth.style.borderColor = '#ff0000';
            widthError.style.display = 'block';
        } else {
            customWidth.style.borderColor = '#000000';
            widthError.style.display = 'none';
        }
        
        if (h > LIMITS.maxHeight) {
            customHeight.style.borderColor = '#ff0000';
            heightError.style.display = 'block';
        } else {
            customHeight.style.borderColor = '#000000';
            heightError.style.display = 'none';
        }

        if (w > LIMITS.maxWidth || h > LIMITS.maxHeight) {
            return;
        }

        // Se non ci sono dimensioni valide, mostra valori vuoti
        if (!w || !h || !selected) {
            if (formatoApertoDisplay) {
                formatoApertoDisplay.textContent = '-- × -- mm';
            }
            if (foglioWarning) {
                foglioWarning.textContent = '';
                foglioWarning.style.display = 'none';
            }
            if (spineDisplay) {
                spineDisplay.textContent = '0.0 mm';
            }
            if (platesDisplay) {
                platesDisplay.textContent = '';
            }
            if (segnatureList) {
                segnatureList.innerHTML = '';
            }
            if (suggerimentoOttimizzato) {
                suggerimentoOttimizzato.innerHTML = '';
            }
            if (ottimizzazioneContainer) {
                ottimizzazioneContainer.style.display = 'none';
            }
            buildSignatureStack([]);
            return;
        }

        dims = {
            width: w,
            height: h,
            depth: d
        };
        
        if (dims) {
            // Calcolo formato aperto
            const openWidth = dims.width * 2;
            const openHeight = dims.height;
            formatoApertoDisplay.textContent = `${openWidth} × ${openHeight} mm`;

            // Verifica compatibilità con il foglio macchina selezionato
            const foglioKey = foglioSelect.value;
            foglioWarning.textContent = '';
            foglioWarning.classList.remove('warning-yellow', 'warning-red');

            if (foglioKey && FOGLI_MACCHINA[foglioKey]) {
                const sheet = FOGLI_MACCHINA[foglioKey];
                
                // Aggiorna segnature e suggerimenti
                updateSegnature(openWidth, openHeight, sheet);
                updateSuggerimenti(foglioKey);

                const openFitsNormal =
                    openWidth <= sheet.width && openHeight <= sheet.height;
                const openFitsRotated =
                    openWidth <= sheet.height && openHeight <= sheet.width;

                const closedFitsNormal =
                    dims.width <= sheet.width && dims.height <= sheet.height;
                const closedFitsRotated =
                    dims.width <= sheet.height && dims.height <= sheet.width;

                const openFits = openFitsNormal || openFitsRotated;
                const closedFits = closedFitsNormal || closedFitsRotated;

                if (!openFits && closedFits) {
                    // Giallo: formato aperto non stampabile, ma pagina singola sì
                    foglioWarning.textContent =
                        'Attenzione: non stampabile in quartini, ma stampabile pagina singola a formato chiuso.';
                    foglioWarning.classList.add('warning-yellow');
                    foglioWarning.style.display = 'block';
                } else if (!openFits && !closedFits) {
                    // Rosso: nemmeno il formato chiuso entra nel foglio macchina
                    foglioWarning.textContent =
                        'Stampa impossibile: il formato chiuso è più grande del foglio macchina selezionato.';
                    foglioWarning.classList.add('warning-red');
                    foglioWarning.style.display = 'block';
                } else {
                    foglioWarning.style.display = 'none';
                }
            } else {
                availableSignatureTypes = new Set();
            }

            // Aggiorna UI segnature dopo aver calcolato disponibilità
            if (!options.skipSignatureRender) {
                updateSignaturesUI();
            }

        // Scala aumentata per rendere la preview più grande
        const scale = 1.5; 
        const w_preview = dims.width;
        const h_preview = dims.height;
        const d_preview = dims.depth;

        // Se le pagine superano 1000, non aggiornare il modello 3D
        const numPagine = parseInt(numeroPagineInput.value, 10);
        if (numPagine > 1000) {
            // Continuiamo comunque con la validazione testuale alla fine
        } else {
            parallelepiped.style.width = `${w_preview * scale}px`;
            parallelepiped.style.height = `${h_preview * scale}px`;

            const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
            const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
            const depths = [];

            if (segnatureConfigurate.length > 0) {
                segnatureConfigurate.forEach((sig) => {
                    const thickness = getSignatureThickness(sig, defaultType, defaultGram);
                    depths.push(thickness);
                });
                if (GLUE_MM > 0) {
                    depths.push(GLUE_MM);
                }
            } else {
                depths.push(d_preview);
            }

            buildSignatureStack(w_preview, h_preview, depths, scale);
            applyPreviewTransform();

            const halfTotalD = (depths.reduce((sum, val) => sum + val, 0) * scale) / 2;

            // Aggiorna etichette
            if (widthLabel) {
                widthLabel.textContent = `L: ${w_preview}mm`;
                widthLabel.style.transform = `translateX(-50%) translateZ(${halfTotalD + 5}px)`;
            }
            if (heightLabel) {
                heightLabel.textContent = `A: ${h_preview}mm`;
                heightLabel.style.transform = `translateY(-50%) rotate(-90deg) translateZ(${halfTotalD + 5}px)`;
            }
            if (depthLabel) {
                depthLabel.style.display = 'none';
            }

            if (platesDisplay) {
                if (tecnicaSelect && tecnicaSelect.value === 'offset') {
                    let totalPlates = 0;
                    let totalColors = 0;
                    segnatureConfigurate.forEach(sig => {
                        const colorCount = getSignatureColorCount(sig);
                        totalColors += colorCount;
                        const multiplier = sig.stampaMode === 'stesso-lato' ? 1 : 2;
                        totalPlates += colorCount * multiplier;
                    });
                    const normalPlates = totalColors * 2;
                    const savedPlates = normalPlates - totalPlates;
                    if (savedPlates > 0) {
                        platesDisplay.textContent = `Lastre totali: ${totalPlates} (${totalColors} colori × 2 fronte/retro, risparmio: ${savedPlates} lastre)`;
                    } else {
                        platesDisplay.textContent = `Lastre totali: ${totalPlates} (${totalColors} colori × 2 fronte/retro)`;
                    }
                } else {
                    platesDisplay.textContent = '';
                }
            }

            // Calcola fogli macchina necessari
            updateFogliMacchina();
        }
        }

        // Validazione pagine (non influisce sulla preview, solo avviso)
        validatePagine();

        // Se siamo in vista imposizione, aggiorna anche quella
        if (currentView === 'imposition') {
            updateImpositionView();
        }
    }
    
    if (orientamentoSelect) {
        orientamentoSelect.addEventListener('change', function() {
            const val = orientamentoSelect.value;
            const w = parseInt(customWidth.value) || 0;
            const h = parseInt(customHeight.value) || 0;

            if (w > 0 && h > 0) {
                isUpdatingFromPreset = true;
                if (val === 'quadrotto' && w !== h) {
                    const maxDim = Math.max(w, h);
                    customWidth.value = maxDim;
                    customHeight.value = maxDim;
                }
                isUpdatingFromPreset = false;
            }

            updateFormatoOptions();
            
            if (val === 'quadrotto') {
                // Converti automaticamente le segnature non valide per quadrotto
                const validTypes = [12, 24];
                segnatureConfigurate.forEach(sig => {
                    if (!validTypes.includes(sig.pagine)) {
                        const converted = validTypes.find(type => type <= sig.pagine) || validTypes[validTypes.length - 1];
                        sig.pagine = converted;
                    }
                });
            }
            updatePreview();
        });
    }

    formatoSelect.addEventListener('change', function() {
        isUpdatingFromPreset = true;
        updatePreview();
        isUpdatingFromPreset = false;
    });

    tecnicaSelect.addEventListener('change', function() {
        updateFoglioMacchina();
        // Mostra/nascondi il calcolo lastre in base alla tecnica
        if (platesGroup) {
            platesGroup.style.display = tecnicaSelect.value === 'offset' ? 'block' : 'none';
        }
        updatePreview();
    });

    foglioSelect.addEventListener('change', function() {
        updatePreview();
    });
    
    swapButton.addEventListener('click', swapDimensions);
    
    [customWidth, customHeight].forEach(input => {
        input.addEventListener('input', function() {
            if (isUpdatingFromPreset) return;

            // Aggiorna la forma in base alle dimensioni inserite
            updateFormaFromDimensions();
            
            // Se l'utente modifica manualmente i valori, controlla se corrisponde a un preset
            const matchingPreset = checkIfMatchesPreset();
            if (!matchingPreset && formatoSelect.value !== 'custom') {
                formatoSelect.value = 'custom';
            } else if (matchingPreset && formatoSelect.value !== matchingPreset) {
                formatoSelect.value = matchingPreset;
            }
            
            updatePreview();
        });
    });
    
    // Inizializzazione - campi lasciati vuoti
    updateFormatoOptions();
    
    // Nascondi il calcolo lastre all'avvio se la tecnica non è offset
    if (platesGroup) {
        platesGroup.style.display = tecnicaSelect && tecnicaSelect.value === 'offset' ? 'block' : 'none';
    }

    // Listener per numero pagine
    numeroPagineInput.addEventListener('input', updatePreview);

    // Listener per tiratura
    if (tiraturaInput) {
        tiraturaInput.addEventListener('input', function() {
            // Chiama updatePreview per assicurarsi che availableSignatureTypes sia aggiornato
            updatePreview();
        });
    }

    if (cartaTipoSelect) {
        cartaTipoSelect.addEventListener('change', updatePreview);
    }
    if (cartaGrammaturaSelect) {
        cartaGrammaturaSelect.addEventListener('change', updatePreview);
    }

    if (previewContainer) {
        let isDragging = false;
        let lastX = 0;
        let lastY = 0;

        previewContainer.addEventListener('mousedown', function(event) {
            isDragging = true;
            lastX = event.clientX;
            lastY = event.clientY;
            previewContainer.classList.add('dragging');
        });

        window.addEventListener('mousemove', function(event) {
            if (!isDragging) {
                return;
            }
            const dx = event.clientX - lastX;
            const dy = event.clientY - lastY;
            rotationY += dx * 0.3;
            rotationX -= dy * 0.3;
            rotationX = Math.max(-80, Math.min(80, rotationX));
            lastX = event.clientX;
            lastY = event.clientY;
            applyPreviewTransform();
        });

        window.addEventListener('mouseup', function() {
            isDragging = false;
            previewContainer.classList.remove('dragging');
        });

        previewContainer.addEventListener('mouseleave', function() {
            isDragging = false;
            previewContainer.classList.remove('dragging');
        });

        previewContainer.addEventListener('wheel', function(event) {
            event.preventDefault();
            const delta = -event.deltaY;
            zoom += delta * 0.001;
            zoom = Math.max(0.5, Math.min(2.5, zoom));
            applyPreviewTransform();
        }, { passive: false });

        // Supporto touch per mobile
        let touchStartDistance = 0;
        let touchStartZoom = 1;
        let touchStartX = 0;
        let touchStartY = 0;
        let touchStartRotationX = 0;
        let touchStartRotationY = 0;

        previewContainer.addEventListener('touchstart', function(event) {
            if (event.touches.length === 1) {
                // Un dito: prepara per rotazione
                touchStartX = event.touches[0].clientX;
                touchStartY = event.touches[0].clientY;
                touchStartRotationX = rotationX;
                touchStartRotationY = rotationY;
                previewContainer.classList.add('dragging');
            } else if (event.touches.length === 2) {
                // Due dita: prepara per zoom
                event.preventDefault();
                const touch1 = event.touches[0];
                const touch2 = event.touches[1];
                touchStartDistance = Math.hypot(
                    touch2.clientX - touch1.clientX,
                    touch2.clientY - touch1.clientY
                );
                touchStartZoom = zoom;
            }
        }, { passive: false });

        previewContainer.addEventListener('touchmove', function(event) {
            if (event.touches.length === 1) {
                // Un dito: ruota il modello
                event.preventDefault();
                const dx = event.touches[0].clientX - touchStartX;
                const dy = event.touches[0].clientY - touchStartY;
                rotationY = touchStartRotationY + dx * 0.3;
                rotationX = touchStartRotationX - dy * 0.3;
                rotationX = Math.max(-80, Math.min(80, rotationX));
                applyPreviewTransform();
            } else if (event.touches.length === 2) {
                // Due dita: zoom
                event.preventDefault();
                const touch1 = event.touches[0];
                const touch2 = event.touches[1];
                const currentDistance = Math.hypot(
                    touch2.clientX - touch1.clientX,
                    touch2.clientY - touch1.clientY
                );
                const scale = currentDistance / touchStartDistance;
                zoom = touchStartZoom * scale;
                zoom = Math.max(0.5, Math.min(2.5, zoom));
                applyPreviewTransform();
            }
        }, { passive: false });

        previewContainer.addEventListener('touchend', function(event) {
            previewContainer.classList.remove('dragging');
            touchStartDistance = 0;
        });
    }

    if (addSignatureBtn) {
        addSignatureBtn.addEventListener('click', function() {
            const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
            const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
            const isQuadrotto = orientamentoSelect && orientamentoSelect.value === 'quadrotto';
            const defaultPagine = isQuadrotto ? 12 : 16;
            
            segnatureConfigurate.push({
                pagine: defaultPagine,
                paperType: defaultType,
                grammatura: defaultGram,
                variante: null,
                stampaMode: 'normale'
            });
            updatePreview();
        });
    }

    if (autoSignaturesBtn) {
        autoSignaturesBtn.addEventListener('click', function() {
            const totalPages = parseInt(numeroPagineInput.value, 10) || 0;
            if (!totalPages || totalPages < 4) {
                return;
            }

            updatePreview();

            const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
            const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
            const isOffset = tecnicaSelect && tecnicaSelect.value === 'offset';

            const isQuadrotto = orientamentoSelect && orientamentoSelect.value === 'quadrotto';
            const available = Array.from(availableSignatureTypes);
            const defaultSizes = isQuadrotto ? [24, 12] : [32, 16, 8, 4];
            const sizes = available.length > 0 ? available.sort((a, b) => b - a) : defaultSizes;

            let remaining = totalPages;
            const generated = [];

            while (remaining > 0) {
                const size = sizes.find(val => val <= remaining);
                
                if (!size) {
                    // Se non c'è una segnatura che entra nelle pagine rimanenti, ferma il ciclo
                    // per evitare di aggiungere pagine extra
                    break;
                }
                
                generated.push({
                    pagine: size,
                    paperType: defaultType,
                    grammatura: defaultGram,
                    colors: isOffset ? { c: true, m: true, y: true, k: true, pantone: [] } : undefined,
                    variante: null,
                    stampaMode: 'normale'
                });
                remaining -= size;
            }

            // Riordina: segnature più piccole in penultima posizione, più grande in ultima
            if (generated.length > 1) {
                // Trova la segnatura più grande
                const maxSize = Math.max(...generated.map(s => s.pagine));
                const maxIndex = generated.findIndex(s => s.pagine === maxSize);
                const maxSignature = generated[maxIndex];
                
                // Trova le segnature più piccole (valore minimo)
                const minSize = Math.min(...generated.map(s => s.pagine));
                const smallSignatures = generated.filter(s => s.pagine === minSize);
                
                // Rimuovi la più grande e le più piccole dall'array originale
                const others = generated.filter((s, i) => 
                    i !== maxIndex && s.pagine !== minSize
                );
                
                // Ricostruisci l'array: altre segnature, poi le più piccole, poi la più grande
                segnatureConfigurate = [...others, ...smallSignatures, maxSignature];
            } else {
                segnatureConfigurate = generated;
            }
            
            updatePreview();
        });
    }

    if (resetSignaturesBtn) {
        resetSignaturesBtn.addEventListener('click', function() {
            segnatureConfigurate = [];
            updatePreview();
        });
    }

    if (resetProjectBtn) {
        resetProjectBtn.addEventListener('click', function() {
            if (nomeProgettoInput) nomeProgettoInput.value = '';
            formatoSelect.value = '';
            if (orientamentoSelect) orientamentoSelect.value = '';
            tecnicaSelect.value = '';
            updateFoglioMacchina();
            updateFormatoOptions();

            customWidth.value = '';
            customHeight.value = '';
            numeroPagineInput.value = '';
            if (cartaTipoSelect) cartaTipoSelect.value = '';
            if (cartaGrammaturaSelect) cartaGrammaturaSelect.value = '';
            if (tiraturaInput) tiraturaInput.value = '';

            segnatureConfigurate = [];
            availableSignatureTypes = new Set();

            updatePreview();
        });
    }

    function generatePDFReport() {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        let yPos = 20;
        const pageWidth = doc.internal.pageSize.getWidth();
        const margin = 20;
        const now = new Date();
        
        const nomeProgetto = nomeProgettoInput ? nomeProgettoInput.value.trim() : '';
        const projectTitle = nomeProgetto || 'Untitled';
        const fileName = `Report Progetto ${projectTitle}.pdf`;

        function addSectionTitle(text) {
            if (yPos > 250) {
                doc.addPage();
                yPos = 20;
            }
            doc.setFontSize(14);
            doc.setFont('helvetica', 'bold');
            doc.text(text, margin, yPos);
            yPos += 10;
            doc.setFontSize(10);
            doc.setFont('helvetica', 'normal');
        }

        function addText(text) {
            doc.text(text, margin, yPos);
            yPos += 7;
        }

        function savePDF() {
            doc.save(fileName);
        }

        doc.setFontSize(20);
        doc.setFont('helvetica', 'bold');
        doc.text('REPORT PROGETTO', pageWidth / 2, yPos, { align: 'center' });
        yPos += 15;

        if (nomeProgetto) {
            doc.setFontSize(14);
            doc.setFont('helvetica', 'bold');
            doc.text(nomeProgetto, pageWidth / 2, yPos, { align: 'center' });
            yPos += 10;
        }

        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        const dateStr = now.toLocaleDateString('it-IT', { 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
        doc.text(`Generato il: ${dateStr}`, pageWidth / 2, yPos, { align: 'center' });
        yPos += 20;

        addSectionTitle('DATI GENERALI');

        const tecnica = tecnicaSelect.value || 'Non specificato';
        const tecnicaLabel = tecnica === 'offset' ? 'Offset' : tecnica === 'digitale' ? 'Digitale' : tecnica;
        addText(`Tecnica di stampa: ${tecnicaLabel}`);

        const foglioLabel = foglioSelect.options[foglioSelect.selectedIndex]?.text || 'Non specificato';
        addText(`Foglio macchina: ${foglioLabel}`);

        const orientamento = orientamentoSelect ? orientamentoSelect.value : '';
        const orientamentoLabel = orientamento === 'quadrotto' ? 'Quadrotto' : orientamento === 'rettangolare' ? 'Rettangolare' : 'Non specificato';
        addText(`Forma: ${orientamentoLabel}`);

        const formatoLabel = formatoSelect.options[formatoSelect.selectedIndex]?.text || formatoSelect.value || 'Non specificato';
        addText(`Formato (chiuso): ${formatoLabel}`);

        const width = parseInt(customWidth.value) || 0;
        const height = parseInt(customHeight.value) || 0;
        if (width && height) {
            addText(`Dimensioni (chiuso): ${width} × ${height} mm`);
            addText(`Formato aperto: ${width * 2} × ${height} mm`);
        }

        const numPagine = parseInt(numeroPagineInput.value) || 0;
        addText(`Numero pagine: ${numPagine || 'Non specificato'}`);

        const cartaTipo = cartaTipoSelect ? cartaTipoSelect.value : '';
        const cartaTipoLabel = cartaTipoSelect ? cartaTipoSelect.options[cartaTipoSelect.selectedIndex]?.text || cartaTipo : 'Non specificato';
        addText(`Tipo carta: ${cartaTipoLabel}`);

        const grammatura = cartaGrammaturaSelect ? cartaGrammaturaSelect.value : '';
        addText(`Grammatura: ${grammatura ? grammatura + ' g/m²' : 'Non specificato'}`);

        const spine = spineDisplay ? spineDisplay.textContent : '0.0 mm';
        addText(`Dorso stimato: ${spine}`);

        if (platesDisplay && platesDisplay.textContent) {
            addText(`Lastre: ${platesDisplay.textContent}`);
        }

        const tiratura = tiraturaInput ? parseInt(tiraturaInput.value, 10) : 0;
        if (tiratura > 0) {
            addText(`Tiratura: ${tiratura.toLocaleString('it-IT')}`);
            
            const fogliData = calculateFogliMacchina();
            if (fogliData.hasOffsetWaste) {
                addText(`Fogli macchina necessari: ${fogliData.total.toLocaleString('it-IT')} (${fogliData.base.toLocaleString('it-IT')} + 15%)`);
            } else {
                addText(`Fogli macchina necessari: ${fogliData.total.toLocaleString('it-IT')}`);
            }
        }

        yPos += 10;

        if (segnatureConfigurate.length > 0) {
            addSectionTitle('SEGNATURE');

            let totalPagineAssegnate = 0;
            segnatureConfigurate.forEach((sig, index) => {
                if (yPos > 270) {
                    doc.addPage();
                    yPos = 20;
                }

                const sigType = sig.pagine || 0;
                const sigPaperType = sig.paperType || cartaTipo || 'Non specificato';
                const sigGrammatura = sig.grammatura || grammatura || 'Non specificato';
                
                doc.setFont('helvetica', 'bold');
                doc.text(`Segnatura ${index + 1}:`, margin, yPos);
                yPos += 6;
                
                doc.setFont('helvetica', 'normal');
                doc.text(`  Tipo: ${sigType}°`, margin + 5, yPos);
                yPos += 6;
                doc.text(`  Carta: ${sigPaperType}`, margin + 5, yPos);
                yPos += 6;
                doc.text(`  Grammatura: ${sigGrammatura} g/m²`, margin + 5, yPos);
                yPos += 6;

                if (tecnica === 'offset' && sig.colors) {
                    ensureSignatureColors(sig);
                    const processColors = [];
                    if (sig.colors.c) processColors.push('C');
                    if (sig.colors.m) processColors.push('M');
                    if (sig.colors.y) processColors.push('Y');
                    if (sig.colors.k) processColors.push('K');
                    const colorsText = processColors.length > 0 ? processColors.join('+') : 'Nessuno';
                    const pantoneList = sig.colors.pantone && sig.colors.pantone.length > 0
                        ? ` + Pantone: ${sig.colors.pantone.map(p => `${p.name} (${p.color})`).join(', ')}`
                        : '';
                    doc.text(`  Colori: ${colorsText}${pantoneList}`, margin + 5, yPos);
                    yPos += 6;
                }

                totalPagineAssegnate += sigType;
                yPos += 3;
            });

            doc.text(`Totale pagine assegnate: ${totalPagineAssegnate} / ${numPagine}`, margin, yPos);
            yPos += 10;
        }

        savePDF();
    }

    const downloadReportBtn = document.getElementById('download-report-btn');
    if (downloadReportBtn) {
        downloadReportBtn.addEventListener('click', generatePDFReport);
    }

    // Inizializza la vista imposition come predefinita
    setPreviewView('imposition');

    function saveProject() {
        const projectData = {
            version: '1.0',
            nomeProgetto: nomeProgettoInput ? nomeProgettoInput.value : '',
            tecnicaStampa: tecnicaSelect ? tecnicaSelect.value : '',
            foglioMacchina: foglioSelect ? foglioSelect.value : '',
            orientamento: orientamentoSelect ? orientamentoSelect.value : '',
            formato: formatoSelect ? formatoSelect.value : '',
            customWidth: customWidth ? customWidth.value : '',
            customHeight: customHeight ? customHeight.value : '',
            numeroPagine: numeroPagineInput ? numeroPagineInput.value : '',
            cartaTipo: cartaTipoSelect ? cartaTipoSelect.value : '',
            cartaGrammatura: cartaGrammaturaSelect ? cartaGrammaturaSelect.value : '',
            tiratura: tiraturaInput ? tiraturaInput.value : '',
            segnatureConfigurate: segnatureConfigurate.map(sig => ({
                pagine: sig.pagine,
                variante: sig.variante || null,
                stampaMode: sig.stampaMode || null,
                paperType: sig.paperType || null,
                grammatura: sig.grammatura || null,
                colors: sig.colors ? {
                    c: sig.colors.c || false,
                    m: sig.colors.m || false,
                    y: sig.colors.y || false,
                    k: sig.colors.k || false,
                    pantone: sig.colors.pantone ? sig.colors.pantone.map(p => 
                        typeof p === 'string' ? { name: p, color: '#000000' } : p
                    ) : []
                } : null
            }))
        };

        const jsonStr = JSON.stringify(projectData, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const nomeProgetto = nomeProgettoInput ? nomeProgettoInput.value.trim() : 'progetto';
        a.download = `${nomeProgetto || 'progetto'}_book-designer.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function loadProject(file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const projectData = JSON.parse(e.target.result);
                
                // Ripristina i campi base
                if (nomeProgettoInput && projectData.nomeProgetto) {
                    nomeProgettoInput.value = projectData.nomeProgetto;
                }
                if (tecnicaSelect && projectData.tecnicaStampa) {
                    tecnicaSelect.value = projectData.tecnicaStampa;
                }
                if (foglioSelect && projectData.foglioMacchina) {
                    foglioSelect.value = projectData.foglioMacchina;
                }
                if (orientamentoSelect && projectData.orientamento) {
                    orientamentoSelect.value = projectData.orientamento;
                }
                if (formatoSelect && projectData.formato) {
                    formatoSelect.value = projectData.formato;
                }
                if (customWidth && projectData.customWidth) {
                    customWidth.value = projectData.customWidth;
                }
                if (customHeight && projectData.customHeight) {
                    customHeight.value = projectData.customHeight;
                }
                if (numeroPagineInput && projectData.numeroPagine) {
                    numeroPagineInput.value = projectData.numeroPagine;
                }
                if (cartaTipoSelect && projectData.cartaTipo) {
                    cartaTipoSelect.value = projectData.cartaTipo;
                }
                if (cartaGrammaturaSelect && projectData.cartaGrammatura) {
                    cartaGrammaturaSelect.value = projectData.cartaGrammatura;
                }
                if (tiraturaInput && projectData.tiratura) {
                    tiraturaInput.value = projectData.tiratura;
                }

                // Ripristina le segnature
                if (projectData.segnatureConfigurate && Array.isArray(projectData.segnatureConfigurate)) {
                    segnatureConfigurate = projectData.segnatureConfigurate.map(sig => {
                        const restored = {
                            pagine: sig.pagine,
                            variante: sig.variante || null,
                            stampaMode: sig.stampaMode || null,
                            paperType: sig.paperType || null,
                            grammatura: sig.grammatura || null
                        };
                        if (sig.colors) {
                            restored.colors = {
                                c: sig.colors.c || false,
                                m: sig.colors.m || false,
                                y: sig.colors.y || false,
                                k: sig.colors.k || false,
                                pantone: sig.colors.pantone ? sig.colors.pantone.map(p => 
                                    typeof p === 'string' ? { name: p, color: '#000000' } : p
                                ) : []
                            };
                        }
                        return restored;
                    });
                } else {
                    segnatureConfigurate = [];
                }

                // Aggiorna le opzioni del foglio e triggera gli eventi necessari
                if (tecnicaSelect) {
                    tecnicaSelect.dispatchEvent(new Event('change'));
                }
                if (foglioSelect) {
                    foglioSelect.dispatchEvent(new Event('change'));
                }
                if (orientamentoSelect) {
                    orientamentoSelect.dispatchEvent(new Event('change'));
                }
                if (formatoSelect) {
                    formatoSelect.dispatchEvent(new Event('change'));
                }

                // Aggiorna la preview
                updatePreview();
                
                alert('Progetto caricato con successo!');
            } catch (error) {
                console.error('Errore nel caricamento del progetto:', error);
                alert('Errore nel caricamento del file. Assicurati che sia un file valido.');
            }
        };
        reader.readAsText(file);
    }

    const saveProjectBtn = document.getElementById('save-project-btn');
    if (saveProjectBtn) {
        saveProjectBtn.addEventListener('click', saveProject);
    }

    const loadProjectBtn = document.getElementById('load-project-btn');
    const loadProjectInput = document.getElementById('load-project-input');
    if (loadProjectBtn && loadProjectInput) {
        loadProjectBtn.addEventListener('click', function() {
            loadProjectInput.click();
        });
        loadProjectInput.addEventListener('change', function(e) {
            if (e.target.files && e.target.files.length > 0) {
                loadProject(e.target.files[0]);
                e.target.value = ''; // Reset per permettere di ricaricare lo stesso file
            }
        });
    }
});
