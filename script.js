// Book Designer - Script principale

document.addEventListener('DOMContentLoaded', function() {
    const formatoSelect = document.getElementById('formato');
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
    const ottimizzazioneContainer = document.getElementById('ottimizzazione-container');
    const suggerimentoOttimizzato = document.getElementById('suggerimento-ottimizzato');
    const signaturesConfigList = document.getElementById('signatures-config-list');
    const addSignatureBtn = document.getElementById('add-signature-btn');
    const autoSignaturesBtn = document.getElementById('auto-signatures-btn');
    const pagineResidueDisplay = document.getElementById('pagine-residue');
    const cartaTipoSelect = document.getElementById('carta-tipo');
    const cartaGrammaturaSelect = document.getElementById('carta-grammatura');
    const resetProjectBtn = document.getElementById('reset-project');

    let segnatureConfigurate = [];
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
        parallelepiped.style.transform = `rotateX(${rotationX}deg) rotateY(${rotationY}deg) scale(${zoom})`;
    }

    const dimensioniPreset = {
        'A4': { width: 210, height: 297, depth: 30 },
        'A5': { width: 148, height: 210, depth: 25 },
        'B5': { width: 176, height: 250, depth: 25 },
        '17x24': { width: 170, height: 240, depth: 25 },
        '15x21': { width: 150, height: 210, depth: 20 },
        '16x23': { width: 160, height: 230, depth: 20 }
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
        const types = [4, 8, 16, 32];
        segnatureList.innerHTML = '';
        availableSignatureTypes = new Set();
        
        if (!sheet) return;

        types.forEach(type => {
            let possible = false;
            
            // Verifica semplificata di imposizione
            // 4 pag = 1 foglio aperto (fronte/retro) -> servono 1 aperti per lato
            // 8 pag = 2 fogli aperti -> servono 2 aperti per lato
            // 16 pag = 4 fogli aperti -> servono 4 aperti per lato
            // 32 pag = 8 fogli aperti -> servono 8 aperti per lato
            const numApertiPerLato = type / 4;

            // Possibili griglie per numApertiPerLato (es. per 4 aperti: 2x2, 4x1)
            const grids = {
                1: [[1, 1]],
                2: [[2, 1], [1, 2]],
                4: [[2, 2], [4, 1], [1, 4]],
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
        
        for (const [key, preset] of Object.entries(dimensioniPreset)) {
            if (preset.width === w && preset.height === h) {
                return key;
            }
        }
        return null;
    }

    function swapDimensions() {
        const temp = customWidth.value;
        customWidth.value = customHeight.value;
        customHeight.value = temp;
        
        // Dopo l'inversione, controlla se corrisponde ancora a un preset
        const matchingPreset = checkIfMatchesPreset();
        if (!matchingPreset) {
            formatoSelect.value = 'custom';
        }
        
        updatePreview();
    }

    function validatePagine() {
        const raw = parseInt(numeroPagineInput.value, 10);

        pagineWarning.textContent = '';
        pagineWarning.classList.remove('warning-red');
        numeroPagineInput.style.borderColor = '#000000';

        if (!raw || raw < 4) {
            pagineWarning.textContent = 'Il numero minimo di pagine è 4.';
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
            const visible = tech === tecnica;
            option.hidden = !visible;

            if (visible && !firstVisibleValue) {
                firstVisibleValue = option.value;
            }
        });

        // Se l'opzione selezionata non è più valida, passa alla prima visibile
        const currentOption = foglioSelect.options[foglioSelect.selectedIndex];
        if (!currentOption || currentOption.hidden) {
            foglioSelect.value = firstVisibleValue || '';
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

    function ensureSignatureColors(sig) {
        if (!sig.colors) {
            sig.colors = { c: true, m: true, y: true, k: true, pantone: [] };
        }
        if (!Array.isArray(sig.colors.pantone)) {
            sig.colors.pantone = [];
        }
    }

    function buildColorControls(sig, index) {
        if (!tecnicaSelect || tecnicaSelect.value !== 'offset') {
            return '';
        }

        ensureSignatureColors(sig);
        const colors = sig.colors;
        const pantoneList = colors.pantone.map(p => (
            `<span class="sig-pantone-chip" data-index="${index}" data-pantone="${p}">
                ${p}<button type="button" class="sig-pantone-remove" aria-label="Rimuovi ${p}">×</button>
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

        depths.forEach(depth => {
            const scaledDepth = Math.max(depth * scale, 0.5);
            const halfD = scaledDepth / 2;

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

        let hasInvalidSignature = false;

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
            item.draggable = true;
            
            totalPagineAssegnate += sig.pagine;

            const gramWarning = getGrammaturaWarning(sig.pagine, sig.grammatura);
            const gramWarningHtml = gramWarning ? `<div class="sig-grammatura-warning">${gramWarning}</div>` : '';

            item.innerHTML = `
                <span class="drag-handle">☰</span>
                <select class="sig-type-select" data-index="${index}">
                    <option value="4" ${sig.pagine === 4 ? 'selected' : ''}>4°</option>
                    <option value="8" ${sig.pagine === 8 ? 'selected' : ''}>8°</option>
                    <option value="12" ${sig.pagine === 12 ? 'selected' : ''}>12°</option>
                    <option value="16" ${sig.pagine === 16 ? 'selected' : ''}>16°</option>
                    <option value="24" ${sig.pagine === 24 ? 'selected' : ''}>24°</option>
                    <option value="32" ${sig.pagine === 32 ? 'selected' : ''}>32°</option>
                </select>
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
                <button class="remove-btn">✕</button>
                ${buildColorControls(sig, index)}
                ${gramWarningHtml}
            `;

            if (availableSignatureTypes.size > 0 && !availableSignatureTypes.has(sig.pagine)) {
                item.classList.add('invalid');
                hasInvalidSignature = true;
            }

            item.querySelector('.sig-type-select').onchange = (e) => {
                segnatureConfigurate[index].pagine = parseInt(e.target.value);
                updatePreview();
            };

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
                const pantoneAdd = item.querySelector('.sig-pantone-add');
                if (pantoneAdd && pantoneInput) {
                    pantoneAdd.addEventListener('click', () => {
                        const value = pantoneInput.value.trim();
                        if (!value) {
                            return;
                        }
                        if (!sig.colors.pantone.includes(value)) {
                            sig.colors.pantone.push(value);
                        }
                        pantoneInput.value = '';
                        updatePreview();
                    });
                }

                item.querySelectorAll('.sig-pantone-remove').forEach(button => {
                    button.addEventListener('click', (e) => {
                        const chip = e.target.closest('.sig-pantone-chip');
                        const pantone = chip?.dataset.pantone;
                        if (!pantone) {
                            return;
                        }
                        sig.colors.pantone = sig.colors.pantone.filter(p => p !== pantone);
                        updatePreview();
                    });
                });
            }

            item.querySelector('.remove-btn').onclick = () => {
                segnatureConfigurate.splice(index, 1);
                updatePreview();
            };

            item.ondragstart = (e) => {
                e.dataTransfer.setData('text/plain', index);
            };
            item.ondragover = (e) => e.preventDefault();
            item.ondrop = (e) => {
                const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
                const toIndex = index;
                const movedItem = segnatureConfigurate.splice(fromIndex, 1)[0];
                segnatureConfigurate.splice(toIndex, 0, movedItem);
                updatePreview();
            };

            signaturesConfigList.appendChild(item);
        });

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

    function updatePreview(options = {}) {
        let dims;
        const selected = formatoSelect.value;

        // Calcola dorso dinamico
        const calculatedDepth = calculateSpine();
        if (spineDisplay) {
            spineDisplay.textContent = `${calculatedDepth.toFixed(2)} mm`;
        }
        if (platesDisplay) {
            platesDisplay.textContent = '';
        }

        // Se è un preset, popola i campi con i valori del preset
        if (selected !== 'custom' && dimensioniPreset[selected] && !isUpdatingFromPreset) {
            const currentW = parseInt(customWidth.value) || 0;
            const currentH = parseInt(customHeight.value) || 0;
            
            const presetDims = dimensioniPreset[selected];
            if ((currentW !== presetDims.width || currentH !== presetDims.height) &&
                currentW <= LIMITS.maxWidth && currentH <= LIMITS.maxHeight) {
                isUpdatingFromPreset = true;
                customWidth.value = presetDims.width;
                customHeight.value = presetDims.height;
                customWidth.style.borderColor = '#000000';
                customHeight.style.borderColor = '#000000';
                widthError.style.display = 'none';
                heightError.style.display = 'none';
                isUpdatingFromPreset = false;
            }
        }

        let w = parseInt(customWidth.value) || 0;
        let h = parseInt(customHeight.value) || 0;
        let d = calculatedDepth;

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
        const w = dims.width;
        const h = dims.height;
        const d = dims.depth;

        parallelepiped.style.width = `${w * scale}px`;
        parallelepiped.style.height = `${h * scale}px`;

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
            depths.push(d);
        }

        buildSignatureStack(w, h, depths, scale);
        applyPreviewTransform();

        const halfTotalD = (depths.reduce((sum, val) => sum + val, 0) * scale) / 2;

        // Aggiorna etichette
        if (widthLabel) {
            widthLabel.textContent = `L: ${w}mm`;
            widthLabel.style.transform = `translateX(-50%) translateZ(${halfTotalD + 5}px)`;
        }
        if (heightLabel) {
            heightLabel.textContent = `A: ${h}mm`;
            heightLabel.style.transform = `translateY(-50%) rotate(-90deg) translateZ(${halfTotalD + 5}px)`;
        }
        if (depthLabel) {
            depthLabel.style.display = 'none';
        }

        if (platesDisplay) {
            if (tecnicaSelect && tecnicaSelect.value === 'offset') {
                let totalColors = 0;
                segnatureConfigurate.forEach(sig => {
                    totalColors += getSignatureColorCount(sig);
                });
                const totalPlates = totalColors * 2;
                platesDisplay.textContent = `Lastre totali: ${totalPlates} (${totalColors} colori × 2 fronte/retro)`;
            } else {
                platesDisplay.textContent = '';
            }
        }
        }

        // Validazione pagine (non influisce sulla preview, solo avviso)
        validatePagine();
    }
    
    formatoSelect.addEventListener('change', function() {
        isUpdatingFromPreset = true;
        updatePreview();
        isUpdatingFromPreset = false;
    });

    tecnicaSelect.addEventListener('change', function() {
        updateFoglioMacchina();
        updatePreview();
    });

    foglioSelect.addEventListener('change', function() {
        updatePreview();
    });
    
    swapButton.addEventListener('click', swapDimensions);
    
    [customWidth, customHeight].forEach(input => {
        input.addEventListener('input', function() {
            // Se l'utente modifica manualmente i valori, controlla se corrisponde a un preset
            if (!isUpdatingFromPreset) {
                const matchingPreset = checkIfMatchesPreset();
                if (!matchingPreset && formatoSelect.value !== 'custom') {
                    formatoSelect.value = 'custom';
                }
            }
            updatePreview();
        });
    });
    
    // Inizializzazione
    updateFoglioMacchina();
    updatePreview();

    // Listener per numero pagine
    numeroPagineInput.addEventListener('input', updatePreview);

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
    }

    if (addSignatureBtn) {
        addSignatureBtn.addEventListener('click', function() {
            const defaultType = cartaTipoSelect ? cartaTipoSelect.value : 'patinata_opaca';
            const defaultGram = parseInt(cartaGrammaturaSelect?.value, 10) || 100;
            segnatureConfigurate.push({
                pagine: 16,
                paperType: defaultType,
                grammatura: defaultGram
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

            const available = Array.from(availableSignatureTypes);
            const sizes = available.length > 0 ? available.sort((a, b) => b - a) : [32, 16, 8, 4];

            let remaining = totalPages;
            const generated = [];

            while (remaining > 0) {
                const size = sizes.find(val => val <= remaining) || 4;
                generated.push({
                    pagine: size,
                    paperType: defaultType,
                    grammatura: defaultGram,
                    colors: isOffset ? { c: true, m: true, y: true, k: true, pantone: [] } : undefined
                });
                remaining -= size;
            }

            segnatureConfigurate = generated;
            updatePreview();
        });
    }

    if (resetProjectBtn) {
        resetProjectBtn.addEventListener('click', function() {
            formatoSelect.value = 'A4';
            tecnicaSelect.value = 'offset';
            updateFoglioMacchina();

            customWidth.value = 210;
            customHeight.value = 297;
            numeroPagineInput.value = 100;
            cartaTipoSelect.value = 'patinata_opaca';
            cartaGrammaturaSelect.value = 100;

            segnatureConfigurate = [];
            availableSignatureTypes = new Set();

            updatePreview();
        });
    }
});
