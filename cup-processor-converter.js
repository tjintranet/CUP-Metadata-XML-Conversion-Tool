// Excel to XML Converter Module
class ConverterModule {
    constructor() {
        this.workbookData = [];
        this.headers = [];
        this.mappedHeaders = [];
        this.generatedXMLs = []; // Store generated XMLs for validation
        
        this.initializeElements();
        this.attachEventListeners();
    }
    
    initializeElements() {
        this.excelFileInput = document.getElementById('excelFile');
        this.downloadAllBtn = document.getElementById('downloadAllBtn');
        this.validateGeneratedBtn = document.getElementById('validateGeneratedBtn');
        this.clearAllBtn = document.getElementById('clearAllBtn');
        this.processingIndicator = document.getElementById('processingIndicator');
        this.processingText = document.getElementById('processingText');
        this.statusAlert = document.getElementById('statusAlert');
        this.summarySection = document.getElementById('summarySection');
        this.totalRecords = document.getElementById('totalRecords');
        this.bindingSummary = document.getElementById('bindingSummary');
        this.colourSummary = document.getElementById('colourSummary');
        this.qualitySummary = document.getElementById('qualitySummary');
    }
    
    attachEventListeners() {
        this.excelFileInput.addEventListener('change', (e) => this.handleFileUpload(e));
        this.downloadAllBtn.addEventListener('click', () => this.downloadAllXMLs());
        this.validateGeneratedBtn.addEventListener('click', () => this.validateGeneratedXMLs());
        this.clearAllBtn.addEventListener('click', () => this.clearAll());
    }
    
    applyPaperMapping(paperValue, qualityValue) {
        if (CONFIG.VALUE_MAPPING.paper && CONFIG.VALUE_MAPPING.paper[paperValue]) {
            return CONFIG.VALUE_MAPPING.paper[paperValue];
        }
        
        if (paperValue === "Matte Coated 90gsm") {
            const qualityStr = String(qualityValue).toLowerCase();
            
            if (qualityStr.includes("standard")) {
                return "Clairjet 90 gsm";
            } else if (qualityStr.includes("premium")) {
                return "Magno Matt 90 gsm";
            }
        }
        
        return paperValue;
    }
    
    async handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        this.showProcessing('Reading Excel file...');
        
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonData.length < 2) {
                this.showStatus('error', 'Excel file must contain headers and at least one data row');
                this.hideProcessing();
                return;
            }
            
            this.headers = jsonData[0].map(h => String(h).trim());
            
            this.mappedHeaders = this.headers.map(header => {
                return CONFIG.COLUMN_MAPPING[header] || header;
            });
            
            const orderedHeaders = [];
            const orderedIndices = [];
            
            CONFIG.COLUMN_ORDER.forEach(orderedCol => {
                const index = this.mappedHeaders.indexOf(orderedCol);
                if (index !== -1) {
                    orderedHeaders.push(orderedCol);
                    orderedIndices.push(index);
                } else if (orderedCol === 'spine' || orderedCol === 'gsm') {
                    orderedHeaders.push(orderedCol);
                    orderedIndices.push(-1);
                }
            });
            
            this.mappedHeaders.forEach((header, index) => {
                if (!orderedHeaders.includes(header)) {
                    orderedHeaders.push(header);
                    orderedIndices.push(index);
                }
            });
            
            this.mappedHeaders = orderedHeaders;
            
            this.workbookData = jsonData.slice(1).filter(row => 
                row.some(cell => cell !== null && cell !== undefined && cell !== '')
            ).map(row => {
                const reorderedRow = orderedIndices.map(i => i === -1 ? '' : row[i]);
                
                const extentIndex = this.mappedHeaders.indexOf('extent');
                const paperIndex = this.mappedHeaders.indexOf('paper');
                const bindingIndex = this.mappedHeaders.indexOf('binding_style');
                const qualityIndex = this.mappedHeaders.indexOf('quality');
                const spineIndex = this.mappedHeaders.indexOf('spine');
                const gsmIndex = this.mappedHeaders.indexOf('gsm');
                
                if (spineIndex !== -1 && extentIndex !== -1 && paperIndex !== -1 && bindingIndex !== -1) {
                    const extent = parseInt(reorderedRow[extentIndex]) || 0;
                    let paperName = reorderedRow[paperIndex];
                    const bindingStyle = reorderedRow[bindingIndex];
                    const qualityValue = qualityIndex !== -1 ? reorderedRow[qualityIndex] : '';
                    
                    if (paperName) {
                        paperName = this.applyPaperMapping(paperName, qualityValue);
                        reorderedRow[paperIndex] = paperName;
                    }
                    
                    if (extent > 0 && paperName && bindingStyle) {
                        const spine = this.calculateSpine(extent, paperName, bindingStyle);
                        reorderedRow[spineIndex] = spine.toString();
                    }
                    
                    if (gsmIndex !== -1 && paperName) {
                        const gsm = this.extractGSM(paperName);
                        reorderedRow[gsmIndex] = gsm;
                    }
                }
                
                return reorderedRow;
            });
            
            if (this.workbookData.length === 0) {
                this.showStatus('error', 'No data rows found in Excel file');
                this.hideProcessing();
                return;
            }
            
            this.updateProcessing(`Processing ${this.workbookData.length} rows...`);
            
            this.displaySummary();
            
            this.downloadAllBtn.disabled = false;
            this.validateGeneratedBtn.disabled = false;
            this.clearAllBtn.disabled = false;
            
            this.showStatus('success', `Successfully loaded ${this.workbookData.length} records from Excel file`);
            this.hideProcessing();
            
        } catch (error) {
            console.error('Error processing Excel file:', error);
            this.showStatus('error', 'Error processing file: ' + error.message);
            this.hideProcessing();
        }
    }
    
    displaySummary() {
        // Calculate summary statistics
        const stats = {
            binding: {},
            colour: {},
            quality: {}
        };
        
        const bindingIndex = this.mappedHeaders.indexOf('binding_style');
        const colourIndex = this.mappedHeaders.indexOf('colour');
        const qualityIndex = this.mappedHeaders.indexOf('quality');
        
        this.workbookData.forEach(row => {
            if (bindingIndex !== -1 && row[bindingIndex]) {
                const binding = String(row[bindingIndex]);
                stats.binding[binding] = (stats.binding[binding] || 0) + 1;
            }
            
            if (colourIndex !== -1 && row[colourIndex]) {
                const colour = String(row[colourIndex]);
                stats.colour[colour] = (stats.colour[colour] || 0) + 1;
            }
            
            if (qualityIndex !== -1 && row[qualityIndex]) {
                const quality = String(row[qualityIndex]);
                stats.quality[quality] = (stats.quality[quality] || 0) + 1;
            }
        });
        
        // Display summary
        this.totalRecords.textContent = this.workbookData.length;
        
        // Format binding summary
        const bindingText = Object.entries(stats.binding)
            .map(([type, count]) => `${type}: ${count}`)
            .join('<br>') || 'N/A';
        this.bindingSummary.innerHTML = bindingText;
        
        // Format colour summary
        const colourText = Object.entries(stats.colour)
            .map(([type, count]) => `${type}: ${count}`)
            .join('<br>') || 'N/A';
        this.colourSummary.innerHTML = colourText;
        
        // Format quality summary
        const qualityText = Object.entries(stats.quality)
            .map(([type, count]) => `${type}: ${count}`)
            .join('<br>') || 'N/A';
        this.qualitySummary.innerHTML = qualityText;
        
        this.summarySection.style.display = 'block';
    }
    
    generateXMLForRow(row, index) {
        let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
        xml += '<record>\n';
        
        const qualityIndex = this.mappedHeaders.indexOf('quality');
        const qualityValue = qualityIndex !== -1 ? row[qualityIndex] : '';
        
        this.mappedHeaders.forEach((header, colIndex) => {
            let value = row[colIndex] !== undefined ? row[colIndex] : '';
            
            if (header === 'paper' && value) {
                value = this.applyPaperMapping(value, qualityValue);
            } else if (CONFIG.VALUE_MAPPING[header] && CONFIG.VALUE_MAPPING[header][value]) {
                value = CONFIG.VALUE_MAPPING[header][value];
            }
            
            const cleanHeader = this.sanitizeXMLTag(header);
            const cleanValue = this.escapeXML(String(value));
            xml += `  <${cleanHeader}>${cleanValue}</${cleanHeader}>\n`;
        });
        
        xml += '</record>';
        return xml;
    }
    
    async downloadAllXMLs() {
        this.showProcessing('Creating ZIP file...');
        
        try {
            const zip = new JSZip();
            this.generatedXMLs = [];
            
            this.workbookData.forEach((row, index) => {
                const xml = this.generateXMLForRow(row, index);
                
                const isbnIndex = this.mappedHeaders.indexOf('isbn');
                let filename = `record_${String(index + 1).padStart(4, '0')}.xml`;
                
                if (isbnIndex !== -1 && row[isbnIndex]) {
                    const isbn = String(row[isbnIndex]).replace(/[^a-zA-Z0-9]/g, '');
                    filename = `${isbn}.xml`;
                }
                
                zip.file(filename, xml);
                this.generatedXMLs.push({ filename, content: xml });
            });
            
            this.updateProcessing('Generating ZIP archive...');
            
            const blob = await zip.generateAsync({ type: 'blob' });
            
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `cup_excel_to_xml_${new Date().getTime()}.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
            this.showStatus('success', `Successfully downloaded ${this.workbookData.length} XML files as ZIP`);
            this.hideProcessing();
            
        } catch (error) {
            console.error('Error creating ZIP:', error);
            this.showStatus('error', 'Error creating ZIP file: ' + error.message);
            this.hideProcessing();
        }
    }
    
    validateGeneratedXMLs() {
        if (this.generatedXMLs.length === 0) {
            this.generatedXMLs = [];
            this.workbookData.forEach((row, index) => {
                const xml = this.generateXMLForRow(row, index);
                const isbnIndex = this.mappedHeaders.indexOf('isbn');
                let filename = `record_${String(index + 1).padStart(4, '0')}.xml`;
                
                if (isbnIndex !== -1 && row[isbnIndex]) {
                    const isbn = String(row[isbnIndex]).replace(/[^a-zA-Z0-9]/g, '');
                    filename = `${isbn}.xml`;
                }
                
                this.generatedXMLs.push({ filename, content: xml });
            });
        }
        
        // Switch to validator tab
        const validatorTab = document.getElementById('validator-tab');
        validatorTab.click();
        
        // Trigger validation in validator module
        if (window.validatorModule) {
            window.validatorModule.validateFromConverter(this.generatedXMLs);
        }
    }
    
    clearAll() {
        this.workbookData = [];
        this.headers = [];
        this.mappedHeaders = [];
        this.generatedXMLs = [];
        this.excelFileInput.value = '';
        this.summarySection.style.display = 'none';
        this.downloadAllBtn.disabled = true;
        this.validateGeneratedBtn.disabled = true;
        this.clearAllBtn.disabled = true;
        this.statusAlert.classList.add('d-none');
    }
    
    calculateSpine(extent, paperName, bindingStyle) {
        const paperSpec = CONFIG.PAPER_SPECS[paperName];
        if (!paperSpec) return 0;
        
        const baseSpineThickness = Math.round(
            (extent * paperSpec.grammage * paperSpec.volume) / CONFIG.SPINE_CALCULATION_FACTOR
        );
        
        return bindingStyle.toLowerCase() === 'cased' ?
            baseSpineThickness + CONFIG.HARDBACK_SPINE_ADDITION :
            baseSpineThickness;
    }
    
    extractGSM(paperName) {
        if (!paperName) return '';
        
        const paperSpec = CONFIG.PAPER_SPECS[paperName];
        if (paperSpec) {
            return paperSpec.grammage.toString();
        }
        
        const match = paperName.match(/(\d+)\s*gsm/i);
        return match ? match[1] : '';
    }
    
    sanitizeXMLTag(tag) {
        return tag.replace(/[^a-zA-Z0-9_-]/g, '_').replace(/^[^a-zA-Z_]/, '_');
    }
    
    escapeXML(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
    
    showProcessing(message) {
        this.processingText.textContent = message;
        this.processingIndicator.style.display = 'block';
    }
    
    updateProcessing(message) {
        this.processingText.textContent = message;
    }
    
    hideProcessing() {
        this.processingIndicator.style.display = 'none';
    }
    
    showStatus(type, message) {
        this.statusAlert.className = 'alert';
        this.statusAlert.classList.remove('d-none');
        
        if (type === 'success') {
            this.statusAlert.classList.add('alert-success');
            this.statusAlert.innerHTML = `<i class="bi bi-check-circle-fill me-2"></i>${message}`;
        } else if (type === 'error') {
            this.statusAlert.classList.add('alert-danger');
            this.statusAlert.innerHTML = `<i class="bi bi-exclamation-triangle-fill me-2"></i>${message}`;
        } else {
            this.statusAlert.classList.add('alert-info');
            this.statusAlert.innerHTML = `<i class="bi bi-info-circle-fill me-2"></i>${message}`;
        }
    }
}
