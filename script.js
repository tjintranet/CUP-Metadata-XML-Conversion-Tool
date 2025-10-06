// Global variables
let workbookData = [];
let headers = [];
let mappedHeaders = [];

// Column mapping from original Excel headers to XML attribute names
const COLUMN_MAPPING = {
    "Master ISBN 13": "isbn",
    "full_title_for_export": "title",
    "pod_ExtentMod2": "extent",
    "pod_xml_TrimHeight": "trim_height",
    "pod_xml_TrimWidth": "trim_width",
    "pod_xml_TJ_PaperStock": "paper",
    "pod_Colourspace": "colour",
    "pod_xml_TJ_QualityOption": "quality",
    "pod_ColourBook_fullBleed": "bleed",
    "pod_TJ_BindType_c": "binding_style",
    "pod_CoverLaminate": "lamination",
    "pod_xml_Jacket": "jacket",
    "pod_xml_TJ_JacketLamination": "jacket_lamination",
    "pod_xml_TJ_FlapSize": "flap_size",
    "pod_xml_TJ_WibalinColour": "wibalin_colour",
    "pod_xml_SetISBN": "set_isbn"
};

// Desired column order for display and XML
const COLUMN_ORDER = [
    "isbn",
    "title",
    "trim_height",
    "trim_width",
    "extent",
    "spine",
    "paper",
    "gsm",
    "colour",
    "quality",
    "bleed",
    "binding_style",
    "lamination",
    "jacket",
    "jacket_lamination",
    "flap_size",
    "wibalin_colour",
    "set_isbn"
];

// Value remapping for specific fields
const VALUE_MAPPING = {
    "paper": {
        "Cream 80gsm": "Munken 80 gsm",
        "White 80gsm": "Navigator 80 gsm",
        "Matte Coated 90gsm": "LetsGo Silk 90 gsm"
    },
    "lamination": {
        "Matte": "Matt"
    }
};

// Paper specifications for spine calculation
const PAPER_SPECS = {
    "Munken 80 gsm": { grammage: 80, volume: 17.5 },
    "Navigator 80 gsm": { grammage: 80, volume: 12.5 },
    "LetsGo Silk 90 gsm": { grammage: 90, volume: 10 }
};

// Constants for spine calculation
const SPINE_CALCULATION_FACTOR = 20000;
const HARDBACK_SPINE_ADDITION = 4;

// DOM elements
const excelFileInput = document.getElementById('excelFile');
const downloadAllBtn = document.getElementById('downloadAllBtn');
const clearAllBtn = document.getElementById('clearAllBtn');
const processingIndicator = document.getElementById('processingIndicator');
const processingText = document.getElementById('processingText');
const statusAlert = document.getElementById('statusAlert');
const previewTableSection = document.getElementById('previewTableSection');
const previewHeaders = document.getElementById('previewHeaders');
const previewBody = document.getElementById('previewBody');
const rowCountInfo = document.getElementById('rowCountInfo');

// Event listeners
excelFileInput.addEventListener('change', handleFileUpload);
downloadAllBtn.addEventListener('click', downloadAllXMLs);
clearAllBtn.addEventListener('click', clearAll);

// Handle file upload
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    showProcessing('Reading Excel file...');
    
    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Get first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (jsonData.length < 2) {
            showStatus('error', 'Excel file must contain headers and at least one data row');
            hideProcessing();
            return;
        }
        
        // Extract headers and create mapping
        headers = jsonData[0].map(h => String(h).trim());
        
        // Map headers to XML attribute names
        mappedHeaders = headers.map(header => {
            return COLUMN_MAPPING[header] || header;
        });
        
        // Reorder headers according to COLUMN_ORDER
        const orderedHeaders = [];
        const orderedIndices = [];
        
        COLUMN_ORDER.forEach(orderedCol => {
            const index = mappedHeaders.indexOf(orderedCol);
            if (index !== -1) {
                orderedHeaders.push(orderedCol);
                orderedIndices.push(index);
            } else if (orderedCol === 'spine' || orderedCol === 'gsm') {
                // Add spine and gsm columns even if not in original data
                orderedHeaders.push(orderedCol);
                orderedIndices.push(-1); // -1 indicates calculated field
            }
        });
        
        // Add any unmapped columns at the end
        mappedHeaders.forEach((header, index) => {
            if (!orderedHeaders.includes(header)) {
                orderedHeaders.push(header);
                orderedIndices.push(index);
            }
        });
        
        mappedHeaders = orderedHeaders;
        
        console.log('Original headers:', headers);
        console.log('Mapped headers:', mappedHeaders);
        
        // Reorder data columns to match new header order and calculate spine and gsm
        workbookData = jsonData.slice(1).filter(row => 
            row.some(cell => cell !== null && cell !== undefined && cell !== '')
        ).map(row => {
            const reorderedRow = orderedIndices.map(i => i === -1 ? '' : row[i]);
            
            // Calculate spine value
            const extentIndex = mappedHeaders.indexOf('extent');
            const paperIndex = mappedHeaders.indexOf('paper');
            const bindingIndex = mappedHeaders.indexOf('binding_style');
            const spineIndex = mappedHeaders.indexOf('spine');
            const gsmIndex = mappedHeaders.indexOf('gsm');
            
            if (spineIndex !== -1 && extentIndex !== -1 && paperIndex !== -1 && bindingIndex !== -1) {
                const extent = parseInt(reorderedRow[extentIndex]) || 0;
                let paperName = reorderedRow[paperIndex];
                const bindingStyle = reorderedRow[bindingIndex];
                
                // Apply paper mapping if needed
                if (VALUE_MAPPING.paper && VALUE_MAPPING.paper[paperName]) {
                    paperName = VALUE_MAPPING.paper[paperName];
                }
                
                // Calculate spine
                if (extent > 0 && paperName && bindingStyle) {
                    const spine = calculateSpine(extent, paperName, bindingStyle);
                    reorderedRow[spineIndex] = spine.toString();
                }
                
                // Extract GSM value from paper name
                if (gsmIndex !== -1 && paperName) {
                    const gsm = extractGSM(paperName);
                    reorderedRow[gsmIndex] = gsm;
                }
            }
            
            return reorderedRow;
        });
        
        if (workbookData.length === 0) {
            showStatus('error', 'No data rows found in Excel file');
            hideProcessing();
            return;
        }
        
        updateProcessing(`Processing ${workbookData.length} rows...`);
        
        // Display preview
        displayPreview();
        
        // Enable buttons
        downloadAllBtn.disabled = false;
        clearAllBtn.disabled = false;
        
        showStatus('success', `Successfully loaded ${workbookData.length} records from Excel file`);
        hideProcessing();
        
    } catch (error) {
        console.error('Error processing file:', error);
        showStatus('error', 'Error processing Excel file: ' + error.message);
        hideProcessing();
    }
}

// Display preview table
function displayPreview() {
    // Clear existing content
    previewHeaders.innerHTML = '';
    previewBody.innerHTML = '';
    
    // Calculate statistics
    const stats = calculateStatistics();
    
    // Create summary
    const summaryInfo = document.createElement('div');
    summaryInfo.className = 'alert alert-info mb-3';
    summaryInfo.innerHTML = `
        <h5 class="alert-heading"><i class="bi bi-info-circle me-2"></i>Processing Summary</h5>
        <hr>
        <div class="row">
            <div class="col-md-6">
                <p class="mb-2"><strong>Total Records:</strong> ${workbookData.length}</p>
                <p class="mb-2"><strong>Binding Style:</strong></p>
                <ul class="mb-2">
                    <li>Cased: ${stats.binding.Cased || 0}</li>
                    <li>Limp: ${stats.binding.Limp || 0}</li>
                </ul>
                <p class="mb-2"><strong>Quality:</strong></p>
                <ul class="mb-2">
                    <li>Standard: ${stats.quality.Standard || 0}</li>
                    <li>Premium: ${stats.quality.Premium || 0}</li>
                </ul>
            </div>
            <div class="col-md-6">
                <p class="mb-2"><strong>Colour:</strong></p>
                <ul class="mb-2">
                    <li>Mono: ${stats.colour.Mono || 0}</li>
                    <li>Colour: ${stats.colour.Colour || 0}</li>
                </ul>
                <p class="mb-2"><strong>Jobs with Bleeds:</strong> ${stats.bleeds}</p>
                <p class="mb-2"><strong>Jackets:</strong> ${stats.jackets}</p>
                <p class="mb-2"><strong>Lamination Types:</strong></p>
                <ul class="mb-0">
                    ${Object.entries(stats.lamination).map(([type, count]) => `<li>${type}: ${count}</li>`).join('')}
                </ul>
            </div>
        </div>
    `;
    
    // Clear preview section and add summary
    const previewContainer = previewTableSection.querySelector('.table-responsive');
    if (previewContainer) {
        previewContainer.innerHTML = '';
        previewContainer.appendChild(summaryInfo);
    }
    
    // Update row count
    rowCountInfo.textContent = `${workbookData.length} records`;
    
    // Show preview section
    previewTableSection.style.display = 'block';
}

// Calculate statistics for summary
function calculateStatistics() {
    const stats = {
        binding: {},
        quality: {},
        colour: {},
        bleeds: 0,
        jackets: 0,
        lamination: {}
    };
    
    const bindingIndex = mappedHeaders.indexOf('binding_style');
    const qualityIndex = mappedHeaders.indexOf('quality');
    const colourIndex = mappedHeaders.indexOf('colour');
    const bleedIndex = mappedHeaders.indexOf('bleed');
    const jacketIndex = mappedHeaders.indexOf('jacket');
    const laminationIndex = mappedHeaders.indexOf('lamination');
    
    workbookData.forEach(row => {
        // Count binding styles
        if (bindingIndex !== -1 && row[bindingIndex]) {
            const binding = String(row[bindingIndex]);
            stats.binding[binding] = (stats.binding[binding] || 0) + 1;
        }
        
        // Count quality
        if (qualityIndex !== -1 && row[qualityIndex]) {
            const quality = String(row[qualityIndex]);
            stats.quality[quality] = (stats.quality[quality] || 0) + 1;
        }
        
        // Count colour
        if (colourIndex !== -1 && row[colourIndex]) {
            const colour = String(row[colourIndex]);
            stats.colour[colour] = (stats.colour[colour] || 0) + 1;
        }
        
        // Count bleeds (Y values)
        if (bleedIndex !== -1 && row[bleedIndex]) {
            const bleed = String(row[bleedIndex]).toUpperCase();
            if (bleed === 'Y' || bleed === 'YES') {
                stats.bleeds++;
            }
        }
        
        // Count jackets (Y values)
        if (jacketIndex !== -1 && row[jacketIndex]) {
            const jacket = String(row[jacketIndex]).toUpperCase();
            if (jacket === 'Y' || jacket === 'YES') {
                stats.jackets++;
            }
        }
        
        // Count lamination types
        if (laminationIndex !== -1 && row[laminationIndex]) {
            let lamination = String(row[laminationIndex]);
            // Apply value mapping
            if (VALUE_MAPPING.lamination && VALUE_MAPPING.lamination[lamination]) {
                lamination = VALUE_MAPPING.lamination[lamination];
            }
            stats.lamination[lamination] = (stats.lamination[lamination] || 0) + 1;
        }
    });
    
    return stats;
}

// Generate XML for a single row
function generateXMLForRow(row, index) {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xml += '<record>\n';
    
    mappedHeaders.forEach((header, colIndex) => {
        let value = row[colIndex] !== undefined ? row[colIndex] : '';
        
        // Apply value mapping if exists for this field
        if (VALUE_MAPPING[header] && VALUE_MAPPING[header][value]) {
            value = VALUE_MAPPING[header][value];
        }
        
        const cleanHeader = sanitizeXMLTag(header);
        const cleanValue = escapeXML(String(value));
        xml += `  <${cleanHeader}>${cleanValue}</${cleanHeader}>\n`;
    });
    
    xml += '</record>';
    return xml;
}

// Download all XMLs as ZIP
async function downloadAllXMLs() {
    showProcessing('Creating ZIP file...');
    
    try {
        const zip = new JSZip();
        
        // Generate XML for each row
        workbookData.forEach((row, index) => {
            const xml = generateXMLForRow(row, index);
            
            // Get ISBN for filename (first column should be isbn)
            const isbnIndex = mappedHeaders.indexOf('isbn');
            let filename = `record_${String(index + 1).padStart(4, '0')}.xml`;
            
            if (isbnIndex !== -1 && row[isbnIndex]) {
                const isbn = String(row[isbnIndex]).replace(/[^a-zA-Z0-9]/g, '');
                filename = `${isbn}.xml`;
            }
            
            zip.file(filename, xml);
        });
        
        updateProcessing('Generating ZIP archive...');
        
        // Generate ZIP file
        const blob = await zip.generateAsync({ type: 'blob' });
        
        // Download
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `excel_to_xml_${new Date().getTime()}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showStatus('success', `Successfully downloaded ${workbookData.length} XML files as ZIP`);
        hideProcessing();
        
    } catch (error) {
        console.error('Error creating ZIP:', error);
        showStatus('error', 'Error creating ZIP file: ' + error.message);
        hideProcessing();
    }
}

// Clear all data
function clearAll() {
    workbookData = [];
    headers = [];
    mappedHeaders = [];
    excelFileInput.value = '';
    previewHeaders.innerHTML = '';
    previewBody.innerHTML = '';
    previewTableSection.style.display = 'none';
    downloadAllBtn.disabled = true;
    clearAllBtn.disabled = true;
    statusAlert.classList.add('d-none');
    rowCountInfo.textContent = '0 records';
}

// Calculate spine size
function calculateSpine(extent, paperName, bindingStyle) {
    const paperSpec = PAPER_SPECS[paperName];
    if (!paperSpec) return 0;
    
    const baseSpineThickness = Math.round(
        (extent * paperSpec.grammage * paperSpec.volume) / SPINE_CALCULATION_FACTOR
    );
    
    return bindingStyle.toLowerCase() === 'cased' ?
        baseSpineThickness + HARDBACK_SPINE_ADDITION :
        baseSpineThickness;
}

// Extract GSM value from paper name
function extractGSM(paperName) {
    if (!paperName) return '';
    
    // First check if it's a known paper type
    const paperSpec = PAPER_SPECS[paperName];
    if (paperSpec) {
        return paperSpec.grammage.toString();
    }
    
    // Otherwise try to extract number followed by 'gsm' (case insensitive)
    const match = paperName.match(/(\d+)\s*gsm/i);
    return match ? match[1] : '';
}

// Utility functions
function sanitizeXMLTag(tag) {
    // Remove invalid XML tag characters and replace spaces with underscores
    return tag.replace(/[^a-zA-Z0-9_-]/g, '_').replace(/^[^a-zA-Z_]/, '_');
}

function escapeXML(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showProcessing(message) {
    processingText.textContent = message;
    processingIndicator.style.display = 'block';
}

function updateProcessing(message) {
    processingText.textContent = message;
}

function hideProcessing() {
    processingIndicator.style.display = 'none';
}

function showStatus(type, message) {
    statusAlert.className = 'alert';
    statusAlert.classList.remove('d-none');
    
    if (type === 'success') {
        statusAlert.classList.add('alert-success');
        statusAlert.innerHTML = `<i class="bi bi-check-circle-fill me-2"></i>${message}`;
    } else if (type === 'error') {
        statusAlert.classList.add('alert-danger');
        statusAlert.innerHTML = `<i class="bi bi-exclamation-triangle-fill me-2"></i>${message}`;
    } else {
        statusAlert.classList.add('alert-info');
        statusAlert.innerHTML = `<i class="bi bi-info-circle-fill me-2"></i>${message}`;
    }
}