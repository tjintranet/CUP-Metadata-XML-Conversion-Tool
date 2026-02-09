// XML Validator Module
class ValidatorModule {
    constructor() {
        this.validationResults = [];
        this.showFailedOnly = false;
        this.validator = new CUPValidator();
        
        this.initializeElements();
        this.attachEventListeners();
    }
    
    initializeElements() {
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.resultsArea = document.getElementById('resultsArea');
        this.downloadBtn = document.getElementById('downloadValidationBtn');
        this.downloadFailedBtn = document.getElementById('downloadFailedBtn');
        this.clearBtn = document.getElementById('clearValidatorBtn');
        this.processingSpinner = document.getElementById('validatorProcessingSpinner');
        this.resultsSummary = document.getElementById('resultsSummary');
        this.totalFilesEl = document.getElementById('totalFiles');
        this.totalShowingEl = document.getElementById('totalShowing');
        this.totalPassedEl = document.getElementById('totalPassed');
        this.totalFailedEl = document.getElementById('totalFailed');
        this.failedOnlyToggle = document.getElementById('failedOnlyToggle');
    }
    
    attachEventListeners() {
        this.failedOnlyToggle.addEventListener('change', (e) => {
            this.showFailedOnly = e.target.checked;
            this.updateResultsDisplay();
        });

        this.fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files).filter(file => 
                file.name.toLowerCase().endsWith('.xml')
            );
            if (files.length > 0) {
                this.handleFiles(files);
            }
        });

        this.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.dropZone.classList.add('dragover');
        });

        this.dropZone.addEventListener('dragleave', () => {
            this.dropZone.classList.remove('dragover');
        });

        this.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.dropZone.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files).filter(file => 
                file.name.toLowerCase().endsWith('.xml')
            );
            if (files.length > 0) {
                this.handleFiles(files);
            }
        });

        this.clearBtn.addEventListener('click', () => {
            this.validationResults = [];
            this.showFailedOnly = false;
            this.failedOnlyToggle.checked = false;
            this.updateResultsDisplay();
            this.downloadBtn.disabled = true;
            this.downloadFailedBtn.disabled = true;
            this.clearBtn.disabled = true;
        });

        this.downloadBtn.addEventListener('click', (e) => {
            e.preventDefault();
            if (this.validationResults.length > 0) {
                this.downloadResults(false);
            }
        });

        this.downloadFailedBtn.addEventListener('click', (e) => {
            e.preventDefault();
            if (this.validationResults.length > 0) {
                this.downloadResults(true);
            }
        });
    }
    
    validateFromConverter(xmlFiles) {
        this.processingSpinner.style.display = 'block';
        this.dropZone.style.display = 'none';
        
        this.validationResults = [];
        
        xmlFiles.forEach(({ filename, content }) => {
            const result = this.validateXMLContent(content, filename);
            this.validationResults.push({
                fileName: filename,
                isbn: result.isbn,
                title: result.title,
                validations: result.validations
            });
        });
        
        this.updateResultsDisplay();
        this.downloadBtn.disabled = false;
        this.downloadFailedBtn.disabled = false;
        this.clearBtn.disabled = false;
        this.processingSpinner.style.display = 'none';
    }
    
    readFileContent(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = () => reject(new Error('Error reading file'));
            reader.readAsText(file);
        });
    }
    
    validateXMLContent(xmlContent, fileName) {
        try {
            if (typeof xmlContent !== 'string') {
                return {
                    validations: [{
                        test: 'XML Format',
                        result: false,
                        message: 'Invalid XML content type'
                    }],
                    isbn: '',
                    title: ''
                };
            }

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(xmlContent, 'text/xml');
            
            const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
            if (parseError) {
                return {
                    validations: [{
                        test: 'XML Format',
                        result: false,
                        message: `Invalid XML format: ${parseError.textContent}`
                    }],
                    isbn: '',
                    title: ''
                };
            }

            const recordNode = xmlDoc.getElementsByTagName('record')[0];
            if (!recordNode) {
                return {
                    validations: [{
                        test: 'XML Structure',
                        result: false,
                        message: 'Missing root record element'
                    }],
                    isbn: '',
                    title: ''
                };
            }

            const isbn = recordNode.querySelector('isbn')?.textContent?.trim() || '';
            const title = recordNode.querySelector('title')?.textContent?.trim() || '';

            const validations = this.validator.validate(recordNode);
            return {
                validations: validations,
                isbn: isbn,
                title: title
            };
        } catch (error) {
            console.error('Error processing XML:', error);
            return {
                validations: [{
                    test: 'Processing Error',
                    result: false,
                    message: `Error: ${error.message}`
                }],
                isbn: '',
                title: ''
            };
        }
    }
    
    async handleFiles(files) {
        if (files.length === 0) return;

        this.processingSpinner.style.display = 'block';
        this.dropZone.style.display = 'none';

        try {
            for (const file of files) {
                try {
                    const content = await this.readFileContent(file);
                    const result = this.validateXMLContent(content, file.name);
                    
                    this.validationResults.push({
                        fileName: file.name,
                        isbn: result.isbn,
                        title: result.title,
                        validations: result.validations
                    });
                } catch (error) {
                    console.error(`Error processing ${file.name}:`, error);
                    this.validationResults.push({
                        fileName: file.name,
                        isbn: '',
                        title: '',
                        validations: [{
                            test: 'File Error',
                            result: false,
                            message: `Failed to process file: ${error.message}`
                        }]
                    });
                }
            }

            this.updateResultsDisplay();
            this.downloadBtn.disabled = false;
            this.downloadFailedBtn.disabled = false;
            this.clearBtn.disabled = false;

        } catch (error) {
            console.error('Error handling files:', error);
            alert('An error occurred while processing files. Please try again.');
        } finally {
            this.processingSpinner.style.display = 'none';
            this.fileInput.value = '';
        }
    }
    
    calculateSummary() {
        const total = this.validationResults.length;
        const passed = this.validationResults.filter(r => 
            r.validations.every(v => v.result)
        ).length;
        const failed = total - passed;

        return {
            total: total,
            showing: total,
            passed: passed,
            failed: failed
        };
    }
    
    updateResultsDisplay() {
        if (this.validationResults.length === 0) {
            this.resultsArea.innerHTML = `
                <div class="col-12 text-center text-muted">
                    <p>Upload XML files to see validation results</p>
                </div>
            `;
            this.resultsSummary.classList.add('d-none');
            this.dropZone.style.display = 'flex';
            return;
        }

        const summary = this.calculateSummary();
        const filteredResults = this.showFailedOnly 
            ? this.validationResults.filter(r => r.validations.some(v => !v.result))
            : this.validationResults;

        this.totalFilesEl.textContent = summary.total;
        this.totalShowingEl.textContent = filteredResults.length;
        this.totalPassedEl.textContent = summary.passed;
        this.totalFailedEl.textContent = summary.failed;

        this.resultsSummary.classList.remove('d-none');

        this.resultsArea.innerHTML = filteredResults
            .map((fileResult) => {
                const originalIndex = this.validationResults.indexOf(fileResult);
                const hasErrors = fileResult.validations.some(v => !v.result);
                const displayName = fileResult.isbn || fileResult.fileName;
                
                return `
                    <div class="col-12 fade-in">
                        <div class="card ${hasErrors ? 'border-danger' : 'border-success'} compact-card">
                            <div class="card-header ${hasErrors ? 'bg-danger' : 'bg-success'} bg-opacity-10 d-flex justify-content-between align-items-center py-2">
                                <div>
                                    <strong>${this.escapeHtml(displayName)}</strong>
                                    ${fileResult.title ? `<span class="ms-2 text-muted small">${this.escapeHtml(fileResult.title)}</span>` : ''}
                                </div>
                                <div class="d-flex align-items-center gap-2">
                                    <button class="btn btn-sm btn-outline-secondary" 
                                            onclick="window.validatorModule.copyToClipboard(${originalIndex})"
                                            data-isbn="${this.escapeHtml(fileResult.isbn)}"
                                            id="copy-btn-${originalIndex}">
                                        <i class="bi bi-clipboard"></i>
                                    </button>
                                    <span class="badge bg-${hasErrors ? 'danger' : 'success'}">
                                        ${hasErrors ? 'Failed' : 'Passed'}
                                    </span>
                                </div>
                            </div>
                            <div class="card-body py-2">
                                <ul class="list-unstyled mb-0 compact-list">
                                    ${fileResult.validations.map(validation => `
                                        <li class="d-flex align-items-start py-1">
                                            <i class="bi ${validation.result ? 'bi-check-circle-fill text-success' : 'bi-x-circle-fill text-danger'} me-2"></i>
                                            <small>
                                                <strong>${this.escapeHtml(validation.test)}:</strong> 
                                                ${this.escapeHtml(validation.message)}
                                            </small>
                                        </li>
                                    `).join('')}
                                </ul>
                            </div>
                        </div>
                    </div>
                `;
            })
            .join('');
    }
    
    generateTextReport(failedOnly = false) {
        if (this.validationResults.length === 0) return '';
        
        const timestamp = new Date().toLocaleString();
        const summary = this.calculateSummary();
        
        // Filter to only failed validations if requested
        const resultsToReport = failedOnly 
            ? this.validationResults.filter(r => r.validations.some(v => !v.result))
            : this.validationResults;
        
        if (failedOnly && resultsToReport.length === 0) {
            return 'No failed validations found. All files passed validation.';
        }
        
        let report = '';
        
        report += '='.repeat(80) + '\n';
        report += failedOnly ? 'CUP XML VALIDATION - FAILED VALIDATIONS ONLY\n' : 'CUP XML VALIDATION ERROR REPORT\n';
        report += '='.repeat(80) + '\n';
        report += `Generated: ${timestamp}\n\n`;
        
        report += 'SUMMARY\n';
        report += '-'.repeat(40) + '\n';
        report += `Total Files Processed: ${summary.total}\n`;
        if (failedOnly) {
            report += `Files with Errors: ${resultsToReport.length}\n`;
        }
        report += '\n';
        
        report += failedOnly ? 'FAILED VALIDATIONS ONLY\n' : 'VALIDATION RESULTS\n';
        report += '='.repeat(80) + '\n\n';
        
        resultsToReport.forEach((fileResult, index) => {
            const hasErrors = fileResult.validations.some(v => !v.result);
            const status = hasErrors ? 'FAILED' : 'PASSED';
            
            report += `${index + 1}. ISBN: ${fileResult.isbn || 'N/A'}\n`;
            if (fileResult.title) {
                report += `   Title: ${fileResult.title}\n`;
            }
            report += `   Status: ${status}\n`;
            report += '-'.repeat(80) + '\n';
            
            const totalTests = fileResult.validations.length;
            const passedTests = fileResult.validations.filter(v => v.result).length;
            const failedTests = totalTests - passedTests;
            
            report += `   Tests: ${passedTests}/${totalTests} passed`;
            if (failedTests > 0) {
                report += ` (${failedTests} failed)`;
            }
            report += '\n';
            
            const failedValidations = fileResult.validations.filter(v => !v.result);
            if (failedValidations.length > 0) {
                report += '   Failed Tests:\n';
                failedValidations.forEach(validation => {
                    report += `   ✗ ${validation.test}: ${validation.message}\n`;
                });
            }
            
            report += '\n';
        });
        
        report += '='.repeat(80) + '\n';
        report += 'End of Report\n';
        report += '='.repeat(80) + '\n';
        
        return report;
    }
    
    downloadResults(failedOnly = false) {
        if (this.validationResults.length === 0) return;
        
        const reportContent = this.generateTextReport(failedOnly);
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
        const filename = failedOnly 
            ? `cup_validation_failed_${timestamp}.txt`
            : `cup_validation_summary_${timestamp}.txt`;
        
        const blob = new Blob([reportContent], { type: 'text/plain;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }
    
    async copyToClipboard(index) {
        const fileResult = this.validationResults[index];
        const isbn = fileResult.isbn || fileResult.fileName;
        const hasErrors = fileResult.validations.some(v => !v.result);
        
        let text = `Validation Results for ISBN: ${isbn}\n`;
        if (fileResult.title) {
            text += `Title: ${fileResult.title}\n`;
        }
        text += `Status: ${hasErrors ? 'Failed' : 'Passed'}\n`;
        text += '----------------------------------------\n';
        
        fileResult.validations.forEach(validation => {
            text += `${validation.result ? '✓' : '✗'} ${validation.test}: ${validation.message}\n`;
        });
        
        text += '----------------------------------------\n';
        
        try {
            await navigator.clipboard.writeText(text);
            
            const button = document.getElementById(`copy-btn-${index}`);
            
            if (button) {
                const originalHTML = button.innerHTML;
                button.innerHTML = '<i class="bi bi-check-circle"></i>';
                button.classList.remove('btn-outline-secondary');
                button.classList.add('btn-outline-success');
                
                setTimeout(() => {
                    button.innerHTML = originalHTML;
                    button.classList.remove('btn-outline-success');
                    button.classList.add('btn-outline-secondary');
                }, 2000);
            }
        } catch (err) {
            console.error('Failed to copy:', err);
            alert('Failed to copy to clipboard. Please try again.');
        }
    }
    
    escapeHtml(unsafe) {
        if (!unsafe) return '';
        return String(unsafe)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }
}

// CUP Validator Class
class CUPValidator {
    constructor() {
        this.validationSteps = [
            this.validateRequiredFields,
            this.validateBinding,
            this.validatePaper,
            this.validateColour,
            this.validateRoute,
            this.validateTrimSize,
            this.validateColourPaperCompatibility,
            this.validateRoutePaperCompatibility
        ];
    }

    validate(xmlDoc) {
        const results = [];
        
        const context = {
            isbn: xmlDoc.querySelector('isbn')?.textContent?.trim(),
            title: xmlDoc.querySelector('title')?.textContent?.trim(),
            trim_height: xmlDoc.querySelector('trim_height')?.textContent?.trim(),
            trim_width: xmlDoc.querySelector('trim_width')?.textContent?.trim(),
            extent: xmlDoc.querySelector('extent')?.textContent?.trim(),
            paper: xmlDoc.querySelector('paper')?.textContent?.trim(),
            colour: xmlDoc.querySelector('colour')?.textContent?.trim(),
            quality: xmlDoc.querySelector('quality')?.textContent?.trim(),
            binding_style: xmlDoc.querySelector('binding_style')?.textContent?.trim(),
            results: results
        };

        for (const step of this.validationSteps) {
            step.call(this, context);
        }

        return results;
    }

    addResult(results, test, result, message) {
        results.push({ test, result, message });
    }

    validateRequiredFields(context) {
        const { isbn, trim_height, trim_width, extent, paper, colour, quality, binding_style, results } = context;
        
        const requiredFields = {
            'ISBN': isbn,
            'Trim Height': trim_height,
            'Trim Width': trim_width,
            'Extent': extent,
            'Paper': paper,
            'Colour': colour,
            'Quality': quality,
            'Binding Style': binding_style
        };

        const missingFields = Object.entries(requiredFields)
            .filter(([_, value]) => !value || value.trim() === '')
            .map(([field]) => field);

        const allFieldsPresent = missingFields.length === 0;

        this.addResult(
            results,
            'Required Fields',
            allFieldsPresent,
            allFieldsPresent 
                ? 'All required fields are present and have values'
                : `Missing or empty fields: ${missingFields.join(', ')}`
        );
    }

    validateBinding(context) {
        const { binding_style, results } = context;
        const isValid = CONFIG.VALID_BINDINGS.has(binding_style);
        
        this.addResult(
            results,
            'Binding Style',
            isValid,
            isValid
                ? `Valid binding: ${binding_style}`
                : `Invalid binding: '${binding_style}' (must be Cased or Limp)`
        );
    }

    validatePaper(context) {
        const { paper, results } = context;
        const isValid = CONFIG.VALID_PAPERS.has(paper);
        
        this.addResult(
            results,
            'Paper Type',
            isValid,
            isValid
                ? `Valid paper: ${paper}`
                : `Invalid paper: '${paper}' (must be CUP MunkenPure 80 gsm, Navigator 80 gsm, Clairjet 90 gsm, or Magno Matt 90 gsm)`
        );
    }

    validateColour(context) {
        const { colour, results } = context;
        const isValid = CONFIG.VALID_COLOURS.has(colour);
        
        this.addResult(
            results,
            'Colour',
            isValid,
            isValid
                ? `Valid colour: ${colour}`
                : `Invalid colour: '${colour}' (must be Mono or Colour)`
        );
    }

    validateRoute(context) {
        const { quality, results } = context;
        const isValid = CONFIG.VALID_ROUTES.has(quality);
        
        this.addResult(
            results,
            'Quality/Route',
            isValid,
            isValid
                ? `Valid route: ${quality}`
                : `Invalid route: '${quality}' (must be Standard or Premium)`
        );
    }

    validateTrimSize(context) {
        const { trim_width, trim_height, results } = context;
        
        const width = parseFloat(trim_width);
        const height = parseFloat(trim_height);
        
        if (isNaN(width) || isNaN(height) || width <= 0 || height <= 0) {
            this.addResult(
                results,
                'Trim Size',
                false,
                `Invalid dimensions: ${trim_width}x${trim_height}mm (must be positive numbers)`
            );
            return;
        }
        
        const trimSize = `${Math.round(width)}x${Math.round(height)}`;
        const isValid = CONFIG.VALID_TRIM_SIZES.has(trimSize);
        
        this.addResult(
            results,
            'Trim Size',
            isValid,
            isValid
                ? `Valid trim size: ${trimSize}mm`
                : `Invalid trim size: ${trimSize}mm (must be one of: 140x216, 152x229, 156x234, 170x244, 189x246, 178x254, 203x254, 216x280)`
        );
    }

    validateColourPaperCompatibility(context) {
        const { paper, colour, results } = context;
        
        if (!CONFIG.VALID_PAPERS.has(paper) || !CONFIG.VALID_COLOURS.has(colour)) {
            return;
        }
        
        const compatibility = CONFIG.PAPER_COMPATIBILITY[paper];
        const isValid = compatibility.allowedColours.has(colour);
        
        if (!isValid) {
            const allowedColours = Array.from(compatibility.allowedColours).join(', ');
            this.addResult(
                results,
                'Colour/Paper Compatibility',
                false,
                `Colour '${colour}' is not compatible with ${paper} (allowed: ${allowedColours})`
            );
        } else {
            this.addResult(
                results,
                'Colour/Paper Compatibility',
                true,
                `Colour '${colour}' is compatible with ${paper}`
            );
        }
    }

    validateRoutePaperCompatibility(context) {
        const { paper, quality, results } = context;
        
        if (!CONFIG.VALID_PAPERS.has(paper) || !CONFIG.VALID_ROUTES.has(quality)) {
            return;
        }
        
        const compatibility = CONFIG.PAPER_COMPATIBILITY[paper];
        const isValid = compatibility.allowedRoutes.has(quality);
        
        if (!isValid) {
            const allowedRoutes = Array.from(compatibility.allowedRoutes).join(', ');
            this.addResult(
                results,
                'Route/Paper Compatibility',
                false,
                `Route '${quality}' is not compatible with ${paper} (allowed: ${allowedRoutes})`
            );
        } else {
            this.addResult(
                results,
                'Route/Paper Compatibility',
                true,
                `Route '${quality}' is compatible with ${paper}`
            );
        }
    }
}
