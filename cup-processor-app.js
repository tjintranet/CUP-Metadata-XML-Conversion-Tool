// Main Application Initialization
let converterModule;
let validatorModule;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    // Initialize modules
    converterModule = new ConverterModule();
    validatorModule = new ValidatorModule();
    
    // Make validator module globally accessible for converter
    window.validatorModule = validatorModule;
    
    console.log('CUP Metadata Processor initialized successfully');
});
