/**
 * Document Generator - Frontend-only Word Template Automation
 * Using PizZip (JSZip fork) for compatibility with docxtemplater
 */

// Global state
const state = {
    isInitialized: false,
    librariesLoaded: false
};

// DOM Elements
let documentForm, nameInput, rollNoInput, sectionInput;
let generateBtn, previewBtn, previewSection;
let previewName, previewRollNo, previewSectionValue;
let statusMessages;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    console.log('📄 Document Generator Initializing...');
    
    // Get DOM elements
    documentForm = document.getElementById('documentForm');
    nameInput = document.getElementById('name');
    rollNoInput = document.getElementById('rollNo');
    sectionInput = document.getElementById('section');
    generateBtn = document.getElementById('generateBtn');
    previewBtn = document.getElementById('previewBtn');
    previewSection = document.getElementById('previewSection');
    previewName = document.getElementById('previewName');
    previewRollNo = document.getElementById('previewRollNo');
    previewSectionValue = document.getElementById('previewSection');
    statusMessages = document.getElementById('statusMessages');
    
    // Check if libraries are loaded
    checkLibraries();
    
    // Initialize form handlers
    initFormHandlers();
    
    // Test template availability
    setTimeout(testTemplateAvailability, 1000);
});

// Check if required libraries are loaded
function checkLibraries() {
    addStatus('🔍 Checking libraries...', 'info');
    
    const checkInterval = setInterval(() => {
        const pizzipLoaded = typeof PizZip !== 'undefined';
        const docxtemplaterLoaded = typeof docxtemplater !== 'undefined';
        const filesaverLoaded = typeof saveAs !== 'undefined';
        
        if (pizzipLoaded && docxtemplaterLoaded && filesaverLoaded) {
            clearInterval(checkInterval);
            state.librariesLoaded = true;
            addStatus('✅ All libraries loaded successfully', 'success');
            console.log('PizZip loaded:', typeof PizZip);
            console.log('docxtemplater loaded:', typeof docxtemplater);
            console.log('FileSaver loaded:', typeof saveAs);
            
            // Mark as initialized
            state.isInitialized = true;
            addStatus('🚀 Application ready to use', 'success');
            
            // Enable form
            enableForm(true);
        }
        
        // Timeout after 10 seconds
        setTimeout(() => {
            if (!state.librariesLoaded) {
                clearInterval(checkInterval);
                addStatus('❌ Failed to load libraries. Please refresh the page.', 'error');
                enableForm(false);
            }
        }, 10000);
    }, 100);
}

// Add status message
function addStatus(message, type = 'info') {
    const timestamp = new Date().toLocaleTimeString();
    const statusDiv = document.createElement('div');
    statusDiv.className = `status-${type}`;
    statusDiv.innerHTML = `[${timestamp}] ${message}`;
    
    if (statusMessages) {
        statusMessages.appendChild(statusDiv);
        statusMessages.scrollTop = statusMessages.scrollHeight;
    }
    
    console.log(`[${type.toUpperCase()}] ${message}`);
}

// Enable/disable form
function enableForm(enabled) {
    if (generateBtn) generateBtn.disabled = !enabled;
    if (previewBtn) previewBtn.disabled = !enabled;
    
    if (enabled) {
        addStatus('✅ Form enabled - ready to generate documents', 'success');
    } else {
        addStatus('⚠️ Form disabled - waiting for libraries', 'warning');
    }
}

// Initialize form handlers
function initFormHandlers() {
    if (!documentForm) return;
    
    // Preview button click handler
    previewBtn.addEventListener('click', function() {
        const name = nameInput.value.trim();
        const rollNo = rollNoInput.value.trim();
        const section = sectionInput.value.trim();
        
        if (!name || !rollNo || !section) {
            alert('Please fill in all fields to preview');
            return;
        }
        
        // Update preview values
        previewName.textContent = name;
        previewRollNo.textContent = rollNo;
        previewSectionValue.textContent = section;
        
        // Show preview section
        previewSection.classList.remove('hidden');
        
        addStatus('📊 Preview updated', 'info');
    });
    
    // Form submission handler
    documentForm.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        if (!state.librariesLoaded) {
            addStatus('❌ Libraries not loaded yet. Please wait.', 'error');
            return;
        }
        
        // Get form values
        const name = nameInput.value.trim();
        const rollNo = rollNoInput.value.trim();
        const section = sectionInput.value.trim();
        
        // Validate inputs
        if (!name || !rollNo || !section) {
            alert('Please fill in all fields');
            return;
        }
        
        // Prepare data
        const templateData = {
            name: name,
            rollNo: rollNo,
            section: section
        };
        
        // Generate document
        await generateAndDownloadDocument(templateData);
    });
    
    // Real-time validation
    [nameInput, rollNoInput, sectionInput].forEach(input => {
        input.addEventListener('input', function() {
            const allFilled = nameInput.value.trim() && 
                             rollNoInput.value.trim() && 
                             sectionInput.value.trim();
            
            // Enable/disable buttons
            if (generateBtn) generateBtn.disabled = !allFilled || !state.librariesLoaded;
            if (previewBtn) previewBtn.disabled = !allFilled;
            
            // Update preview if section is visible
            if (previewSection && !previewSection.classList.contains('hidden')) {
                previewBtn.click();
            }
        });
    });
    
    // Initialize button states
    enableForm(false);
}

// Load template file
async function loadTemplateFile() {
    addStatus('📥 Loading template.docx...', 'info');
    
    try {
        // Use fetch API
        const response = await fetch('template.docx');
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        // Get as array buffer
        const arrayBuffer = await response.arrayBuffer();
        
        addStatus(`✅ Template loaded: ${arrayBuffer.byteLength} bytes`, 'success');
        return arrayBuffer;
        
    } catch (error) {
        addStatus(`❌ Failed to load template: ${error.message}`, 'error');
        throw error;
    }
}

// Generate and download document
async function generateAndDownloadDocument(data) {
    if (!state.librariesLoaded) {
        addStatus('❌ Libraries not loaded', 'error');
        return;
    }
    
    // Disable button and show loading state
    const originalText = generateBtn.innerHTML;
    generateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
    generateBtn.disabled = true;
    previewBtn.disabled = true;
    
    addStatus('🚀 Starting document generation...', 'info');
    addStatus(`Data: ${JSON.stringify(data)}`, 'info');
    
    try {
        // Load template
        const templateArrayBuffer = await loadTemplateFile();
        
        if (!templateArrayBuffer || templateArrayBuffer.byteLength === 0) {
            throw new Error('Template is empty');
        }
        
        addStatus('🔧 Creating document from template...', 'info');
        
        // Convert array buffer to binary string for PizZip
        const bytes = new Uint8Array(templateArrayBuffer);
        let binaryString = '';
        
        for (let i = 0; i < bytes.length; i++) {
            binaryString += String.fromCharCode(bytes[i]);
        }
        
        addStatus(`Converted to binary string: ${binaryString.length} characters`, 'info');
        
        // Create PizZip instance with the content
        // PizZip accepts content in constructor (unlike JSZip 3.0)
        const zip = new PizZip(binaryString);
        
        // Create docxtemplater instance
        const doc = new docxtemplater();
        doc.loadZip(zip);
        
        // Set the data
        doc.setData(data);
        
        // Render the document
        try {
            doc.render();
            addStatus('✅ Template rendered successfully', 'success');
        } catch (renderError) {
            addStatus(`❌ Template render error: ${renderError.message}`, 'error');
            
            let errorDetails = 'Template Error:\n' + renderError.message;
            
            if (renderError.properties && renderError.properties.key) {
                errorDetails += `\nAt placeholder: ${renderError.properties.key}`;
                
                // Check for common issues
                if (renderError.message.includes('not found')) {
                    errorDetails += '\n\nMake sure your template contains exactly:';
                    errorDetails += '\n• {{name}}';
                    errorDetails += '\n• {{rollNo}}';
                    errorDetails += '\n• {{section}}';
                }
            }
            
            throw new Error(errorDetails);
        }
        
        // Generate the output
        addStatus('💾 Generating output file...', 'info');
        
        const out = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });
        
        addStatus(`✅ Document generated: ${out.size} bytes`, 'success');
        
        // Create filename
        const safeName = data.name.replace(/[^\w\s.-]/gi, '_');
        const safeRollNo = data.rollNo.replace(/[^\w\s.-]/gi, '_');
        const filename = `${safeName}_${safeRollNo}.docx`;
        
        // Save the file
        saveAs(out, filename);
        
        addStatus(`📥 File saved as: ${filename}`, 'success');
        
        // Show success message
        setTimeout(() => {
            alert(`✅ Document generated successfully!\n\nFile: ${filename}\n\nCheck your downloads folder.`);
        }, 500);
        
    } catch (error) {
        console.error('Generation error:', error);
        
        // Show user-friendly error message
        let errorMessage = 'Failed to generate document.\n\n';
        
        if (error.message.includes('Template file') || error.message.includes('Failed to load') || error.message.includes('HTTP')) {
            errorMessage += 'Could not load the template file.\n\n';
            errorMessage += 'Please ensure:\n';
            errorMessage += '1. template.docx exists in the same folder\n';
            errorMessage += '2. The file name is exactly "template.docx"\n';
            errorMessage += '3. The file is a valid Word document (.docx)';
        } else if (error.message.includes('Template Error')) {
            errorMessage += error.message;
        } else if (error.message.includes('PizZip') || error.message.includes('constructor')) {
            errorMessage += 'Library compatibility issue.\n\n';
            errorMessage += 'Please try:\n';
            errorMessage += '1. Refreshing the page\n';
            errorMessage += '2. Using a different browser\n';
            errorMessage += '3. Checking browser console (F12) for details';
        } else {
            errorMessage += 'Error: ' + error.message;
        }
        
        alert(errorMessage);
        addStatus(`❌ Error: ${error.message}`, 'error');
        
    } finally {
        // Restore button state
        generateBtn.innerHTML = originalText;
        generateBtn.disabled = !state.librariesLoaded;
        previewBtn.disabled = false;
        
        addStatus('=== Generation complete ===', 'info');
    }
}

// Test template availability
async function testTemplateAvailability() {
    addStatus('🔍 Testing template availability...', 'info');
    
    try {
        const response = await fetch('template.docx', { method: 'HEAD' });
        
        if (response.ok) {
            // Get file size
            const sizeResponse = await fetch('template.docx');
            const blob = await sizeResponse.blob();
            
            addStatus(`✅ Template found: ${blob.size} bytes (${(blob.size / 1024).toFixed(2)} KB)`, 'success');
            
            // Check if it's a valid DOCX (ZIP file)
            const arrayBuffer = await blob.slice(0, 4).arrayBuffer();
            const view = new Uint8Array(arrayBuffer);
            
            // DOCX files start with "PK" (ZIP signature)
            const isZipFile = view[0] === 0x50 && view[1] === 0x4B;
            
            if (isZipFile) {
                addStatus('✅ Valid DOCX/ZIP file detected', 'success');
            } else {
                addStatus('⚠️ File may not be a valid DOCX (missing ZIP signature)', 'warning');
            }
            
            return true;
        } else {
            addStatus(`❌ Template not found (HTTP ${response.status})`, 'error');
            addStatus('Please ensure template.docx exists in the same folder as index.html', 'warning');
            return false;
        }
    } catch (error) {
        addStatus(`❌ Template check failed: ${error.message}`, 'error');
        return false;
    }
}

// Expose utility functions for debugging
window.debugTools = {
    testLibraries: function() {
        return {
            PizZip: typeof PizZip !== 'undefined',
            docxtemplater: typeof docxtemplater !== 'undefined',
            FileSaver: typeof saveAs !== 'undefined',
            state: state
        };
    },
    
    testTemplate: testTemplateAvailability,
    
    createTestDocument: async function() {
        const testData = {
            name: "Test Student",
            rollNo: "TEST001",
            section: "Test Section"
        };
        
        addStatus('🧪 Creating test document...', 'info');
        return generateAndDownloadDocument(testData);
    },
    
    clearStatus: function() {
        if (statusMessages) {
            statusMessages.innerHTML = '';
            addStatus('Status cleared', 'info');
        }
    }
};

// Add keyboard shortcut for debugging (Ctrl+Shift+D)
document.addEventListener('keydown', function(e) {
    if (e.ctrlKey && e.shiftKey && e.key === 'D') {
        e.preventDefault();
        const debugInfo = window.debugTools.testLibraries();
        console.table(debugInfo);
        addStatus('Debug info logged to console (F12)', 'info');
    }
});

// Initial status
addStatus('Document Generator starting up...', 'info');
addStatus('Using PizZip + docxtemplater + FileSaver.js', 'info');
