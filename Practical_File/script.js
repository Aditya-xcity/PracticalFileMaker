/**
 * Document Generator - Frontend-only Word Template Automation
 * This script runs entirely in the browser and generates Word documents
 * by replacing placeholders in a template.docx file
 * Using JSZip 2.7.1 for compatibility with docxtemplater
 */

// Global state
let appInitialized = false;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    console.log('📄 Document Generator Initializing...');
    
    // Check libraries are loaded
    setTimeout(initializeApp, 100);
});

function initializeApp() {
    // Check if required libraries are loaded
    if (typeof JSZip === 'undefined') {
        showError('JSZip library not loaded. Please check your internet connection and refresh the page.');
        return;
    }
    
    if (typeof docxtemplater === 'undefined') {
        showError('docxtemplater library not loaded. Please check your internet connection and refresh the page.');
        return;
    }
    
    if (typeof saveAs === 'undefined') {
        showError('FileSaver library not loaded. Please check your internet connection and refresh the page.');
        return;
    }
    
    console.log('✅ All libraries loaded:');
    console.log('JSZip version:', JSZip.version);
    console.log('docxtemplater version:', docxtemplater.version);
    
    // Initialize the application
    initApplication();
    appInitialized = true;
}

function showError(message) {
    console.error('❌', message);
    alert('Error: ' + message);
}

function initApplication() {
    console.log('🚀 Initializing Document Generator Application');
    
    // DOM Elements
    const documentForm = document.getElementById('documentForm');
    const nameInput = document.getElementById('name');
    const rollNoInput = document.getElementById('rollNo');
    const sectionInput = document.getElementById('section');
    const generateBtn = document.getElementById('generateBtn');
    const previewBtn = document.getElementById('previewBtn');
    const previewSection = document.getElementById('previewSection');
    const previewName = document.getElementById('previewName');
    const previewRollNo = document.getElementById('previewRollNo');
    const previewSectionValue = document.getElementById('previewSection');
    const debugPanel = document.getElementById('debugPanel');
    const debugInfo = document.getElementById('debugInfo');
    
    // Show debug info function
    function addDebugInfo(message) {
        if (debugPanel && debugInfo) {
            const timestamp = new Date().toLocaleTimeString();
            debugInfo.innerHTML += `[${timestamp}] ${message}\n`;
            debugPanel.scrollTop = debugPanel.scrollHeight;
        }
        console.log(message);
    }
    
    // Show/hide debug panel
    window.toggleDebug = function() {
        debugPanel.classList.toggle('hidden');
    };
    
    // Clear debug info
    window.clearDebug = function() {
        if (debugInfo) {
            debugInfo.innerHTML = '';
        }
    };
    
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
        
        addDebugInfo('Preview updated with values');
    });
    
    // Function to load template file
    function loadTemplate() {
        addDebugInfo('Loading template.docx...');
        
        return new Promise(function(resolve, reject) {
            // Use XMLHttpRequest for better binary handling
            const xhr = new XMLHttpRequest();
            xhr.open('GET', 'template.docx', true);
            xhr.responseType = 'arraybuffer';
            
            xhr.onload = function() {
                if (xhr.status === 200) {
                    addDebugInfo(`✅ Template loaded: ${xhr.response.byteLength} bytes`);
                    resolve(xhr.response);
                } else {
                    addDebugInfo(`❌ Failed to load template: HTTP ${xhr.status}`);
                    reject(new Error(`Failed to load template: HTTP ${xhr.status}`));
                }
            };
            
            xhr.onerror = function() {
                addDebugInfo('❌ Network error loading template');
                reject(new Error('Network error loading template'));
            };
            
            xhr.send();
        });
    }
    
    // Function to generate document
    async function generateDocument(data) {
        try {
            addDebugInfo('Starting document generation...');
            addDebugInfo(`Data: ${JSON.stringify(data)}`);
            
            // Load the template
            const arrayBuffer = await loadTemplate();
            
            // Convert array buffer to binary string (compatible with JSZip 2.x)
            const bytes = new Uint8Array(arrayBuffer);
            let binaryString = '';
            
            for (let i = 0; i < bytes.length; i++) {
                binaryString += String.fromCharCode(bytes[i]);
            }
            
            addDebugInfo(`Converted to binary string: ${binaryString.length} chars`);
            
            // Create JSZip instance with the content
            // JSZip 2.x accepts content in constructor
            const zip = new JSZip(binaryString);
            
            // Create docxtemplater instance
            const doc = new docxtemplater();
            doc.loadZip(zip);
            
            // Set the data
            doc.setData(data);
            
            // Render the document
            try {
                doc.render();
                addDebugInfo('✅ Template rendered successfully');
            } catch (renderError) {
                addDebugInfo(`❌ Template render error: ${renderError.message}`);
                
                // Provide helpful error message
                let errorMessage = 'Template Error: ' + renderError.message;
                
                if (renderError.properties && renderError.properties.key) {
                    errorMessage += `\nError at placeholder: ${renderError.properties.key}`;
                    
                    // Check if placeholder exists
                    const placeholders = ['name', 'rollNo', 'section'];
                    const missingPlaceholders = placeholders.filter(p => !data[p]);
                    
                    if (missingPlaceholders.length > 0) {
                        errorMessage += `\nMissing data for: ${missingPlaceholders.join(', ')}`;
                    }
                }
                
                throw new Error(errorMessage);
            }
            
            // Generate the output
            addDebugInfo('Generating output file...');
            
            // Generate as blob
            const out = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            
            addDebugInfo(`✅ Document generated: ${out.size} bytes`);
            return out;
            
        } catch (error) {
            addDebugInfo(`❌ Document generation failed: ${error.message}`);
            throw error;
        }
    }
    
    // Form submission handler
    documentForm.addEventListener('submit', async function(e) {
        e.preventDefault();
        
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
        
        // Disable button and show loading state
        const originalText = generateBtn.innerHTML;
        generateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generating...';
        generateBtn.disabled = true;
        previewBtn.disabled = true;
        
        // Show debug panel during generation
        debugPanel.classList.remove('hidden');
        addDebugInfo('=== Starting Document Generation ===');
        
        try {
            // Generate the document
            const startTime = Date.now();
            const docBlob = await generateDocument(templateData);
            const endTime = Date.now();
            
            addDebugInfo(`Generation time: ${endTime - startTime}ms`);
            
            // Create filename
            const safeName = name.replace(/[^\w\s.-]/gi, '_');
            const safeRollNo = rollNo.replace(/[^\w\s.-]/gi, '_');
            const filename = `${safeName}_${safeRollNo}.docx`;
            
            // Save the file
            saveAs(docBlob, filename);
            
            addDebugInfo(`✅ File saved as: ${filename}`);
            
            // Show success message
            setTimeout(function() {
                alert(`✅ Document generated successfully!\n\nFile: ${filename}\n\nCheck your downloads folder.`);
            }, 500);
            
        } catch (error) {
            console.error('Error:', error);
            
            // Show user-friendly error message
            let errorMessage = 'Failed to generate document.\n\n';
            
            if (error.message.includes('Template file') || error.message.includes('Failed to load')) {
                errorMessage += 'Could not load the template file.\n\n';
                errorMessage += 'Please ensure:\n';
                errorMessage += '1. template.docx exists in the same folder\n';
                errorMessage += '2. The file name is exactly "template.docx"\n';
                errorMessage += '3. The file is a valid Word document';
            } else if (error.message.includes('Template Error')) {
                errorMessage += error.message + '\n\n';
                errorMessage += 'Make sure your template contains:\n';
                errorMessage += '• {{name}}\n';
                errorMessage += '• {{rollNo}}\n';
                errorMessage += '• {{section}}';
            } else {
                errorMessage += 'Error: ' + error.message;
            }
            
            alert(errorMessage);
            
        } finally {
            // Restore button state
            generateBtn.innerHTML = originalText;
            generateBtn.disabled = false;
            previewBtn.disabled = false;
            
            addDebugInfo('=== Generation Complete ===\n');
        }
    });
    
    // Real-time validation
    [nameInput, rollNoInput, sectionInput].forEach(input => {
        input.addEventListener('input', function() {
            const allFilled = nameInput.value.trim() && 
                             rollNoInput.value.trim() && 
                             sectionInput.value.trim();
            
            // Enable/disable buttons
            generateBtn.disabled = !allFilled;
            previewBtn.disabled = !allFilled;
            
            // Update preview if section is visible
            if (!previewSection.classList.contains('hidden')) {
                previewBtn.click();
            }
        });
    });
    
    // Initialize button states
    generateBtn.disabled = true;
    previewBtn.disabled = true;
    
    // Test template on startup
    testTemplateAvailability();
    
    addDebugInfo('✅ Application initialized');
    addDebugInfo('Ready to generate documents!');
    
    // Expose some utility functions for debugging
    window.app = {
        testTemplate: testTemplateAvailability,
        generateTest: function() {
            const testData = {
                name: "Test Student",
                rollNo: "TEST001",
                section: "Test Section"
            };
            
            addDebugInfo('Running test generation...');
            return generateDocument(testData);
        },
        getTemplateSize: async function() {
            try {
                const response = await fetch('template.docx');
                const blob = await response.blob();
                return blob.size;
            } catch (error) {
                return 'Error: ' + error.message;
            }
        }
    };
    
    // Function to test template availability
    async function testTemplateAvailability() {
        addDebugInfo('Testing template availability...');
        
        try {
            const response = await fetch('template.docx', { method: 'HEAD' });
            
            if (response.ok) {
                const sizeResponse = await fetch('template.docx');
                const blob = await sizeResponse.blob();
                
                addDebugInfo(`✅ Template found: ${blob.size} bytes`);
                
                // Check if it's a valid DOCX
                const arrayBuffer = await blob.slice(0, 4).arrayBuffer();
                const view = new Uint8Array(arrayBuffer);
                const signature = view[0] === 0x50 && view[1] === 0x4B; // "PK" signature
                
                if (signature) {
                    addDebugInfo('✅ Valid DOCX/ZIP file detected');
                } else {
                    addDebugInfo('⚠️ File may not be a valid DOCX');
                }
                
                return true;
            } else {
                addDebugInfo(`❌ Template not found: HTTP ${response.status}`);
                return false;
            }
        } catch (error) {
            addDebugInfo(`❌ Template check failed: ${error.message}`);
            return false;
        }
    }
}
