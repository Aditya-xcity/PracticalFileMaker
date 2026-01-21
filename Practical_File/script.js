/**
 * Document Generator - Frontend-only Word Template Automation
 * This script runs entirely in the browser and generates Word documents
 * by replacing placeholders in a template.docx file
 */

// Wait for libraries to load
window.addEventListener('DOMContentLoaded', () => {
    // Check if required libraries are loaded
    checkLibraries();
});

function checkLibraries() {
    const loadingIndicator = document.getElementById('loadingIndicator');
    
    // Show loading indicator
    loadingIndicator.classList.remove('hidden');
    
    // Check each library with timeout
    const libraries = {
        'JSZip': typeof JSZip,
        'docxtemplater': typeof docxtemplater,
        'FileSaver': typeof saveAs
    };
    
    let checkCount = 0;
    const maxChecks = 50; // 5 seconds max
    
    function checkLibraryStatus() {
        checkCount++;
        
        const allLoaded = Object.values(libraries).every(type => type !== 'undefined');
        
        if (allLoaded) {
            console.log('✅ All libraries loaded successfully');
            console.log('JSZip version:', JSZip.version);
            console.log('docxtemplater version:', docxtemplater.version);
            
            // Hide loading indicator
            loadingIndicator.classList.add('hidden');
            
            // Initialize the application
            initApplication();
            return;
        }
        
        if (checkCount >= maxChecks) {
            console.error('❌ Some libraries failed to load after timeout');
            console.log('Current library status:', libraries);
            
            // Hide loading indicator and show error
            loadingIndicator.classList.add('hidden');
            showLibraryError();
            return;
        }
        
        // Update library status
        libraries.JSZip = typeof JSZip;
        libraries.docxtemplater = typeof docxtemplater;
        libraries.FileSaver = typeof saveAs;
        
        setTimeout(checkLibraryStatus, 100);
    }
    
    checkLibraryStatus();
}

function showLibraryError() {
    alert(
        '❌ Required libraries failed to load.\n\n' +
        'Please:\n' +
        '1. Check your internet connection\n' +
        '2. Refresh the page\n' +
        '3. If problem persists, check browser console (F12)\n\n' +
        'The application needs JSZip, docxtemplater, and FileSaver.js libraries.'
    );
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
    
    // Preview button click handler
    previewBtn.addEventListener('click', () => {
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
    });
    
    // Load the template file using Fetch API
    async function loadTemplateFile() {
        try {
            console.log('📥 Loading template file...');
            
            // Use fetch API
            const response = await fetch('template.docx');
            
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            // Get the file as array buffer
            const arrayBuffer = await response.arrayBuffer();
            
            console.log('✅ Template loaded, size:', arrayBuffer.byteLength, 'bytes');
            return arrayBuffer;
            
        } catch (error) {
            console.error('❌ Failed to load template:', error);
            
            // Provide helpful error messages
            if (error.message.includes('Failed to fetch') || error.message.includes('404')) {
                throw new Error(
                    'Template file (template.docx) not found. Please ensure:\n\n' +
                    '1. The file is named exactly "template.docx"\n' +
                    '2. It is in the same folder as index.html\n' +
                    '3. GitHub Pages is serving it correctly\n\n' +
                    'For GitHub Pages:\n' +
                    '• Make sure template.docx is committed to the repository\n' +
                    '• Check the repository file structure'
                );
            }
            throw error;
        }
    }
    
    // Generate the document with replaced placeholders
    async function generateDocument(data) {
        try {
            console.log('⚙️ Starting document generation...');
            console.log('Data to insert:', data);
            
            // Load the template
            const templateContent = await loadTemplateFile();
            
            if (!templateContent || templateContent.byteLength === 0) {
                throw new Error('Template content is empty or invalid');
            }
            
            console.log('🔧 Creating document from template...');
            
            // Create a JSZip instance with the template content
            const zip = new JSZip(templateContent);
            
            // Create a docxtemplater instance
            const doc = new docxtemplater();
            doc.loadZip(zip);
            
            // Set the data to replace placeholders
            doc.setData(data);
            
            // Render the document (replace all placeholders)
            try {
                doc.render();
                console.log('✅ Template rendered successfully');
            } catch (renderError) {
                console.error('❌ Error rendering template:', renderError);
                
                let errorDetails = renderError.message;
                
                // Check for common template errors
                if (renderError.properties) {
                    const prop = renderError.properties;
                    errorDetails += `\n\nError at: ${prop.key || 'unknown'}`;
                    
                    if (prop.rootError) {
                        errorDetails += `\nRoot error: ${prop.rootError}`;
                    }
                    
                    // Check if placeholders exist in template
                    if (errorDetails.includes('not found')) {
                        errorDetails += '\n\nMake sure your template.docx contains these exact placeholders:';
                        errorDetails += '\n• {{name}}';
                        errorDetails += '\n• {{rollNo}}';
                        errorDetails += '\n• {{section}}';
                    }
                }
                
                throw new Error(`Template error:\n${errorDetails}`);
            }
            
            // Generate the output as a blob
            console.log('💾 Generating output file...');
            const out = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            
            console.log('✅ Document generated, size:', out.size, 'bytes');
            return out;
            
        } catch (error) {
            console.error('❌ Document generation failed:', error);
            throw error;
        }
    }
    
    // Form submission handler
    documentForm.addEventListener('submit', async (e) => {
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
        
        // Prepare data for template replacement
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
        
        try {
            // Generate the document
            console.log('🚀 Generating document...');
            const docBlob = await generateDocument(templateData);
            
            // Create filename
            const safeName = name.replace(/[^\w\s.-]/gi, '_');
            const safeRollNo = rollNo.replace(/[^\w\s.-]/gi, '_');
            const filename = `${safeName}_${safeRollNo}.docx`;
            
            // Save the file using FileSaver.js
            saveAs(docBlob, filename);
            
            console.log('📥 Document downloaded:', filename);
            
            // Show success message
            setTimeout(() => {
                alert(`✅ Document generated successfully!\n\n📄 File: ${filename}\n\n📁 Check your downloads folder.`);
            }, 500);
            
        } catch (error) {
            console.error('❌ Error:', error);
            
            // Show user-friendly error message
            let errorMessage = '❌ Failed to generate document.\n\n';
            
            if (error.message.includes('Template file') || error.message.includes('not found')) {
                errorMessage += 'File Error:\n';
                errorMessage += error.message;
            } else if (error.message.includes('Template error')) {
                errorMessage += 'Template Error:\n';
                errorMessage += error.message;
            } else if (error.message.includes('network') || error.message.includes('fetch')) {
                errorMessage += 'Network Error:\nPlease check your internet connection and try again.';
            } else {
                errorMessage += 'Error:\n' + error.message;
            }
            
            alert(errorMessage);
            
        } finally {
            // Restore button state
            generateBtn.innerHTML = originalText;
            generateBtn.disabled = false;
            previewBtn.disabled = false;
        }
    });
    
    // Real-time validation
    [nameInput, rollNoInput, sectionInput].forEach(input => {
        input.addEventListener('input', () => {
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
    
    console.log('✅ Application initialized successfully');
    console.log('📝 Ready to generate documents!');
    
    // Test template availability on page load
    checkTemplateAvailability();
    
    async function checkTemplateAvailability() {
        try {
            const response = await fetch('template.docx', { method: 'HEAD' });
            if (response.ok) {
                console.log('✅ Template file is accessible');
                
                // Get actual size
                const sizeResponse = await fetch('template.docx');
                const blob = await sizeResponse.blob();
                console.log(`📁 Template size: ${blob.size} bytes (${(blob.size / 1024).toFixed(2)} KB)`);
                
                if (blob.size === 0) {
                    console.warn('⚠️ Warning: Template file is empty (0 bytes)');
                }
            } else {
                console.warn('⚠️ Template file not found (HTTP ' + response.status + ')');
                console.warn('Make sure template.docx exists in the same directory');
            }
        } catch (error) {
            console.warn('⚠️ Could not check template file:', error.message);
        }
    }
    
    // Add debug utility
    window.testTemplate = async function() {
        console.log('🧪 Testing template loading...');
        
        try {
            const response = await fetch('template.docx');
            const blob = await response.blob();
            
            console.log('✅ Template accessible');
            console.log('Size:', blob.size, 'bytes');
            console.log('Type:', blob.type);
            
            // Try to read first few bytes to verify it's a DOCX
            const arrayBuffer = await blob.slice(0, 4).arrayBuffer();
            const view = new Uint8Array(arrayBuffer);
            const signature = Array.from(view).map(b => b.toString(16).padStart(2, '0')).join(' ');
            
            console.log('File signature (first 4 bytes):', signature);
            
            // DOCX should start with PK (ZIP format)
            if (signature.includes('50 4b')) {
                console.log('✅ File appears to be a valid DOCX (ZIP format)');
            } else {
                console.warn('⚠️ File may not be a valid DOCX');
            }
            
            return true;
        } catch (error) {
            console.error('❌ Template test failed:', error);
            return false;
        }
    };
    
    // Test the libraries work
    window.testLibraries = function() {
        console.log('🧪 Testing libraries...');
        
        const tests = {
            'JSZip': typeof JSZip !== 'undefined',
            'docxtemplater': typeof docxtemplater !== 'undefined',
            'FileSaver': typeof saveAs !== 'undefined',
            'JSZip version': JSZip ? JSZip.version : 'Not loaded',
            'docxtemplater version': docxtemplater ? docxtemplater.version : 'Not loaded'
        };
        
        console.table(tests);
        
        return Object.values(tests).every(test => test !== false && test !== 'Not loaded');
    };
    
    // Run initial library test
    console.log('🔧 Library test results:');
    window.testLibraries();
}