/**
 * Document Generator using Mammoth.js
 * Simple and reliable .docx template replacement
 */

class DocumentGenerator {
    constructor() {
        this.state = {
            initialized: false,
            templateLoaded: false,
            templateData: null
        };
        
        this.init();
    }
    
    async init() {
        this.log('🚀 Initializing Document Generator...', 'info');
        
        // Get DOM elements
        this.elements = {
            form: document.getElementById('documentForm'),
            nameInput: document.getElementById('name'),
            rollNoInput: document.getElementById('rollNo'),
            sectionInput: document.getElementById('section'),
            generateBtn: document.getElementById('generateBtn'),
            previewBtn: document.getElementById('previewBtn'),
            testBtn: document.getElementById('testBtn'),
            previewSection: document.getElementById('previewSection'),
            previewName: document.getElementById('previewName'),
            previewRollNo: document.getElementById('previewRollNo'),
            previewSectionValue: document.getElementById('previewSection'),
            statusLog: document.getElementById('statusLog'),
            clearLogBtn: document.getElementById('clearLogBtn'),
            exportLogBtn: document.getElementById('exportLogBtn')
        };
        
        // Initialize event listeners
        this.initEventListeners();
        
        // Check if libraries are loaded
        this.checkLibraries();
        
        // Test template availability
        await this.testTemplate();
        
        this.state.initialized = true;
        this.log('✅ Document Generator initialized successfully', 'success');
    }
    
    log(message, type = 'info') {
        const timestamp = new Date().toLocaleTimeString();
        const logEntry = document.createElement('div');
        logEntry.className = `log-entry ${type}`;
        logEntry.textContent = `[${timestamp}] ${message}`;
        
        if (this.elements.statusLog) {
            this.elements.statusLog.appendChild(logEntry);
            this.elements.statusLog.scrollTop = this.elements.statusLog.scrollHeight;
        }
        
        console.log(`[${type.toUpperCase()}] ${message}`);
    }
    
    checkLibraries() {
        if (typeof mammoth === 'undefined') {
            this.log('❌ Mammoth.js not loaded', 'error');
            return false;
        }
        
        if (typeof JSZip === 'undefined') {
            this.log('❌ JSZip not loaded', 'error');
            return false;
        }
        
        if (typeof saveAs === 'undefined') {
            this.log('❌ FileSaver not loaded', 'error');
            return false;
        }
        
        this.log('✅ All libraries loaded successfully', 'success');
        return true;
    }
    
    initEventListeners() {
        // Preview button
        if (this.elements.previewBtn) {
            this.elements.previewBtn.addEventListener('click', () => this.showPreview());
        }
        
        // Test button
        if (this.elements.testBtn) {
            this.elements.testBtn.addEventListener('click', () => this.testTemplate());
        }
        
        // Form submission
        if (this.elements.form) {
            this.elements.form.addEventListener('submit', (e) => {
                e.preventDefault();
                this.generateDocument();
            });
        }
        
        // Clear log button
        if (this.elements.clearLogBtn) {
            this.elements.clearLogBtn.addEventListener('click', () => {
                if (this.elements.statusLog) {
                    this.elements.statusLog.innerHTML = '';
                    this.log('Log cleared', 'info');
                }
            });
        }
        
        // Export log button
        if (this.elements.exportLogBtn) {
            this.elements.exportLogBtn.addEventListener('click', () => {
                if (this.elements.statusLog) {
                    const logText = this.elements.statusLog.textContent;
                    const blob = new Blob([logText], { type: 'text/plain' });
                    saveAs(blob, 'document_generator_log.txt');
                    this.log('Log exported', 'success');
                }
            });
        }
        
        // Real-time validation
        const inputs = [this.elements.nameInput, this.elements.rollNoInput, this.elements.sectionInput];
        inputs.forEach(input => {
            if (input) {
                input.addEventListener('input', () => this.validateForm());
            }
        });
        
        this.validateForm();
    }
    
    validateForm() {
        const name = this.elements.nameInput?.value.trim() || '';
        const rollNo = this.elements.rollNoInput?.value.trim() || '';
        const section = this.elements.sectionInput?.value.trim() || '';
        
        const allFilled = name && rollNo && section;
        
        if (this.elements.generateBtn) {
            this.elements.generateBtn.disabled = !allFilled;
        }
        
        if (this.elements.previewBtn) {
            this.elements.previewBtn.disabled = !allFilled;
        }
        
        return allFilled;
    }
    
    showPreview() {
        if (!this.validateForm()) {
            alert('Please fill in all fields to preview');
            return;
        }
        
        const name = this.elements.nameInput.value.trim();
        const rollNo = this.elements.rollNoInput.value.trim();
        const section = this.elements.sectionInput.value.trim();
        
        this.elements.previewName.textContent = name;
        this.elements.previewRollNo.textContent = rollNo;
        this.elements.previewSectionValue.textContent = section;
        
        this.elements.previewSection.classList.remove('hidden');
        this.log('Preview updated', 'info');
    }
    
    async testTemplate() {
        this.log('🔍 Testing template availability...', 'info');
        
        try {
            const response = await fetch('template.docx');
            
            if (!response.ok) {
                this.log(`❌ Template not found (HTTP ${response.status})`, 'error');
                this.state.templateLoaded = false;
                return false;
            }
            
            const blob = await response.blob();
            this.log(`✅ Template found: ${blob.size} bytes`, 'success');
            
            // Test if it's a valid .docx
            if (blob.size === 0) {
                this.log('⚠️ Template is empty', 'warning');
                this.state.templateLoaded = false;
                return false;
            }
            
            // Try to read with Mammoth to test
            const arrayBuffer = await blob.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            
            if (result.value) {
                this.log('✅ Template is a valid .docx file', 'success');
                
                // Check for placeholders
                const text = result.value;
                const placeholders = {
                    name: text.includes('{name}'),
                    rollNo: text.includes('{rollNo}'),
                    section: text.includes('{section}')
                };
                
                this.log('Placeholder check:', 'info');
                this.log(`  {name}: ${placeholders.name ? '✅ Found' : '❌ Not found'}`, 
                        placeholders.name ? 'success' : 'error');
                this.log(`  {rollNo}: ${placeholders.rollNo ? '✅ Found' : '❌ Not found'}`, 
                        placeholders.rollNo ? 'success' : 'error');
                this.log(`  {section}: ${placeholders.section ? '✅ Found' : '❌ Not found'}`, 
                        placeholders.section ? 'success' : 'error');
                
                this.state.templateLoaded = true;
                return true;
            } else {
                this.log('⚠️ Could not read template content', 'warning');
                this.state.templateLoaded = false;
                return false;
            }
            
        } catch (error) {
            this.log(`❌ Template test failed: ${error.message}`, 'error');
            this.state.templateLoaded = false;
            return false;
        }
    }
    
    async generateDocument() {
        if (!this.validateForm()) {
            alert('Please fill in all fields');
            return;
        }
        
        if (!this.state.templateLoaded) {
            const shouldContinue = confirm('Template not loaded or invalid. Continue anyway?');
            if (!shouldContinue) return;
        }
        
        this.log('🚀 Starting document generation...', 'info');
        
        const name = this.elements.nameInput.value.trim();
        const rollNo = this.elements.rollNoInput.value.trim();
        const section = this.elements.sectionInput.value.trim();
        
        this.log(`Data: Name="${name}", RollNo="${rollNo}", Section="${section}"`, 'info');
        
        // Disable buttons during generation
        this.setButtonsState(false);
        
        try {
            // Step 1: Load the template
            this.log('📥 Loading template...', 'info');
            const templateResponse = await fetch('template.docx');
            const templateBlob = await templateResponse.blob();
            const templateArrayBuffer = await templateBlob.arrayBuffer();
            
            // Step 2: Convert .docx to a JSZip object
            this.log('🔧 Processing .docx file...', 'info');
            const zip = await JSZip.loadAsync(templateArrayBuffer);
            
            // Step 3: Extract and modify document.xml (main content)
            this.log('✏️ Replacing placeholders...', 'info');
            await this.processDocumentXML(zip, { name, rollNo, section });
            
            // Step 4: Generate the new .docx
            this.log('💾 Generating output file...', 'info');
            const outputBlob = await zip.generateAsync({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            
            // Step 5: Save the file
            const safeName = name.replace(/[^\w\s.-]/gi, '_');
            const safeRollNo = rollNo.replace(/[^\w\s.-]/gi, '_');
            const filename = `${safeName}_${safeRollNo}.docx`;
            
            saveAs(outputBlob, filename);
            
            this.log(`✅ Document generated successfully: ${filename}`, 'success');
            this.log(`📦 File size: ${outputBlob.size} bytes`, 'info');
            
            // Show success message
            setTimeout(() => {
                alert(`✅ Document generated successfully!\n\nFile: ${filename}\n\nCheck your downloads folder.`);
            }, 500);
            
        } catch (error) {
            this.log(`❌ Document generation failed: ${error.message}`, 'error');
            console.error('Generation error:', error);
            
            alert(`Failed to generate document:\n\n${error.message}\n\nCheck the status log for details.`);
            
        } finally {
            // Re-enable buttons
            this.setButtonsState(true);
            this.log('=== Generation complete ===', 'info');
        }
    }
    
    async processDocumentXML(zip, data) {
        try {
            // Get the main document XML file
            const xmlFile = zip.file('word/document.xml');
            if (!xmlFile) {
                throw new Error('Could not find document.xml in the .docx file');
            }
            
            // Read the XML content
            const xmlContent = await xmlFile.async('text');
            
            // Replace placeholders in the XML
            let modifiedContent = xmlContent;
            
            // Replace {name}
            if (xmlContent.includes('{name}')) {
                modifiedContent = modifiedContent.replace(/{name}/g, data.name);
                this.log(`  Replaced {name} with "${data.name}"`, 'success');
            } else {
                this.log(`  ⚠️ {name} not found in template`, 'warning');
            }
            
            // Replace {rollNo}
            if (xmlContent.includes('{rollNo}')) {
                modifiedContent = modifiedContent.replace(/{rollNo}/g, data.rollNo);
                this.log(`  Replaced {rollNo} with "${data.rollNo}"`, 'success');
            } else {
                this.log(`  ⚠️ {rollNo} not found in template`, 'warning');
            }
            
            // Replace {section}
            if (xmlContent.includes('{section}')) {
                modifiedContent = modifiedContent.replace(/{section}/g, data.section);
                this.log(`  Replaced {section} with "${data.section}"`, 'success');
            } else {
                this.log(`  ⚠️ {section} not found in template`, 'warning');
            }
            
            // Update the XML file in the zip
            zip.file('word/document.xml', modifiedContent);
            
            // Also check and update header/footer files if they exist
            await this.processHeaderFooter(zip, data, 'header');
            await this.processHeaderFooter(zip, data, 'footer');
            
            this.log('✅ Placeholders replaced successfully', 'success');
            
        } catch (error) {
            this.log(`❌ Error processing document XML: ${error.message}`, 'error');
            throw error;
        }
    }
    
    async processHeaderFooter(zip, data, type) {
        try {
            // Get all header/footer files
            const files = [];
            zip.forEach((relativePath, file) => {
                if (relativePath.includes(`word/${type}`) && relativePath.endsWith('.xml')) {
                    files.push({ path: relativePath, file: file });
                }
            });
            
            if (files.length > 0) {
                this.log(`Found ${files.length} ${type} file(s)`, 'info');
                
                for (const { path, file } of files) {
                    const content = await file.async('text');
                    let modifiedContent = content;
                    
                    // Replace placeholders
                    modifiedContent = modifiedContent.replace(/{name}/g, data.name);
                    modifiedContent = modifiedContent.replace(/{rollNo}/g, data.rollNo);
                    modifiedContent = modifiedContent.replace(/{section}/g, data.section);
                    
                    // Update the file
                    zip.file(path, modifiedContent);
                    this.log(`  Updated ${path}`, 'success');
                }
            }
        } catch (error) {
            this.log(`⚠️ Could not process ${type} files: ${error.message}`, 'warning');
            // Don't throw error for header/footer issues
        }
    }
    
    setButtonsState(enabled) {
        const buttons = [
            this.elements.generateBtn,
            this.elements.previewBtn,
            this.elements.testBtn
        ];
        
        buttons.forEach(btn => {
            if (btn) {
                if (!enabled) {
                    btn.disabled = true;
                    const originalText = btn.innerHTML;
                    btn.setAttribute('data-original-text', originalText);
                    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
                } else {
                    btn.disabled = false;
                    const originalText = btn.getAttribute('data-original-text');
                    if (originalText) {
                        btn.innerHTML = originalText;
                    }
                }
            }
        });
    }
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.documentGenerator = new DocumentGenerator();
});

// Expose utility functions
window.debugTools = {
    testTemplate: async () => {
        if (window.documentGenerator) {
            return await window.documentGenerator.testTemplate();
        }
        return false;
    },
    
    extractTemplateText: async () => {
        try {
            const response = await fetch('template.docx');
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            console.log('Template text:', result.value);
            return result.value;
        } catch (error) {
            console.error('Error:', error);
            return null;
        }
    },
    
    showAllFiles: async () => {
        try {
            const response = await fetch('template.docx');
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);
            
            console.log('Files in template:');
            zip.forEach((relativePath, file) => {
                console.log(`  ${relativePath} (${file._data.uncompressedSize} bytes)`);
            });
            
            return zip;
        } catch (error) {
            console.error('Error:', error);
            return null;
        }
    }
};
