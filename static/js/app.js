// JavaScript para o Dashboard de Comparação de Nomes
class ComparisonDashboard {
    constructor() {
        this.currentResults = null;
        this.init();
    }

    init() {
        this.bindEvents();
        this.updateThresholdDisplay();
        this.setupFileUploadAreas();
    }

    setupFileUploadAreas() {
        // Setup upload areas for drag and drop and click functionality
        this.setupUploadArea('upload1', 'file1');
        this.setupUploadArea('upload2', 'file2');
    }

    setupUploadArea(uploadAreaId, fileInputId) {
        const uploadArea = document.getElementById(uploadAreaId);
        const fileInput = document.getElementById(fileInputId);

        if (!uploadArea || !fileInput) return;

        // Click to upload
        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });

        // File input change
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.displaySelectedFile(uploadAreaId, fileInputId, e.target.files[0]);
            }
        });

        // Drag and drop events
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                this.displaySelectedFile(uploadAreaId, fileInputId, files[0]);
            }
        });
    }

    displaySelectedFile(uploadAreaId, fileInputId, file) {
        const uploadArea = document.getElementById(uploadAreaId);
        const fileNumber = fileInputId.replace('file', '');
        const infoDiv = document.getElementById(`info${fileNumber}`);

        // Hide upload area and show file info
        uploadArea.style.display = 'none';
        infoDiv.style.display = 'block';

        // Update file info
        const fileName = infoDiv.querySelector('.file-name');
        const fileDetails = infoDiv.querySelector('.file-details');
        
        fileName.textContent = file.name;
        fileDetails.textContent = `Tamanho: ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    }

    removeFile(fileNumber) {
        const uploadArea = document.getElementById(`upload${fileNumber}`);
        const infoDiv = document.getElementById(`info${fileNumber}`);
        const fileInput = document.getElementById(`file${fileNumber}`);

        // Reset file input
        fileInput.value = '';

        // Show upload area and hide file info
        uploadArea.style.display = 'block';
        infoDiv.style.display = 'none';
    }

    bindEvents() {
        // Upload form
        document.getElementById('uploadForm').addEventListener('submit', (e) => {
            e.preventDefault();
            this.handleFileUpload();
        });

        // Compare button
        document.getElementById('compareBtn').addEventListener('click', () => {
            this.handleComparison();
        });

        // Export button
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.handleExport();
        });

        // Threshold slider
        document.getElementById('threshold').addEventListener('input', (e) => {
            this.updateThresholdDisplay();
        });

        // Column selects
        document.getElementById('column1').addEventListener('change', () => {
            this.checkCompareButtonState();
        });

        document.getElementById('column2').addEventListener('change', () => {
            this.checkCompareButtonState();
        });
    }

    updateThresholdDisplay() {
        const threshold = document.getElementById('threshold').value;
        document.getElementById('thresholdValue').textContent = threshold;
    }

    checkCompareButtonState() {
        const file1 = document.getElementById('file1').files.length > 0;
        const file2 = document.getElementById('file2').files.length > 0;
        
        const compareBtn = document.getElementById('compareBtn');
        compareBtn.disabled = !(file1 && file2);
    }

    async handleFileUpload() {
        const formData = new FormData();
        const file1 = document.getElementById('file1').files[0];
        const file2 = document.getElementById('file2').files[0];

        if (!file1 || !file2) {
            this.showToast('Por favor, selecione ambos os arquivos.', 'error');
            return;
        }

        formData.append('file1', file1);
        formData.append('file2', file2);

        try {
            this.showLoading('Fazendo upload dos arquivos...');
            
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.displayFileInfo(result);
                this.showToast('Arquivos carregados com sucesso!', 'success');
            } else {
                this.showToast(result.error || 'Erro no upload dos arquivos', 'error');
            }
        } catch (error) {
            this.showToast('Erro de conexão: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    displayFileInfo(data) {
        // File 1 info
        document.getElementById('file1Name').textContent = data.file1_info.name;
        document.getElementById('file1Rows').textContent = data.file1_info.rows.toLocaleString();
        
        const col1Select = document.getElementById('column1');
        col1Select.innerHTML = '<option value="">Selecione uma coluna...</option>';
        data.file1_info.columns.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            col1Select.appendChild(option);
        });

        // File 2 info
        document.getElementById('file2Name').textContent = data.file2_info.name;
        document.getElementById('file2Rows').textContent = data.file2_info.rows.toLocaleString();
        
        const col2Select = document.getElementById('column2');
        col2Select.innerHTML = '<option value="">Selecione uma coluna...</option>';
        data.file2_info.columns.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            col2Select.appendChild(option);
        });

        // Show sections
        document.getElementById('fileInfoSection').style.display = 'block';
        document.getElementById('configSection').style.display = 'block';
        
        // Add fade-in animation
        document.getElementById('fileInfoSection').classList.add('fade-in');
        document.getElementById('configSection').classList.add('fade-in');
    }

    async handleComparison() {
        const threshold = document.getElementById('threshold').value;
        const algorithm = document.getElementById('algorithm').value;

        const requestData = {
            threshold: parseInt(threshold),
            algorithm: algorithm
        };

        try {
            this.showLoading('Executando comparação...');
            document.getElementById('loadingSection').style.display = 'block';
            
            const response = await fetch('/compare', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(requestData)
            });

            const result = await response.json();

            if (result.success) {
                this.currentResults = result;
                this.displayResults(result);
                this.showToast('Comparação concluída com sucesso!', 'success');
            } else {
                this.showToast(result.error || 'Erro na comparação', 'error');
            }
        } catch (error) {
            this.showToast('Erro de conexão: ' + error.message, 'error');
        } finally {
            this.hideLoading();
            document.getElementById('loadingSection').style.display = 'none';
        }
    }

    displayResults(data) {
        // Update statistics
        document.getElementById('totalToefl').textContent = data.statistics.total_toefl.toLocaleString();
        document.getElementById('matchedCount').textContent = data.statistics.matched.toLocaleString();
        document.getElementById('unmatchedCount').textContent = data.statistics.unmatched.toLocaleString();
        document.getElementById('matchPercentage').textContent = data.statistics.match_percentage + '%';

        // Update results table
        const tbody = document.getElementById('resultsTableBody');
        tbody.innerHTML = '';

        if (data.results.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center text-muted">Nenhuma correspondência encontrada</td></tr>';
            return;
        }

        data.results.forEach((result, index) => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${index + 1}</td>
                <td>${result.toefl_name}</td>
                <td>${result.matched_name}</td>
                <td>${result.class || 'N/A'}</td>
                <td>
                    <span class="badge bg-success">${result.score}%</span>
                </td>
            `;
            tbody.appendChild(row);
        });

        // Show results section
        document.getElementById('resultsSection').style.display = 'block';
        document.getElementById('resultsSection').classList.add('fade-in');
        
        // Scroll to results
        document.getElementById('resultsSection').scrollIntoView({ 
            behavior: 'smooth', 
            block: 'start' 
        });
    }

    getScoreClass(score) {
        if (score >= 90) return 'match-high';
        if (score >= 70) return 'match-medium';
        return 'match-low';
    }

    async handleExport() {
        if (!this.currentResults) {
            this.showToast('Nenhum resultado para exportar.', 'error');
            return;
        }

        try {
            this.showLoading('Preparando exportação...');
            
            const response = await fetch('/export', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    results: this.currentResults.results
                })
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `comparacao_nomes_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                this.showToast('Arquivo exportado com sucesso!', 'success');
            } else {
                const error = await response.json();
                this.showToast(error.error || 'Erro na exportação', 'error');
            }
        } catch (error) {
            this.showToast('Erro de conexão: ' + error.message, 'error');
        } finally {
            this.hideLoading();
        }
    }

    showLoading(message = 'Carregando...') {
        // You can implement a loading overlay here if needed
        console.log('Loading:', message);
    }

    hideLoading() {
        // Hide loading overlay
        console.log('Loading finished');
    }

    showToast(message, type = 'info') {
        const toast = document.getElementById('toast');
        const toastBody = document.getElementById('toastBody');
        
        // Set message
        toastBody.textContent = message;
        
        // Set toast type
        toast.className = 'toast';
        if (type === 'success') {
            toast.classList.add('bg-success', 'text-white');
        } else if (type === 'error') {
            toast.classList.add('bg-danger', 'text-white');
        } else if (type === 'warning') {
            toast.classList.add('bg-warning', 'text-dark');
        } else {
            toast.classList.add('bg-info', 'text-white');
        }
        
        // Show toast
        const bsToast = new bootstrap.Toast(toast);
        bsToast.show();
        
        // Reset classes after hiding
        toast.addEventListener('hidden.bs.toast', () => {
            toast.className = 'toast';
        });
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// Initialize dashboard when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.dashboard = new ComparisonDashboard();
});