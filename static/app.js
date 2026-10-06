document.addEventListener('DOMContentLoaded', () => {
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const fileNameDisplay = document.getElementById('fileName');
    const fileBadge = document.getElementById('fileBadge');
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const loading = document.getElementById('loading');
    const selectAllBtn = document.getElementById('selectAllBtn');
    const deselectAllBtn = document.getElementById('deselectAllBtn');

    // Click to select
    dropzone.addEventListener('click', (e) => {
        if (e.target !== fileInput) {
            fileInput.click();
        }
    });

    // File selection
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            showFile(e.target.files[0].name);
        }
    });

    // Drag and drop
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('bg-blue-50', 'border-blue-500');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('bg-blue-50', 'border-blue-500');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('bg-blue-50', 'border-blue-500');
        
        if (e.dataTransfer.files.length > 0) {
            fileInput.files = e.dataTransfer.files;
            showFile(e.dataTransfer.files[0].name);
        }
    });

    function showFile(name) {
        fileNameDisplay.textContent = name;
        fileBadge.classList.remove('hidden');
    }

    // Select/Deselect all checkboxes
    if (selectAllBtn) {
        selectAllBtn.addEventListener('click', () => {
            form.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
        });
    }
    if (deselectAllBtn) {
        deselectAllBtn.addEventListener('click', () => {
            form.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
        });
    }

    // Toggle modern typography options
    const modernOptions = document.getElementById('modernOptions');
    const decreeBox = form.querySelector('input[name="decree_30"]');
    form.querySelectorAll('input[name="style_mode"]').forEach(radio => {
        radio.addEventListener('change', () => {
            const isModern = form.querySelector('input[name="style_mode"]:checked').value === 'modern';
            modernOptions.classList.toggle('hidden', !isModern);
            if (decreeBox) {
                decreeBox.disabled = isModern;
                decreeBox.closest('label').classList.toggle('opacity-50', isModern);
            }
        });
    });

    // Form submission
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        if (!fileInput.files || fileInput.files.length === 0) {
            alert('Vui lòng chọn một file Word (.docx) để xử lý.');
            return;
        }

        const formData = new FormData(form);
        
        submitBtn.classList.add('hidden');
        loading.classList.remove('hidden');

        try {
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'standardized_' + fileInput.files[0].name;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
            } else {
                const errorData = await response.json().catch(() => ({}));
                alert('Có lỗi xảy ra: ' + (errorData.detail || 'Không thể xử lý file.'));
            }
        } catch (error) {
            console.error('Error:', error);
            alert('Lỗi kết nối tới máy chủ ứng dụng.');
        } finally {
            submitBtn.classList.remove('hidden');
            loading.classList.add('hidden');
        }
    });
});
