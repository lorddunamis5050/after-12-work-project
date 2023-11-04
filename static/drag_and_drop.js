const fileInput = document.getElementById('file-input');
const fileDropArea = document.getElementById('file-drop-area');

fileDropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    fileDropArea.classList.add('file-input-container-hover');
});

fileDropArea.addEventListener('dragleave', () => {
    fileDropArea.classList.remove('file-input-container-hover');
});

fileDropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    fileDropArea.classList.remove('file-input-container-hover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        fileInput.files = files;
    }
});