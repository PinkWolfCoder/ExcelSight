const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const loadButton = document.getElementById('loadButton');
  const readButton = document.getElementById('readButton');
  const exitButton = document.getElementById('exitButton');
  const loadTitle = document.getElementById('loadTitle');
  const clearButton = document.getElementById('clearButton');

  let container = document.getElementById("jsoneditor");
  let editor = null;

  container.classList.add("hidden");

  function displayRow(rowData) {
    if (editor) {
      editor.destroy();
    }
    container.innerHTML = ''; // Clear previous content
    editor = new JSONEditor(container, {}); // Create a new editor instance
    editor.set(rowData); // Set JSON data to editor
    adjustEditorSize(); // Adjust the size initially
  }


  ipcRenderer.on('load-result', (event, success) => {
    if (success) {
      loadTitle.classList.add('load-success');
      loadTitle.classList.remove('load-failure');

      const rowNumber = 1;
      container.classList.remove("hidden");
    } else {
      loadTitle.classList.add('load-failure');
      loadTitle.classList.remove('load-success');
    }
  });

  ipcRenderer.on('row-data', (event, rowData) => {
      displayRow(rowData);
  });

  function adjustEditorSize() {
    container.style.width = '100%'; // Set width to 100%
    container.style.height = '100%'; // Set height to 100%
    if (editor) {
      editor.resize(); // Resize the editor to fit the updated size
    }
  }

  window.addEventListener('resize', adjustEditorSize);

  loadButton.addEventListener('click', async () => {
    const filePath = document.getElementById('filePath').value;
    loadButton.disabled = true;
    try {
      const response = await ipcRenderer.invoke('load-excel', filePath);
      if (response.success) {
        console.log(response.message);
      } else {
        console.error(response.message);
      }
    } catch (err) {
      console.error(err);
    } finally {
      loadButton.disabled = false;
    }
  });

  readButton.addEventListener('click', () => {
    const rowNumber = document.getElementById('readRow').value;
    container.classList.remove("hidden");

    readButton.disabled = true;
    try {
      ipcRenderer.send('read-row', rowNumber);
    } catch (err) {
      console.error(err);
    } finally {
      readButton.disabled = false;
    }
  });

  clearButton.addEventListener('click', async () => {
    const filePath = document.getElementById('filePath').value;
    clearButton.disabled = true;
    try {
      const response = await ipcRenderer.invoke('on-clear', filePath);
      if (response.success) {
        console.log(response.message);
      } else {
        console.error(response.message);
      }
    } catch (err) {
      console.error(err);
    } finally {
      clearButton.disabled = false;
    }
  });

  exitButton.addEventListener('click', () => {
    ipcRenderer.send('exit-app');
  });
});
