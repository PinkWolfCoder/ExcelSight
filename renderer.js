const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const loadButton = document.getElementById('loadButton');
  const readButton = document.getElementById('readButton');
  const exitButton = document.getElementById('exitButton');
  const loadTitle = document.getElementById('loadTitle');

  let container = document.getElementById("jsoneditor");
  let editor = new JSONEditor(container);
  container.classList.add("hidden");

  ipcRenderer.on('load-result', (event, success) => {
    if (success) {
      loadTitle.classList.add('load-success'); // Add success class
      loadTitle.classList.remove('load-failure'); // Remove failure class
    } else {
      loadTitle.classList.add('load-failure'); // Add failure class
      loadTitle.classList.remove('load-success'); // Remove success class
    }
  });

  ipcRenderer.on('row-data', (event, rowData) => {
    editor.set(rowData); // Set JSON data to editor
  });
  

  loadButton.addEventListener('click', async () => {
    const filePath = document.getElementById('filePath').value;
    loadButton.disabled = true;  // disable the button
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
      loadButton.disabled = false;  // enable the button again
    }
  });

  readButton.addEventListener('click', () => {
    const rowNumber = document.getElementById('readRow').value;
    const container = document.getElementById("jsoneditor");
    container.classList.remove("hidden");

    readButton.disabled = true;  // disable the button
    try {
      ipcRenderer.send('read-row', rowNumber);
    } catch (err) {
      console.error(err);
    } finally {
      readButton.disabled = false;  // enable the button again
    }
  });

  exitButton.addEventListener('click', () => {
    ipcRenderer.send('exit-app');
  });
});
