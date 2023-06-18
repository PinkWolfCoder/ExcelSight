const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const loadButton = document.getElementById('loadButton');
  const readButton = document.getElementById('readButton');
  const appendForm = document.getElementById('appendForm');
  const exitButton = document.getElementById('exitButton');
  const loadTitle = document.getElementById('loadTitle');

  ipcRenderer.on('load-result', (event, success) => {
    if (success) {
      loadTitle.classList.add('load-success'); // Add success class
      loadTitle.classList.remove('load-failure'); // Remove failure class
    } else {
      loadTitle.classList.add('load-failure'); // Add failure class
      loadTitle.classList.remove('load-success'); // Remove success class
    }
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

  readButton.addEventListener('click', async () => {
    const rowNumber = document.getElementById('readRow').value;
    readButton.disabled = true;  // disable the button
    try {
      const rowData = await ipcRenderer.invoke('read-row', rowNumber);
      const readResult = document.getElementById('readResult');
      const formattedRowData = JSON.stringify(rowData, null, 2);
      console.log('Formatted row data:', formattedRowData);  // New line
      readResult.textContent = formattedRowData;
      console.log('readResult text content:', readResult.textContent);  // New line
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
