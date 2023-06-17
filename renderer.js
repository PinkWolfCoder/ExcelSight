// renderer.js
const { ipcRenderer } = require('electron');

document.addEventListener('DOMContentLoaded', () => {
  const loadButton = document.getElementById('loadButton');
  const readButton = document.getElementById('readButton');
  const appendForm = document.getElementById('appendForm');
  const exitButton = document.getElementById('exitButton');

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
      readResult.textContent = JSON.stringify(rowData, null, 2);
    } catch (err) {
      console.error(err);
    } finally {
      readButton.disabled = false;  // enable the button again
    }
  });

  appendForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const col1 = document.getElementById('col1').value;
    const col2 = document.getElementById('col2').value;
    const col3 = document.getElementById('col3').value;
    const appendButton = appendForm.querySelector('button');
    appendButton.disabled = true;  // disable the button
    try {
      const response = await ipcRenderer.invoke('append-row', { col1, col2, col3 });
      if (response.success) {
        console.log(response.message);
      } else {
        console.error(response.message);
      }
    } catch (err) {
      console.error(err);
    } finally {
      appendButton.disabled = false;  // enable the button again
    }
  });

  exitButton.addEventListener('click', () => {
    ipcRenderer.send('exit-app');
  });
});
