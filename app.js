let coilData = [];
let loadedFileName = '';

// Use cache-busting version of GitHub Excel file
const GITHUB_EXCEL_URL = 'https://raw.githubusercontent.com/gmanand23/Coil_Info/main/coil-data.xlsx?' + new Date().getTime();

function clearLocalStorage() {
  try {
    localStorage.removeItem('coilData');
    localStorage.removeItem('loadedFileName');
    alert('Local storage cleared successfully!');
    console.log('Local storage cleared.');
    fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
  } catch (e) {
    console.error('Error clearing local storage:', e);
    alert('Failed to clear local storage. Check console for details.');
  }
}

async function fetchAndLoadExcelFromUrl(url) {
  const downloadButton = document.querySelector('button[onclick="downloadExcel()"]');
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.textContent = 'Loading Data...';
  }

  try {
    document.getElementById('fileName').textContent = `Loading data from GitHub...`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    coilData = XLSX.utils.sheet_to_json(sheet);

    const urlParts = url.split('/');
    loadedFileName = urlParts[urlParts.length - 1].split('?')[0] || 'coil-data.xlsx';

    document.getElementById('fileName').textContent = `Loaded File (from GitHub): ${loadedFileName}`;
    alert('Excel loaded successfully from GitHub!');

    if (downloadButton) {
      downloadButton.disabled = false;
      downloadButton.textContent = 'Download Loaded Excel';
    }

  } catch (error) {
    console.error('Error loading Excel from URL:', error);
    document.getElementById('fileName').textContent = `Failed to load from GitHub.`;
    alert('Failed to load Excel from GitHub. Check console for details.');
    coilData = [];
    if (downloadButton) {
      downloadButton.disabled = true;
      downloadButton.textContent = 'Failed to Load Data';
    }
  }
}

document.getElementById('excelFile').addEventListener('change', (e) => {
  coilData = [];
  document.getElementById('result').innerHTML = '';
  document.getElementById('coilInput').value = '';
  document.getElementById('suggestions').style.display = 'none';

  const downloadButton = document.querySelector('button[onclick="downloadExcel()"]');
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.textContent = 'Loading Local File...';
  }

  const file = e.target.files[0];
  if (file) {
    loadedFileName = file.name;
    document.getElementById('fileName').textContent = `Loaded File (from local upload): ${loadedFileName}`;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      coilData = XLSX.utils.sheet_to_json(sheet);
      alert('Excel loaded successfully from local file!');

      if (downloadButton) {
        downloadButton.disabled = false;
        downloadButton.textContent = 'Download Loaded Excel';
      }
    };
    reader.onerror = (evt) => {
      console.error('Error reading local file:', evt);
      alert('Error reading local Excel file.');
      if (downloadButton) {
        downloadButton.disabled = true;
        downloadButton.textContent = 'Failed to Load Data';
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    if (downloadButton) {
      downloadButton.disabled = true;
      downloadButton.textContent = 'Download Loaded Excel';
    }
  }
});

function downloadExcel() {
  if (coilData.length === 0) {
    alert('No Excel data to download.');
    return;
  }

  if (!loadedFileName) {
    loadedFileName = 'downloaded_coil_data.xlsx';
  }

  const worksheet = XLSX.utils.json_to_sheet(coilData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Coil Data");
  XLSX.writeFile(workbook, loadedFileName);
}

document.addEventListener('DOMContentLoaded', async () => {
  await fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
});

function searchCoil() {
  const coilNumber = document.getElementById('coilInput').value.trim().toUpperCase();
  const result = coilData.find(row => {
    const keys = Object.keys(row);
    const matchingKey = keys.find(k => k.trim().toUpperCase() === 'MILL COIL NO');
    if (!matchingKey) return false;
    const sheetCoil = String(row[matchingKey]).trim().toUpperCase();
    return sheetCoil === coilNumber;
  });
  displayResult(result);
}

function displayResult(data) {
  const resultDiv = document.getElementById('result');
  if (data) {
    let tableHTML = '<table style="font-family: Comic Sans MS; width:100%; border-collapse: collapse;">';
    tableHTML += '<thead><tr><th style="border: 1px solid #fff; padding: 8px; color: white;">Field</th><th style="border: 1px solid #fff; padding: 8px; color: white;">Value</th></tr></thead><tbody>';
    for (const [key, val] of Object.entries(data)) {
      tableHTML += `<tr><td style="border: 1px solid #fff; padding: 8px; color: white;">${key.trim()}</td><td style="border: 1px solid #fff; padding: 8px; color: white;">${val}</td></tr>`;
    }
    tableHTML += '</tbody></table>';
    resultDiv.innerHTML = tableHTML;
  } else {
    resultDiv.innerHTML = '<p>Coil number not found.</p>';
  }
}

let qrReader = null;

function startScanner() {
  if (qrReader) closeScanner();
  const readerDiv = document.getElementById("reader");
  readerDiv.innerHTML = '';
  qrReader = new Html5Qrcode("reader", { verbose: false });
  qrReader.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: { width: 250, height: 250 } },
    (decodedText) => {
      document.getElementById('coilInput').value = decodedText;
      toggleClearButton(); // Ensure clear button shows if text is scanned
      searchCoil();
      closeScanner();
      scrollToResult(); // Scroll to the result after scan and search
    },
    (errorMessage) => {
      console.warn(`QR error: ${errorMessage}`);
    }
  );
}

function closeScanner() {
  if (qrReader) {
    qrReader.stop().then(() => {
      qrReader.clear();
      document.getElementById("reader").innerHTML = '';
      qrReader = null;
    }).catch(err => console.error("Error stopping scanner", err));
  }
}

function showSuggestions() {
  const input = document.getElementById('coilInput').value.trim().toUpperCase();
  const suggestionsDiv = document.getElementById('suggestions');
  suggestionsDiv.innerHTML = '';

  if (!input || coilData.length === 0) {
    suggestionsDiv.style.display = 'none';
    return;
  }

  const keys = Object.keys(coilData[0]);
  const matchingKey = keys.find(k => k.trim().toUpperCase() === 'MILL COIL NO');
  if (!matchingKey) return;

  const suggestions = coilData
    .map(row => String(row[matchingKey]).trim().toUpperCase())
    .filter(coil => coil.includes(input))
    .slice(0, 10);

  if (suggestions.length > 0) {
    suggestions.forEach(s => {
      const div = document.createElement('div');
      div.textContent = s;
      div.onclick = () => {
        document.getElementById('coilInput').value = s;
        suggestionsDiv.style.display = 'none';
        searchCoil();
        toggleClearButton(); // Ensure clear button shows when a suggestion is clicked
      };
      suggestionsDiv.appendChild(div);
    });
    suggestionsDiv.style.display = 'block';
  } else {
    suggestionsDiv.style.display = 'none';
  }
}

function toggleClearButton() {
  const input = document.getElementById('coilInput');
  const clearBtn = document.getElementById('clearInputBtn');
  if (clearBtn) {
    clearBtn.style.display = input.value.trim() !== '' ? 'block' : 'none';
  }
}

function clearInput() {
  document.getElementById('coilInput').value = '';
  document.getElementById('result').innerHTML = '';
  document.getElementById('suggestions').style.display = 'none';
  const clearBtn = document.getElementById('clearInputBtn');
  if (clearBtn) clearBtn.style.display = 'none';
}

function scrollToResult() {
  const resultDiv = document.getElementById('result');
  if (resultDiv) {
    resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}