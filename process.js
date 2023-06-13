function handleFile() {
  const input = document.getElementById('excel-file');
  const file = input.files[0];

  if (!file) {
    alert('Please choose a file to upload');
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const formattedData = [];

    for (let i = 0; i < jsonData.length; i += 13) {
      let obj = {};
      for (let j = 0; j < 12; j++) {
        for (let k = 0; k < jsonData[i + j].length; k += 2) {
          let key = jsonData[i + j][k];
          let value = jsonData[i + j][k + 1];
          obj[key] = value;
        }
      }
      formattedData.push(obj);
    }

    displayData(formattedData);
  };

  reader.readAsArrayBuffer(file);
}

function displayData(data) {
  const outputDiv = document.getElementById('output');
  // Clear the output div
  outputDiv.innerHTML = '';

  data.forEach((entry, index) => {
    // Create a container for each entry
    const entryDiv = document.createElement('div');
    entryDiv.classList.add('entry');

    const title = document.createElement('h2');
    title.textContent = `Entry ${index + 1}`;
    entryDiv.appendChild(title);

    // Create a div for each key-value pair
    for (let key in entry) {
      const div = document.createElement('div');
      div.classList.add('key-value');

      const keySpan = document.createElement('span');
      keySpan.textContent = key;
      keySpan.classList.add('key');
      div.appendChild(keySpan);

      const valueSpan = document.createElement('span');
      valueSpan.textContent = entry[key];
      div.appendChild(valueSpan);

      entryDiv.appendChild(div);
    }

    outputDiv.appendChild(entryDiv);
  });
}
