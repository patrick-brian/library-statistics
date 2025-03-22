// Load the 'gate-count-module' content immediately on page load
window.onload = function() {

  setActiveTab(document.getElementById('referencestats')); // Set the second tab as active by default
  while (currentTime <= 19.5) {
      let hour = Math.floor(currentTime);
          let minute = (currentTime % 1 === 0.5) ? '30' : '00';
          let period = hour < 12 ? 'AM' : 'PM';

          // Convert to 12-hour format and add leading zero for single-digit hours
          if (hour > 12) hour -= 12;
          let formattedHour = hour < 10 ? `0${hour}` : `${hour}`; // Add leading zero for single digits

          times.push(`${formattedHour}:${minute} ${period}`);
          currentTime += 0.5; // Move to next time slot
  }

  // Loop through the days of the week and time intervals
  for (let day of days) {
      for (let time of times) {
          // Generate a random headcount for each time slot (adjust the range as needed)
          const headcount = getRandomHeadcount(10, 35);

          // Append the generated data to the list
          dummyData.push({ day, time, headcount });
      }
  }
};

// Function to handle the file upload and data conversion
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = e.target.result;

            // Parse the Excel file
            const workbook = XLSX.read(data, {
                type: 'binary'
            });

            // Get the first sheet in the workbook
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert the sheet data to JSON, using the first row as headers
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            convertExcelDates(jsonData.slice(1))

            // Display the JSON data as a list
            const newData = displayList(jsonData);

            refStats = newData
                .map(obj => {
                    // Apply the filter for "Type of Inquiry:" conditions before continuing
                    if (obj["Type of Inquiry:"] !== "Gate Count" && obj["Type of Inquiry:"] !== "Roving") {
                        return Object.keys(obj)
                            .filter(key => refStatsHeaders.includes(key))  // Filter only the required properties
                            .reduce((newObj, key) => {
                                newObj[key] = obj[key];
                                return newObj;
                            }, {});
                    }
                    // If the filter condition is not met, return an empty object or nothing
                    return null; // You can change this to `return {};` if you prefer to keep an empty object
                })
                .filter(item => item !== null);

            gateCountData = newData.filter(item =>
                   item["Gate Count:"] != "" || item["Computer Lab"] != ""
                )
                .map((item, index, array) => {
                  // Get the previous item (if not the first item)
                  const previousItem = array[index - 1];

                  // Calculate "Gate Count - Daily Total" as the difference from the previous "Gate Count"
                  const gateCountDailyTotal = previousItem
                    ? previousItem["Gate Count:"] - item["Gate Count:"]  // Subtract previous "Gate Count" from current
                    : 0; // For the first item, there's no previous item, so we default to 0

                  // Calculate "Gate Count - Unique Head Count" as half of "Gate Count - Daily Total"
                  const gateCountUniqueHeadCount = gateCountDailyTotal / 2;

                  // Calculate "Computer Lab - Daily Total" as the difference from the previous "Computer Lab"
                  const computerLabDailyTotal = previousItem
                    ? previousItem["Computer Lab"] - item["Computer Lab"]
                    : 0; // For the first item, default to 0

                  // Calculate "Computer Lab - Unique Head Count" as half of "Computer Lab - Daily Total"
                  const computerLabUniqueHeadCount = computerLabDailyTotal / 2;

                  return {
                    "Submission ID": item["Submission ID"],
                    "Submitted": item["Submitted"],
                    "Gate Count:": item["Gate Count:"],
                    "Gate Count - Daily Total": gateCountDailyTotal,
                    "Gate Count - Unique Head Count": gateCountUniqueHeadCount,
                    "Computer Lab": item["Computer Lab"],
                    "Computer Lab - Daily Total": computerLabDailyTotal,
                    "Computer Lab - Unique Head Count": computerLabUniqueHeadCount,
                    "Subject(s) of Inquiry:": item["Subject(s) of Inquiry:"],
                    "Additional Information:": item["Additional Information:"]
                  };
                });

			typeOfInquiryData = newData.filter(item =>
			   item["Type of Inquiry:"] != "" && item["Type of Inquiry:"] != "Gate Count" && item["Type of Inquiry:"] != "Roving"
			)

			rovingData = newData.filter(item =>
			   item["Type of Inquiry:"] == "Roving"
			)



			loanableTechData = newData.filter(item =>
			   item["Type of Inquiry:"] == "Loanable Tech"
			)
			rawTable = createTable(refStatsHeaders, refStats, '#reference-stats-data-table')
			techTable = createTable(loanableTechHeaders, loanableTechData, '#loanable-tech-data-table');
			rovingTable = createTable(rovingHeaders, rovingData, '#roving-data-table');
            gateCountTable = createTable(gateCountHeaders, gateCountData, '#gate-count-data-table');
			inquiryTable = createTable(typeOfInquiryDataHeaders, typeOfInquiryData, '#type-of-inquiry-data-table');
        };
        reader.readAsBinaryString(file);

    }
}

// Function to display the JSON data as a list
function displayList(data, tableName) {
    // Extract the headers (first row)
    const headers = data[0];
	handleDuplicates(headers)
	const allData = [{}];
    // Loop through each row in the JSON data, starting from the second row (index 1)
    data.slice(1).forEach((item, rowIndex) => {
        // Create an object where keys are from the header and values are from the current row
        const rowData = {}

		rowData["Submission ID"] = data.length - rowIndex - 1
        // Iterate over headers and assign each value from the current row
        headers.forEach((header, index) => {
            // Assign the value to the rowData object. If the value is undefined, set it to an empty string.
            rowData[header] = item[index] !== undefined ? item[index] : '';  // Replace undefined with ""
        });
		allData[rowIndex] = rowData
    });
    headers.unshift("Submission ID");
    return allData
}

function handleDuplicates(list) {
  const textCount = {};  // To track occurrences of text
  list.forEach((item, index) => {
    // Check if text has been encountered before
    if (textCount[item]) {
      list[index] += '-';  // Append '-' if duplicate (update the array directly)
    }
    // Increment the count of the text
    textCount[item] = (textCount[item] || 0) + 1;
  });
  return list;
}

function createTable(headers, data, tableName){
    let table = $(tableName).DataTable()
    let pageInfo = table.page.info()
    let currentPage = pageInfo !== undefined ? pageInfo.page : 0;
    currentSearch = table.search();
    table.clear().destroy();
	tableColumns = headers.map(
	    item => ({
            name: item,
            title: item,
            data: item
        })
    );

    tableColumns.push({ data: null });

    // Select all elements with the class 'my-class'
    var elements = document.getElementsByClassName('table-wrapper');

    // Loop through the elements and set the background color to white
    for (var i = 0; i < elements.length; i++) {
        elements[i].style.backgroundColor = 'white';
    }

    let scrollValue;


    if (tableName === "#reference-stats-data-table") {
      scrollValue = 270;
    } else if (tableName === "#roving-data-table") {
      scrollValue = 215;
    } else if (tableName === "#gate-count-data-table") {
      scrollValue = 500;
    }

	dataTable = new DataTable(tableName, {
        data: data,
        searching: true,
        pageLength: 100,
        scrollX: false,
        scrollY: scrollValue,
        paging: true,
        columns: tableColumns,
        columnDefs: [{
            "targets": "_all",  // Disable sorting on Name and Country columns
            "orderable": false
        }, {
            targets: -1, // Target the last column
            data: null, // Do not use any data for the delete button
            render: function(data, type, row, meta) {

                let buttons = `<div style="display: flex; gap: 5px;">`;

                // Conditionally show the buttons based on tableName
                if (tableName === '#gate-count-data-table') {
                    // Show all buttons (edit, delete, add) for this table
                    buttons += `
                        <button onclick='editRow(${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-edit'></i></button>
                        <button onclick='deleteRow(false, ${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-trash'></i></button>
                        <button onclick='addRow(${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-plus'></i></button>
                    `;
                } else if (tableName === '#roving-data-table') {
                    // Only show the delete button for this table
                    buttons += `<button onclick='deleteRow(false, ${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-trash'></i></button>`;
                } else if (tableName === '#reference-stats-data-table') {
                    // Only show the edit button for this table
                    buttons += `<button onclick='editRow(${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-edit'></i></button>`;
                }

                buttons += `</div>`;
                return buttons;
              }
        }],
        footerCallback: function (row, data, start, end, display) {
            // Create the custom footer row
            var footerHTML = `
            <tr>
            <th></th>
            <th>Additional Gate Count</th>
            <th><input type="number" id="input1" inputmode="numeric" placeholder="Enter value" ></th>
            <th></th>
            <th>Additional Computer Lab</th>
            <th><input type="number" id="input2" inputmode="numeric" placeholder="Enter value"></th>
            <th><button id="calculate" onclick='calculate()'>Calculate</button></th>
            </tr>
            `;
            // Add the row to the footer
            $('#gate-count-data-table_wrapper tfoot').html(footerHTML);
        }
    })

    dataTable.search(currentSearch)

    dataTable.page(currentPage)

    dataTable.draw(false);

    return dataTable
}

// Handle Calculate button click
function calculate(){
    let tableName = "#gate-count-data-table"
	addedGateCount = parseFloat($('#input1').val()) || 0;
	addedComputerLab = parseFloat($('#input2').val()) || 0;

	let scrollPosition = $(tableName).parent().scrollTop();

	gateCountData[0]["Gate Count - Daily Total"] = addedGateCount - gateCountData[0]["Gate Count:"];
	gateCountData[0]["Gate Count - Unique Head Count"] = gateCountData[0]["Gate Count - Daily Total"]/2;
	gateCountData[0]["Computer Lab - Daily Total"] = addedComputerLab - gateCountData[0]["Computer Lab"];
	gateCountData[0]["Computer Lab - Unique Head Count"] = gateCountData[0]["Computer Lab - Daily Total"]/2;

    createTable(gateCountHeaders, gateCountData, tableName);
    document.getElementById('input1').value = addedGateCount
    document.getElementById('input2').value = addedComputerLab
    $(tableName).parent().scrollTop(scrollPosition);
    calculateTotals();
}

function convertExcelDates(list) {
  return list.map(item => {
    // Convert each item's excelDate and add it as a readable date
    item[0] = localizedDate(excelDateToJSDate(item[0]));
    item[15] !== undefined && (item[15] = localizedDate(excelDateToJSDate(item[15])));
    return item;
  });
}

function excelDateToJSDate(excelDate) {
  const epoch = new Date(1899, 11, 30); // Excel's epoch date (Dec 30, 1899)
  return new Date(epoch.getTime() + excelDate * 86399956.66); // Multiply by 86400000 to convert to milliseconds
}

function localizedDate(dateStr) {

    // Create a Date object from the input string
    const date = new Date(dateStr);

    // Get the current seconds and minutes
    const seconds = date.getSeconds();

    // Round seconds to the nearest 30
    const roundedSeconds = Math.round(seconds / 30) * 30;

    // Set the rounded seconds and reset milliseconds
    date.setSeconds(roundedSeconds);
    date.setMilliseconds(0);

    const options = {
        year: 'numeric',
        month: 'numeric',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true,  // Use 12-hour format (AM/PM)
        timeZone: 'America/Edmonton'  // Set time zone to Edmonton, Alberta
    };
    return new Intl.DateTimeFormat('en-CA', options).format(date);
}

// Function to handle the delete action
function deleteRow(cancelButton, tableName, rowIndex) {
    let deleteData;
    let headers;

    if (tableName === '#gate-count-data-table') {
        deleteData = gateCountData
        headers = gateCountHeaders
    } else if (tableName === '#roving-data-table') {
        deleteData = rovingData
        headers = rovingHeaders
    }
    let table = $(tableName).DataTable();
    let row = table.row(rowIndex).node()
    sortByIdDescending(deleteData)

    let scrollPosition = $(tableName).parent().scrollTop();
    const rows = document.querySelectorAll('tr');
    rows.forEach(row => row.classList.remove('highlighted'));

    if(cancelButton) createTable(headers, deleteData, tableName);
    // Confirmation dialog
    else if (confirm("Are you sure you want to delete Submission = " + row.cells[0].textContent)) {
        // Get the row's data or submission ID (assuming "Submission ID" is in the first column)
        // Adjust this if the "Submission ID" is in a different column
        let submissionID = row.cells[0].textContent;
        const indexToDelete = deleteData.findIndex(item => item["Submission ID"] === Number(submissionID));
        if (tableName === '#gate-count-data-table') {

            if(indexToDelete < deleteData.length - 1 && indexToDelete > 0) {
                deleteData[indexToDelete+1]["Gate Count - Daily Total"] = deleteData[indexToDelete -1]["Gate Count:"] - deleteData[indexToDelete+1]["Gate Count:"]
                deleteData[indexToDelete+1]["Gate Count - Unique Head Count"] = deleteData[indexToDelete +1]["Gate Count - Daily Total"]/2
                deleteData[indexToDelete+1]["Computer Lab - Daily Total"] = deleteData[indexToDelete - 1]["Computer Lab"] - deleteData[indexToDelete+1]["Computer Lab"]
                deleteData[indexToDelete+1]["Computer Lab - Unique Head Count"] = deleteData[indexToDelete+1]["Computer Lab - Daily Total"]/2
            }
        }
        deleteData.splice(indexToDelete, 1)
        createTable(headers, deleteData, tableName);
    } else row.classList.remove('highlighted');

    $(tableName).parent().scrollTop(scrollPosition);
    calculateTotals()
}

// Function to highlight the row and make it editable
function editRow(tableName, rowIndex) {
    removeEditOrAdd(tableName)
    let table = $(tableName).DataTable();

    let editData;
    let headers;
    let input;
    if (tableName === '#gate-count-data-table') {
        editData = gateCountData
        headers = gateCountHeaders
    } else if (tableName === '#reference-stats-data-table') {
        editData = refStats
        headers = refStatsHeaders
    }

    sortByIdDescending(editData)

    let scrollPosition = $(tableName).parent().scrollTop();
    let rowData = {}

    let row = table.row(rowIndex).node()

    row.classList.add('highlighted');
    const cells = row.querySelectorAll('td');

    // Make each cell in the row editable
    cells.forEach((cell, index) => {
        const originalText = index === 1 ? convertTo24HourFormat(cell.textContent) : cell.textContent;

        const originalWidth = cell.offsetWidth - 30; // Get the current column width
        if(tableName == "#gate-count-data-table"){
              if (index == 1 || index == 2 || index == 5 || index == 9) { // Skip the last column (buttons column)
                // Create an input field with the same width as the column
                input = index === 1
                    ? `<input type="datetime-local" value="${originalText}" class="form-control" style="width: ${originalWidth}px !important;" />`
                    : `<input type="text" value="${originalText}" class="form-control" style="width: ${originalWidth}px !important;" />`;
                cell.innerHTML = input;

              }
        } else if(tableName == '#reference-stats-data-table') {
              let options = '';

              if (index == 3) {
                options = typeOfInquiry.map(option =>
                    `<option value="${option}" ${originalText === option ? 'selected' : ''}>${option}</option>`
                ).join('');
              } else if (index == 4) {
                options = typeOfReference.map(option =>
                    `<option value="${option}" ${originalText === option ? 'selected' : ''}>${option}</option>`
                ).join('');
              } else if (index == 5) {
                options = typeOfFacilitativeInquiry.map(option =>
                    `<option value="${option}" ${originalText === option ? 'selected' : ''}>${option}</option>`
                ).join('');
              } else if (index == 6) {
                options = typeOfDigitalSupportInquiry.map(option =>
                    `<option value="${option}" ${originalText === option ? 'selected' : ''}>${option}</option>`
                ).join('');
              } else if (index == 7) {
                options = technologyType.map(option =>
                    `<option value="${option}" ${originalText === option ? 'selected' : ''}>${option}</option>`
                ).join('');
              }

              if (index == 3 || index == 4 || index == 5 || index == 6 || index == 7) {
                  input = `<select class="form-control" style="width: ${originalWidth}px !important;" >
                             ${options}
                           </select>`;
                  cell.innerHTML = input;
              }
        }
    });

    // Add a Save button to the row for saving changes
    const saveButtonHtml = `
        <div>
            <button style="margin: 3px" class='btn btn-success' onclick='saveRow(this,${JSON.stringify(rowData)}, ${JSON.stringify(tableName)})'><i class='fas fa-check'></i></button>
            <button style="margin: 3px" class='btn btn-secondary' id="cancelButton" onclick='cancelEdit(${JSON.stringify(tableName)})'>
                  <i class='fas fa-times'></i>
            </button>
        </div>
        `;

    editing = true
    row.querySelector('td:last-child').innerHTML = saveButtonHtml;
    $(tableName).parent().scrollTop(scrollPosition);
}

function addRow(tableName, rowIndex) {
    removeEditOrAdd(tableName)
    sortByIdDescending(gateCountData)
    let table = $(tableName).DataTable();

    let scrollPosition = $(tableName).parent().scrollTop();
    rowData = table.row(rowIndex).data()

    let emptyRow ={
        "Submission ID": rowData["Submission ID"]+0.1,
        "Submitted": convertTo24HourFormat("2025-01-01, 12:00:00 a.m."),
        "Gate Count:": 0,
        "Gate Count - Daily Total": "",
        "Gate Count - Unique Head Count": "",
        "Computer Lab": 0,
        "Computer Lab - Daily Total": "",
        "Computer Lab - Unique Head Count": "",
        "Subject(s) of Inquiry:": "",
        "Additional Information:": ""
    };

    // Add the new row
    var newRow = table.row.add(emptyRow).draw().node();

    // Insert the new row before the target row (rowIndex)
    var targetRow = table.row(rowIndex).node();
    var editRow = targetRow.nextSibling
    editRow.classList.add('highlighted');

    //Make each cell in the row editable
    const cells = editRow.querySelectorAll('td');

    cells.forEach((cell, index) => {
        if (index == 1 || index == 2 || index == 5 || index == 9) { // Skip the last column (buttons column)
            const originalText = cell.textContent;
            const originalWidth = cell.offsetWidth - 30; // Get the current column width

            // Create an input field with the same width as the column
            const input = index === 1
            ? `<input type="datetime-local" value="${originalText}" class="form-control" style="width: ${originalWidth}px !important;" />`
            : `<input type="text" value="${originalText}" class="form-control" style="width: ${originalWidth}px !important;" />`;
            cell.innerHTML = input;
        }
    });

    // Add a Save button to the row for saving changes
    const saveButtonHtml = `
        <button class='btn btn-success' onclick='saveRow(this,${JSON.stringify(emptyRow)}, ${JSON.stringify(tableName)})'><i class='fas fa-check'></i></button>
        <button class='btn btn-secondary' onclick='deleteRow(true, ${JSON.stringify(tableName)})'>
        <i class='fas fa-times'></i>
        </button>
    `;
    adding = true
    editRow.querySelector('td:last-child').innerHTML = saveButtonHtml;
    targetRow.parentNode.insertBefore(newRow, editRow);  // Insert before the target row
    $(tableName).parent().scrollTop(scrollPosition);
}

// Function to remove highlight and save edited values (this can be triggered when you want to save the edits)
function saveRow(button, rowData, tableName) {
    let saveData;
    let headers;

    if (tableName === '#gate-count-data-table') {
        saveData = gateCountData
        headers = gateCountHeaders
    } else if (tableName === '#reference-stats-data-table') {
        saveData = refStats
        headers = refStatsHeaders
    }

    sortByIdDescending(saveData)
    let scrollPosition = $(tableName).parent().scrollTop();

    const row = button.closest('tr');

    // Remove the highlight class
    row.classList.remove('highlighted');

    // Get all cells in the row
    const cells = row.querySelectorAll('td');

    let submissionId = Number(cells[0].innerText);

    // Loop through the cells and extract the values
    cells.forEach((cell, index) => {
        if (index !== cells.length - 1) { // Skip the last column (buttons column)
            let input;

            if (tableName === '#gate-count-data-table')
                input = index === 1 ? cell.querySelector('.form-control') : cell.querySelector('input'); // Get the input element
            else if (tableName === '#reference-stats-data-table') input = cell.querySelector('select')

            if (input) {
                const columnName = headers[index]; // Get column name (assuming tableColumns array contains column names)
                // Map the input values to the corresponding properties
                switch (columnName) {
                    case "Type of Inquiry:":
                        rowData["Type of Inquiry:"] = input.value;
                        break;
                    case "Type of Reference:":
                        rowData["Type of Reference:"] = input.value;
                        break;
                    case "Type of Facilitative Inquiry:":
                        rowData["Type of Facilitative Inquiry:"] = input.value;
                        break;
                    case "Type of  Digital Support Inquiry:":
                        rowData["Type of  Digital Support Inquiry:"] = input.value;
                        break;
                    case "Technology Item Type:":
                        rowData["Technology Item Type:"] = input.value;
                        break;
                    case "Submitted":
                        rowData["Submitted"] = convertTo12HourFormat(input.value);
                        break;
                    case "Gate Count:":
                        rowData["Gate Count:"] = Number(input.value);
                        break;
                    case "Computer Lab":
                        rowData["Computer Lab"] = Number(input.value);
                        break;
                    case "Additional Information:":
                        rowData["Additional Information:"] = input.value;
                        break;
                    default:
                    // If the column is not recognized, just skip it or handle as needed
                    break;
                }

                // Replace input field with the updated value in the cell
                cell.innerHTML = input.value;
            }
        }
    });

    // Find the corresponding row in gateCountData and update it
    const indexToUpdate = saveData.findIndex(item => item["Submission ID"] === Number(submissionId));
    if (indexToUpdate !== -1) {
        // Update the found item with the new rowData values
        saveData[indexToUpdate] = { ...saveData[indexToUpdate], ...rowData };
        if (tableName === '#gate-count-data-table') {

            saveData[indexToUpdate]["Gate Count - Daily Total"] = saveData[indexToUpdate - 1]["Gate Count:"] - saveData[indexToUpdate]["Gate Count:"]
            saveData[indexToUpdate]["Gate Count - Unique Head Count"] = saveData[indexToUpdate]["Gate Count - Daily Total"]/2
            saveData[indexToUpdate]["Computer Lab - Daily Total"] = saveData[indexToUpdate - 1]["Computer Lab"] - saveData[indexToUpdate]["Computer Lab"]
            saveData[indexToUpdate]["Computer Lab - Unique Head Count"] = saveData[indexToUpdate]["Computer Lab - Daily Total"]/2
        }
    } else {
        // Find the position to insert the new element
        for (let i = 0; i < gateCountData.length - 1; i++) {
            if (saveData[i]["Submission ID"] > submissionId && saveData[i + 1]["Submission ID"] < submissionId) {
                rowData["Gate Count - Daily Total"] = saveData[i]["Gate Count:"] - rowData["Gate Count:"]
                rowData["Gate Count - Unique Head Count"] = rowData["Gate Count - Daily Total"]/2
                rowData["Computer Lab - Daily Total"] = saveData[i]["Computer Lab"] - rowData["Computer Lab"]
                rowData["Computer Lab - Unique Head Count"] = rowData["Computer Lab - Daily Total"]/2
                saveData[i+1]["Gate Count - Daily Total"] = rowData["Gate Count:"] - saveData[i+1]["Gate Count:"]
                saveData[i+1]["Gate Count - Unique Head Count"] = saveData[i+1]["Gate Count - Daily Total"]/2
                saveData[i+1]["Computer Lab - Daily Total"] = rowData["Computer Lab"] - saveData[i+1]["Computer Lab"]
                saveData[i+1]["Computer Lab - Unique Head Count"] = saveData[i+1]["Computer Lab - Daily Total"]/2
                saveData.splice(i + 1, 0, rowData);  // Insert the new element between i and i+1
                break;
            }
        }
    }
    createTable(headers, saveData, tableName);
    $(tableName).parent().scrollTop(scrollPosition);

    if (tableName === '#gate-count-data-table') calculateTotals();
}

// Function to cancel editing and revert the changes
function cancelEdit(tableName) {
    let cancelData;
    let headers;

    if (tableName === '#gate-count-data-table') {
        cancelData = gateCountData
        headers = gateCountHeaders
    } else if (tableName === '#reference-stats-data-table') {
        cancelData = refStats
        headers = refStatsHeaders
    }

    sortByIdDescending(cancelData)
    let scrollPosition = $(tableName).parent().scrollTop();

    const rows = document.querySelectorAll('tr');
    rows.forEach(row => row.classList.remove('highlighted'));

    createTable(headers, cancelData, tableName);
    $(tableName).parent().scrollTop(scrollPosition);
}

function convertTo24HourFormat(dateString) {

    // Example: "2025-01-02, 08:20:00 a.m." or "2025-01-02, 08:20:00 p.m."

    // First, remove the comma and 'a.m.' / 'p.m.' part
    const parts = dateString.split(",");  // ["2025-01-02", "08:20:00 a.m."]

    let datePart = parts[0].trim(); // "2025-01-02"
    let timePart = parts[1].trim(); // "08:20:00 a.m."

    // Split the time and period (AM/PM)
    const timeAndPeriod = timePart.split(" ");
    let time = timeAndPeriod[0];  // "08:20:00"
    const period = timeAndPeriod[1].toLowerCase();  // "a.m." or "p.m."

    // Convert time to 24-hour format
    let [hours, minutes, seconds] = time.split(":").map(num => parseInt(num));

    if (period === "p.m." && hours !== 12) hours += 12; // Add 12 to hours for PM, except for 12 PM
    else if (period === "a.m." && hours === 12) hours = 0; // Convert 12 AM to 00:00


    // Format hours and minutes to 2 digits
    const formattedTime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
    // Combine the date and formatted time
    return `${datePart}T${formattedTime}`;
}

function convertTo12HourFormat(datetime) {
    // Split the datetime into date and time parts
    const [date, time] = datetime.split('T');

    // Split time into hours and minutes
    let [hours, minutes] = time.split(":").map(num => parseInt(num));

    // Determine if it's AM or PM
    const period = hours >= 12 ? 'p.m.' : 'a.m.';

    // Convert hours to 12-hour format
    if (hours > 12) hours -= 12;
    else if (hours === 0) hours = 12; // Midnight (00:00) is 12 a.m.

    // Format the time in 12-hour format with leading zeros if needed
    const formattedTime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:00 ${period}`;

    // Combine date and formatted time
    return `${date}, ${formattedTime}`;
  }

function calculateTotals() {
    totalGateCount = 0;
    totalComputerLab = 0;

    if(document.getElementById("gateCountTab").classList.contains("active")) {
        // Loop through each item and accumulate the cost and profit
        gateCountData.forEach(item => {
            totalGateCount = Number(totalGateCount) + Number(item["Gate Count - Unique Head Count"]);   // Sum the cost
            totalComputerLab = Number(totalComputerLab) + Number(item["Computer Lab - Unique Head Count"]); // Sum the profit
        });

        document.getElementById('gateNumber').innerHTML = rollNumber("gateNumber", totalGateCount);
        document.getElementById('computerNumber').innerHTML = rollNumber("computerNumber", totalComputerLab);

        totalDays = document.getElementById('total-days').value;
        lastYear = document.getElementById('last-year').value;
        totalGateCountAverage = (totalGateCount/totalDays).toFixed(2)
        totalLabAverage = (totalComputerLab/totalDays).toFixed(2)

        document.getElementById('gate-count-average').innerHTML =
            `<span class="number input"> ${(!isFinite(totalGateCountAverage) || isNaN(totalGateCountAverage)) ? "0" : totalGateCountAverage}</span>`;
        document.getElementById('computer-lab-average').innerHTML =
            `<span class="number input">  ${(!isFinite(totalLabAverage) || isNaN(totalLabAverage)) ? "0" : totalLabAverage}</span>`;

        let changePercentage = (((totalGateCount/lastYear)-1)*100).toFixed(2)
        changeText = `${(!isFinite(changePercentage) || isNaN(changePercentage)) ? "0" : Math.abs(changePercentage)}%` +
                     (lastYear > 0 ? ` ${changePercentage < 0 ? changeColor('decrease') : changeColor('increase')}` : '');
        document.getElementById('overallCount').innerHTML = changeText;
    }
}

function changeColor(status) {
    let card = document.getElementById("changeCard")
    if (status === 'decrease') { // warm green (#4CAF50)
        card.style.backgroundColor = '#F44336'; // warm red (#F44336)
    } else {
        card.style.backgroundColor = '#4CAF50'; // warm green (#4CAF50)
    }
    return status
}
function rollNumber(elementId, targetNumber) {
    //const cards = document.querySelectorAll('.card');
    //cards.forEach(card => {
        const numberElement = document.getElementById(elementId);

        // Add the slot rolling class to trigger the animation
        void numberElement.offsetWidth; // Trigger reflow to restart animation

        // Function to simulate the slot machine effect
        let counter = 0;
        const rollInterval = setInterval(() => {
            // Randomize a number between 1 and 100
            numberElement.textContent = targetNumber;
            counter++;

            // After showing numbers for a set number of intervals, stop the slot machine
            if (counter >= 10) { // Number of "spins" before stopping
                clearInterval(rollInterval);

                // Show the final random number after stopping
                const randomNumber = targetNumber;
                numberElement.textContent = randomNumber;
            }
        }, 50); // Interval between number changes (100ms for fast "spinning")
    //});
    return targetNumber
}
function exportReport() {
    sortByIdAscending(gateCountData)
    sortByIdAscending(rovingData)
    sortByIdAscending(refStats)

    let rovingColumns = ["Roving Time", "Study Rooms", "Group Tables", "Study Carrels", "Computer Lab-", "Additional Information:", "Subject(s) of Inquiry:"]
    let exportRoving = rovingData.map(item => {
            let filteredItem = {};
            rovingColumns.forEach(property => {
                if (item[property] !== undefined) {
                    filteredItem[property] = item[property];
                }
            });
            return filteredItem;
        });

    let exportRefStats = refStats.map(item => {
            const { 'Submission ID': submissionId, ...rest } = item; // Destructure to exclude 'Submitted'
            return rest; // Return the new object without 'Submitted'
        });

    // Create a new workbook
    const workbook = new ExcelJS.Workbook();

    // Create the first worksheet 'KC Library Ref Stats'
    const wsRefStats = workbook.addWorksheet('KC Library Ref Stats');
    wsRefStats.columns = [
		{ header: 'Submitted', key: 'Submitted'},
        { header: 'Method of Inquiry:', key: 'Method of Inquiry:'},
        { header: 'Type of Inquiry:', key: 'Type of Inquiry:'},
        { header: 'Type of Reference:', key: 'Type of Reference:'},
		{ header: 'Type of Facilitative Inquiry:', key: 'Type of Facilitative Inquiry:'},
		{ header: 'Type of Digital Support Inquiry:', key: 'Type of Digital Support Inquiry:'},
		{ header: 'Technology Item Type:', key: 'Technology Item Type:'},
		{ header: 'Software/Application Type:', key: 'Software/Application Type:'},
		{ header: "Student's Program", key: "Student's Program"},
		{ header: 'Year of Program', key: 'Year of Program'},
		{ header: 'Was Tech available at the time of request?', key: 'Was Tech available at the time of request?'},
		{ header: 'Subject(s) of Inquiry:', key: 'Subject(s) of Inquiry:'},
		{ header: 'Additional Information:', key: 'Additional Information:'},
    ];

	const refStatsHeaders = ['Submission ID', 'Submitted', 'Method of Inquiry:', 'Type of Inquiry:', 'Type of Reference:', 'Type of Facilitative Inquiry:',
                'Type of  Digital Support Inquiry:', 'Technology Item Type:', 'Software/Application Type:', "Student's Program", 'Year of Program',
                'Was Tech available at the time of request?', 'Subject(s) of Inquiry:', 'Additional Information:']
    exportRefStats.forEach(item => {
        wsRefStats.addRow(item);
    });

    // Create the second worksheet 'Gate Count Summary'
    const wsGateCountSummary = workbook.addWorksheet('Gate Count Summary');
    wsGateCountSummary.columns = [
        { header: 'Date', key: 'Submitted'},
        { header: 'Gate Count', key: 'Gate Count:'},
        { header: 'Daily Total', key: 'Gate Count - Daily Total'},
        { header: 'Unique Head Count (Daily Total/2)', key: 'Gate Count - Unique Head Count'},
        { header: 'Computer Lab', key: 'Computer Lab'},
        { header: 'Daily Total', key: 'Computer Lab - Daily Total'},
        { header: 'Unique Head Count (Daily Total/2)', key: 'Computer Lab - Unique Head Count'},
        { header: 'Subject(s) of Inquiry', key: 'Subject(s) of Inquiry:'},
        { header: 'Additional Information', key: 'Additional Information:'}
    ];
    gateCountData.forEach(item => {
        wsGateCountSummary.addRow(item);
    });

    // Create the third worksheet 'Roving Count'
    const wsRoving = workbook.addWorksheet('Roving Count');
    wsRoving.columns = [
        { header: 'Roving Time', key: 'Roving Time'},
        { header: 'Study Rooms', key: 'Study Rooms'},
        { header: 'Group Tables', key: 'Group Tables'},
        { header: 'Study Carrels', key: 'Study Carrels'},
        { header: 'Computer Lab', key: 'Computer Lab-'},
        { header: 'Additional Information:', key: 'Additional Information:'},
        { header: 'Subject(s) of Inquiry:', key: 'Subject(s) of Inquiry:'}
    ];

    exportRoving.forEach(item => {
        wsRoving.addRow(item);
    });

    const lastRow = wsGateCountSummary.lastRow

    wsGateCountSummary.insertRow(lastRow.number + 2, ['Date', addedGateCount, '', '', addedComputerLab]);
    wsGateCountSummary.insertRow(lastRow.number + 5, ['','', 'Total', totalGateCount, '', '', 'Total Lab', totalComputerLab]);
    wsGateCountSummary.insertRow(lastRow.number + 6, ['','', 'Average per day:', parseFloat(totalGateCountAverage), '', '', 'Average per day', parseFloat(totalLabAverage)]);
    wsGateCountSummary.insertRow(lastRow.number + 9, ['', '', 'Year Over Year Comparison']);
    wsGateCountSummary.insertRow(lastRow.number + 10, ['','', 'Last year', lastYear]);
    wsGateCountSummary.insertRow(lastRow.number + 11, ['','', 'Increase / Decrease:', changeText]);

    // Apply styles to the first row (header) for all sheets
    [wsRefStats, wsGateCountSummary, wsRoving].forEach(sheet => {
        const headerRow = sheet.getRow(1);
        let maxLength = 0;
        headerRow.eachCell((cell, colNumber) => {
            const value = cell.value ? cell.value.toString() : '';
            maxLength = Math.max(maxLength, value.length);
            cell.font = { bold: true }; // Bold text
            cell.alignment = { horizontal: 'center', vertical: 'middle' }; // Center alignment
            cell.alignment.wrapText = true; // Enable text wrapping

            const column = sheet.getColumn(colNumber);
                if(maxLength > 30)
                    column.width = maxLength - 18;
                else if(maxLength > 15)
                    column.width = maxLength - 5; // Add padding
                else
                    column.width = maxLength + 10
        });

        // Enable filter for the header row
        sheet.autoFilter = {
            from: { row: 1, column: 1 }, // Enable autofilter for the whole header row (row 1)
            to: { row: 1, column: sheet.columnCount } // End at the last column
        };

        // Freeze the top row (header)
        sheet.views = [
            {
                state: 'frozen',
                ySplit: 1, // Freeze row 1 (index 0 is the first row)
                topLeftCell: 'A2', // Ensure that the first row is frozen and starting from A2
            }
        ];
    });

    // Apply auto column width for each worksheet
    [wsRefStats, wsGateCountSummary, wsRoving].forEach((worksheet) => {
        worksheet.columns.forEach(column => {
            if (column.width > 10) {
                column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
                    if (rowNumber > 1) {  // Exclude the header row from wrap text
                        cell.alignment = cell.alignment || {}; // Ensure alignment object exists
                        cell.alignment.wrapText = true;  // Enable wrap text for non-header rows
                    }
                });
            }

        });
    });

    // Write the workbook to a buffer and download the file
    workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'raw_data.xlsx';
        link.click();
    });

}

function sortByIdAscending(data) {
    return data.sort((a, b) => a["Submission ID"] - b["Submission ID"]);
}

function sortByIdDescending(data) {
    return data.sort((a, b) => b["Submission ID"] - a["Submission ID"]);
}

// Function to trigger the file input dialog
function triggerFileInput() {
    document.getElementById("excel-upload").click();
}

// Function to toggle the active class when a tab is clicked
function setActiveTab(selectedTab) {
    if(!selectedTab.classList.contains("active")) {
        let contentArea = document.getElementById("content-area");
        let fileToLoad = "";
        // Define the content for each module
        let content = "";

        switch (selectedTab.innerText) {
            case 'KC Library Ref Stats':
                fileToLoad = "modules/reference-stats-module.html";
                headers = refStatsHeaders
                data = refStats
                tableName = "#reference-stats-data-table"
                break;
            case 'Gate Count':
                fileToLoad = "modules/gate-count-module.html";
                headers = gateCountHeaders
                data = gateCountData
                tableName = "#gate-count-data-table"
                break;
            case 'Roving Count':
                fileToLoad = "modules/roving-count-module.html";
                headers = rovingHeaders
                data = rovingData
                tableName = "#roving-data-table"
                break;
            case 'module4':
                content = "<h2>Module 4 Content</h2><p>This is the content for Module 4.</p>";
                break;
            case 'module5':
                content = "<h2>Module 5 Content</h2><p>This is the content for Module 5.</p>";
                break;
            default:
                break;
        }

        // Update the content area with the selected module's content
        if (fileToLoad) {
            // Use Fetch API to load the external HTML file
            fetch(fileToLoad)
            .then(response => {
                if (response.ok) {
                    return response.text();
                }
                throw new Error('Failed to load the content.');
            })
            .then(html => {
                contentArea.innerHTML = html; // Insert the HTML content into the content area
            })
            .catch(error => {
                contentArea.innerHTML = `<p>Error loading content: ${error.message}</p>`;
            });
        } else {
            // If there's no external file to load, display the predefined content
            contentArea.innerHTML = content;
        }

        setTimeout(function () {
            createTable(headers, data, tableName);
            if(selectedTab.innerText == "Gate Count") {
                document.getElementById('total-days').value = totalDays;
                document.getElementById('last-year').value = lastYear;
                document.getElementById('input1').value = addedGateCount
                document.getElementById('input2').value = addedComputerLab
                calculateTotals()
            } else if (selectedTab.innerText == "Roving Count") {
                groupRovingData();
                generateTable("study-room-chart", computerLabAvgHeadCounts);
                generateTable("group-table-chart", groupTablesAvgHeadCounts);
                generateTable("study-carrel-chart", studyCarrelsAvgHeadCounts);
                generateTable("computer-lab-chart", studyRoomAvgHeadCounts);

            }
        }, 50)
        // Remove active class from all tabs
        const tabs = document.querySelectorAll('.side-tab ul li');
        tabs.forEach(tab => {
            tab.classList.remove('active');
        });

        // Add active class to the clicked tab
        selectedTab.classList.add('active');

    }
}

// Modify the onClick handlers in the HTML to trigger the active class change
document.querySelectorAll('.side-tab ul li').forEach(tab => {
    tab.addEventListener('click', function() {
        if (tab.innerText == "Upload Excel File") triggerFileInput()
        else if (tab.innerText == "Export Report") exportReport()
        else setActiveTab(tab);
    });
});

function removeEditOrAdd(tableName) {
   if (adding) {
        deleteRow(true, tableName)
        adding = false
   } else if (editing) {
        cancelEdit(tableName)
        editing = false
   }
}
const mainData = [];

function groupRovingData() {
   rovingData.forEach(item => {
       item["Roving Time"] = item["Roving Time"] && item["Roving Time"].trim() !== ""
           ? roundToNearestHalfHour(item["Roving Time"])
           : roundToNearestHalfHour(item["Submitted"]) ; // Assign an empty string if Submitted is undefined or an empty string
   });
   separateHeadCounts(calculateHeadCountByDay())
}

const computerLabHeadCount = [];
const groupTablesHeadCount = [];
const studyCarrelsHeadCount = [];
const studyRoomHeadCount = [];
const computerLabAvgHeadCounts = [];
const groupTablesAvgHeadCounts = [];
const studyCarrelsAvgHeadCounts = [];
const studyRoomAvgHeadCounts = [];

function getMin(data) {
    // Ensure data is not empty
    if (!data || data.length === 0) {
        return null; // or return some default value like `Infinity` if needed
    }

    // Use Math.min with the spread operator to get the minimum value from the array of averages
    return Math.min(...data.map(item => parseFloat(item.average)));
}

function getMax(data) {
    // Ensure data is not empty
    if (!data || data.length === 0) {
        return null; // or return some default value like `-Infinity` if needed
    }

    // Use Math.max with the spread operator to get the maximum value from the array of averages
    return Math.max(...data.map(item => parseFloat(item.average)));
}

function calculateHeadCountByDay() {
     // Days of the week array
     const daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

     // Initialize an object to store headcount data grouped by day and time
     const headCountByDayAndTime = {};

     // Iterate through rovingData and calculate headcount based on day of the week and rounded time
     rovingData.forEach(item => {
         // Convert Roving Time to 24-hour format for proper parsing
         const rovingTimeFormatted = convertTo24HourFormat(item["Roving Time"]);
         const rovingDate = new Date(rovingTimeFormatted.replace(",", "")); // Remove comma for proper date parsing

         // Get the day of the week (0-6, 0 = Sunday, 1 = Monday, ...)
         const dayOfWeek = rovingDate.getDay();
         const timeRounded = rovingDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

         // Create a key for the combination of day and rounded time
         const dayTimeKey = `${daysOfWeek[dayOfWeek]} ${timeRounded}`;


         // Get the headcounts for each category, assigning 0 if empty, null or empty string
         const getValidHeadCount = (value) => {
             return value === null || value === "" || value === undefined ? 0 : value;
         };

         // Get the headcounts for each category
         const computerLabCount = getValidHeadCount(item["Computer Lab-"]);
         const groupTablesCount = getValidHeadCount(item["Group Tables"]);
         const studyCarrelsCount = getValidHeadCount(item["Study Carrels"]);
         const studyRoomsCount = getValidHeadCount(item["Study Rooms"]);

         // Add the headcount data to the corresponding day and time slot
         if (!headCountByDayAndTime[dayTimeKey]) {
             headCountByDayAndTime[dayTimeKey] = {
                 computerLabHeadCounts: [],
                 groupTablesHeadCounts: [],
                 studyCarrelsHeadCounts: [],
                 studyRoomHeadCounts: []
             };
         }

         headCountByDayAndTime[dayTimeKey].computerLabHeadCounts.push(computerLabCount);
         headCountByDayAndTime[dayTimeKey].groupTablesHeadCounts.push(groupTablesCount);
         headCountByDayAndTime[dayTimeKey].studyCarrelsHeadCounts.push(studyCarrelsCount);
         headCountByDayAndTime[dayTimeKey].studyRoomHeadCounts.push(studyRoomsCount);
     });

     // Calculate the average headcounts for each day-time combination
     const calculateAverage = (data) => {
         return data.map(item => {
             const avgComputerLab = (item.computerLabHeadCounts.reduce((acc, val) => acc + val, 0) / item.computerLabHeadCounts.length).toFixed(2);
             const avgGroupTables = (item.groupTablesHeadCounts.reduce((acc, val) => acc + val, 0) / item.groupTablesHeadCounts.length).toFixed(2);
             const avgStudyCarrels = (item.studyCarrelsHeadCounts.reduce((acc, val) => acc + val, 0) / item.studyCarrelsHeadCounts.length).toFixed(2);
             const avgStudyRooms = (item.studyRoomHeadCounts.reduce((acc, val) => acc + val, 0) / item.studyRoomHeadCounts.length).toFixed(2);

             // Return the separate fields for day and time along with the averages
             const [day, time, ampm] = item.dayTime.split(" ");
             return {
                 day: day, // Day of the week (e.g., Monday)
                 time: time + " " + ampm, // Time in "HH:mm" format (e.g., 07:30)
                 averageComputerLab: avgComputerLab,
                 averageGroupTables: avgGroupTables,
                 averageStudyCarrels: avgStudyCarrels,
                 averageStudyRooms: avgStudyRooms
             };
         });
     };

     const result = Object.keys(headCountByDayAndTime).map(dayTimeKey => ({
         dayTime: dayTimeKey,
         ...headCountByDayAndTime[dayTimeKey]
     }));


     return calculateAverage(result);
}

function addHeadCount(headCountArray, dayOfWeek, headCount) {
    // Check if the day already exists in the array
    if (!headCountArray[dayOfWeek]) {
        // Initialize the day if it doesn't exist
        headCountArray[dayOfWeek] = { day: dayOfWeek, headCounts: [] };
    }
    // Add the headcount for that specific day
    headCountArray[dayOfWeek].headCounts.push(headCount);
}

function fixHourFormat(timeString) {
    const [date, time, period] = timeString.split(" ");
    let [hours, minutes, seconds] = time.split(":");

    if (period === "p.m." && hours !== "12") {
        hours = (parseInt(hours) + 12).toString(); // Add 12 to the hour for PM times (except 12 PM)
    } else if (period === "a.m." && hours === "12") {
        hours = "00"; // Convert 12 AM to 00 hours
    }

    return `${date} ${hours}:${minutes}:${seconds}`;
}

function roundToNearestHalfHour(timeString) {
    // Split the date and time part
    let [datePart, timePart] = timeString.split(", ");

    // Parse the time to 24-hour format
    let [time, period] = timePart.split(" ");
    let [hour, minute, second] = time.split(":").map(Number);

    // Convert to 24-hour format
    if (period === "p.m." && hour !== 12) {
        hour += 12; // Convert PM times to 24-hour format
    }
    if (period === "a.m." && hour === 12) {
        hour = 0; // Midnight case (12:00 AM)
    }

    // Create a new Date object from the parsed values
    let date = new Date(`${datePart}T${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}:${second.toString().padStart(2, '0')}`);

    // Get the minutes of the current time
    let minutes = date.getMinutes();

    // Round the minutes to the nearest 30th minute (either 00 or 30)
    if (minutes < 15) {
        date.setMinutes(0, 0, 0); // Round down to the start of the hour
    } else if (minutes < 45) {
        date.setMinutes(30, 0, 0); // Round to the 30th minute
    } else {
        date.setMinutes(0, 0, 0); // Round up to the next hour
        date.setHours(date.getHours() + 1); // Increment the hour
    }

    // Format the date back into the desired string
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    let formattedHour = date.getHours();
    const formattedMinute = date.getMinutes().toString().padStart(2, '0');
    const formattedSecond = date.getSeconds().toString().padStart(2, '0');
    const periodFinal = formattedHour >= 12 ? 'p.m.' : 'a.m.';

    // Convert to 12-hour format
    formattedHour = formattedHour > 12 ? formattedHour - 12 : (formattedHour === 0 ? 12 : formattedHour);

    // Return the formatted string
    return `${year}-${month}-${day}, ${formattedHour}:${formattedMinute}:${formattedSecond} ${periodFinal}`;
}

function separateHeadCounts(data) {
    // Initialize separate variables for each headcount category


    // Iterate over each record in the data
    data.forEach(item => {
        // Prepare the object with day, time, and average for each headcount category
        const dayTimeData = {
            day: item.day,
            time: item.time,
            average: item.averageComputerLab // Starting with averageComputerLab, will handle others similarly
        };

        // Push each category into its respective array
        computerLabAvgHeadCounts.push({
            ...dayTimeData,
            average: item.averageComputerLab // For computer lab headcount
        });

        groupTablesAvgHeadCounts.push({
            ...dayTimeData,
            average: item.averageGroupTables // For group tables headcount
        });

        studyCarrelsAvgHeadCounts.push({
            ...dayTimeData,
            average: item.averageStudyCarrels // For study carrels headcount
        });

        studyRoomAvgHeadCounts.push({
            ...dayTimeData,
            average: item.averageStudyRooms // For study rooms headcount
        });
    });

//    // Return the 4 variables containing the separated data
//    return {
//        computerLabAvgHeadCounts,
//        groupTablesAvgHeadCounts,
//        studyCarrelsAvgHeadCounts,
//        studyRoomAvgHeadCounts
//    };
}

// Function to generate the table
function generateTable(tableName, tableData) {
    console.log(tableData.filter(item => item.day === "Monday"))
    let tooltip;
    const table = document.getElementById(tableName);
    const headerRow = document.createElement('tr');

    // Create the header row with days
    const blankHeader = document.createElement('th'); // Empty top-left corner cell
    headerRow.appendChild(blankHeader);
    days.forEach(day => {
        const th = document.createElement('th');
        th.innerText = day;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);



    // Function to calculate the red intensity based on headcount
    function getRedShade(headcount) {
        const minHeadcount = getMin(tableData); // Minimum headcount
        const maxHeadcount = getMax(tableData); // Maximum headcount
        // Normalize the headcount to a value between 0 and 1
        const normalized = Math.min(Math.max((headcount - maxHeadcount) / (minHeadcount - maxHeadcount), 0), 1);

        // Calculate the red intensity (255 being the darkest red)
        const redIntensity = Math.floor(130 + (125 * normalized));

        // Return the background color in RGB format
        return `rgb(${redIntensity}, 0, 0)`; // Red, no green, no blue
    }

    // Create rows for each time slot
    times.forEach(time => {
        const row = document.createElement('tr');
        const timeCell = document.createElement('td');
        timeCell.innerText = time;
        row.appendChild(timeCell);

        // Create a cell for each day and insert the corresponding headcount
        days.forEach(day => {
            const td = document.createElement('td');
            const cellData = tableData.find(entry => entry.day === day && entry.time === time);
            if (cellData && cellData.average > 0) {
                td.innerText = cellData.average;
                td.style.backgroundColor = getRedShade(cellData.average); // Apply background color

                // Set the tooltip content to show day, time, and headcount
                td.addEventListener('mouseenter', () => {
                     // Create a tooltip div (will be used for showing tooltips)
                    tooltip = document.createElement('div');
                    tooltip.classList.add('tooltip');
                    document.body.appendChild(tooltip);
                    tooltip.innerText = `Day: ${day}\nTime: ${time}\nAverage Headcount: ${cellData.average}`;
                    tooltip.style.left = `${td.getBoundingClientRect().left}px`;  // Position tooltip horizontally
                    tooltip.style.top = `${td.getBoundingClientRect().bottom + 5}px`;  // Position tooltip below the cell (5px space)
                    tooltip.classList.add('visible');  // Show tooltip
                });

                td.addEventListener('mouseleave', () => {
                    tooltip.remove();  // Hide tooltip
                });
            } else {
                td.innerText = ''; // If no data, leave empty
                td.style.backgroundColor = ''; // No background color if no data
            }

            row.appendChild(td);
        });

        table.appendChild(row);
    });
}

let refStatsHeaders = ['Submission ID', 'Submitted', 'Method of Inquiry:', 'Type of Inquiry:', 'Type of Reference:', 'Type of Facilitative Inquiry:',
                'Type of  Digital Support Inquiry:', 'Technology Item Type:', 'Software/Application Type:', "Student's Program", 'Year of Program',
                'Was Tech available at the time of request?', 'Subject(s) of Inquiry:', 'Additional Information:']

let gateCountHeaders = [
    "Submission ID",
    "Submitted", "Gate Count:", "Gate Count - Daily Total",  "Gate Count - Unique Head Count",
    "Computer Lab",  "Computer Lab - Daily Total",  "Computer Lab - Unique Head Count",
     "Subject(s) of Inquiry:", "Additional Information:"]

let typeOfInquiryDataHeaders = [
    "Submission ID", "Submitted", "Method of Inquiry:", "Type of Inquiry:", "Type of Facilitative Inquiry:", "Type of  Digital Support Inquiry:", "Type of Reference:", "Additional Information:", "Subject(s) of Inquiry:"
]
let rovingHeaders = [
    "Submission ID", "Submitted", "Roving Time", "Method of Inquiry:", "Study Rooms", "Group Tables", "Study Carrels", "Computer Lab-", "Additional Information:", "Subject(s) of Inquiry:"
]

let loanableTechHeaders = [
    "Submission ID", "Submitted", "Method of Inquiry:", "Technology Item Type:", "Was Tech available at the time of request?", "Student's Program", "Additional Information:", "Subject(s) of Inquiry:"
]

let typeOfInquiry = [
    "Basic Reference", "Complex Reference", "Facilitative", "Loanable Tech", "Digital Support", "Gate Count", "Roving"
]

let typeOfReference = [
    "Citation Help", "Copyright", "Database Help", "Find a Resource (print or online)"
]

let typeOfFacilitativeInquiry = [
    "Accessible Format Request", "Community User", "General Library Information (e.g. hours, borrowing period, etc.)", "Interlibrary Loans/Requests/Holds", "Library Account (e.g. pin, renewals, fines, etc.)", "Referral/Directional (External - Bookstore, Registrar, Academic Success Centre, etc.)", "Referral/Directional (In Library - BAL, Copyright, EdTech, Instruction, etc.)", "Reserve Request", "Scan-On-Demand", "Study Room", "Supplies (e.g. stapler, pen, hole punch, etc.)"
]

let typeOfDigitalSupportInquiry = [
    "Document Assistance (e.g. Microsoft Word, Excel, PDF, Google Docs, etc.)", "Internet/Wifi Connectivity", "Keyano Account Access (e.g. Webmail, Moodle, or Self-Service)", "LMS (Moodle. McGraw, MyLAB IT)", "Online Navigation (e.g. opening a browser or searching in Google)", "Print/Scan/Copy Assistance or Troubleshooting", "Software (M365, Respondus, Safe Exam, etc.)"
]

let technologyType = [
    "Calculator", "Camera", "Charger, Adapter, etc.", "Chromebooks", "Chromecast", "DVD Player", "Headphones", "Keyboard", "Laptops", "MFA Token", "Power Bank", "Projector", "SAD Light", "WebCam", "WIreless Mouse"
]

let adding = false;
let editing = false;
let tableHeaders;
let tableData;
let gateCountData = [];
let typeOfInquiryData;
let rovingData;
let loanableTechData;
let savedData = [];
let refStats;
let addedGateCount = 0;
let addedComputerLab = 0;
let totalGateCount = 0;
let totalComputerLab = 0;
let totalGateCountAverage = 0;
let totalLabAverage = 0;
let lastYear = 0;
let totalDays = 0
let changeText = '';
let techTable;
let rovingTable;
let gateCountTable;
let inquiryTable;
let currentSearch = "";
let dataTable;

let days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

// Time slots (7:30 AM to 7:30 PM, in 30-minute intervals)
let times = [];
let currentTime = 7.5; // 7:30 AM in 24-hour format


let dummyData = [];



function getRandomHeadcount(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}