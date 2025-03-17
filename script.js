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

let tableHeaders;
let tableData;
let gateCountData;
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
let changeText = '';
let techTable;
let rovingTable;
let gateCountTable;
let inquiryTable;

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

            //console.log("json ", jsonData.slice(1))
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
            console.log(rovingData)
			rawTable = createTable(refStatsHeaders, refStats, '#raw-data-data-table')
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
    //console.log(headers)



	handleDuplicates(headers)
	const allData = [{}];

    // Loop through each row in the JSON data, starting from the second row (index 1)
    data.slice(1).forEach((item, rowIndex) => {

        // Create an object where keys are from the header and values are from the current row
        const rowData = {}
        //rowData["Submission ID"] = rowIndex + 1
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

function createTable(headers, data, tableName, dataTable, editable){
	tableColumns = headers.map(item => ({
                            name: item,
                            title: item,
                            data: item
                          }));
    tableColumns.push({ data: null });
	return dataTable = new DataTable(tableName, {
                data: data,
                searching: (tableName === '#gate-count-data-table') ? false : true,
                order:{name: "Submission ID", dir: "asc"},
                pageLength: 100,
                scrollX: true,
                scrollY: 450,
                paging: true,
                columns: tableColumns,
                columnDefs: [{
                  targets: 0,
                  width: 5
                }, {
                  targets: -1, // Target the last column
                  data: null, // Do not use any data for the delete button
                  render: function(data, type, row, meta) {
                      // Return the delete button HTML with inline click event handler
                      return `
                        <div style="display: flex; gap: 5px;">
                        <button onclick='editRow(this, ${JSON.stringify(tableName)})'><i class='fas fa-edit'></i></button>
                        <button onclick='deleteRow(this, false, ${JSON.stringify(tableName)})'><i class='fas fa-trash'></i></button>
                        <button onclick='addRow(this, ${JSON.stringify(tableName)}, ${JSON.stringify(meta.row)})'><i class='fas fa-plus'></i></button>
                        </div>
                        `;
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

    gateCountTable.clear();
    gateCountTable.rows.add(gateCountData)
    gateCountTable.draw();
    $(tableName).parent().scrollTop(scrollPosition);
    document.getElementById('input1').value = addedGateCount
    document.getElementById('input2').value = addedComputerLab
    calculateTotals();
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

function excelDateToJSDate(excelDate) {

  const epoch = new Date(1899, 11, 30); // Excel's epoch date (Dec 30, 1899)
  return new Date(epoch.getTime() + excelDate * 86399956.66); // Multiply by 86400000 to convert to milliseconds
}

function convertExcelDates(list) {
  return list.map(item => {
    // Convert each item's excelDate and add it as a readable date
    item[0] = localizedDate(excelDateToJSDate(item[0]));
    item[15] !== undefined && (item[15] = localizedDate(excelDateToJSDate(item[15])));
    return item;
  });
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
function deleteRow(button, cancelButton, tableName) {

    sortByIdDescending(gateCountData)
    let scrollPosition = $(tableName).parent().scrollTop();

    // Find the closest row of the button
    let row = button.closest('tr');
    let submissionID = row.cells[0].textContent;
    row.classList.add('highlighted');

    setTimeout(function() {
        if(cancelButton)
            gateCountTable.row(row).remove().draw();
        // Confirmation dialog
        else if (confirm("Are you sure you want to delete Submission = " + submissionID)) {
            // Get the row's data or submission ID (assuming "Submission ID" is in the first column)
           // Adjust this if the "Submission ID" is in a different column


                const indexToDelete = gateCountData.findIndex(item => item["Submission ID"] === Number(submissionID));
                console.log(gateCountData)
                if(indexToDelete < gateCountData.length - 1) {
                    gateCountData[indexToDelete+1]["Gate Count - Daily Total"] = gateCountData[indexToDelete -1]["Gate Count:"] - gateCountData[indexToDelete+1]["Gate Count:"]
                    gateCountData[indexToDelete+1]["Gate Count - Unique Head Count"] = gateCountData[indexToDelete +1]["Gate Count - Daily Total"]/2
                    gateCountData[indexToDelete+1]["Computer Lab - Daily Total"] = gateCountData[indexToDelete - 1]["Computer Lab"] - gateCountData[indexToDelete+1]["Computer Lab"]
                    gateCountData[indexToDelete+1]["Computer Lab - Unique Head Count"] = gateCountData[indexToDelete+1]["Computer Lab - Daily Total"]/2
                }
                gateCountData.splice(indexToDelete, 1)
                // Remove the row from DataTable and redraw the table
                //
                gateCountTable.clear(); // Clear the current data
                gateCountTable.rows.add(gateCountData); // Add the new data
                gateCountTable.draw(); // Redraw the table with the new data

        } else {
            row.classList.remove('highlighted');
        }
    }, 50)
    $(tableName).parent().scrollTop(scrollPosition);
    calculateTotals()
}

// Function to highlight the row and make it editable
function editRow(button, tableName) {
    sortByIdDescending(gateCountData)
    const row = button.closest('tr');
    let scrollPosition = $(tableName).parent().scrollTop();

    // Add a class to highlight the row
    row.classList.add('highlighted');
    let rowData = {}
    // Make each cell in the row editable
    const cells = row.querySelectorAll('td');

    cells.forEach((cell, index) => {
      if (index == 1 || index == 2 || index == 5 || index == 9) { // Skip the last column (buttons column)
        //console.log()
        const originalText = index === 1 ? convertTo24HourFormat(cell.textContent) : cell.textContent;
        console.log(originalText)
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
        <button class='btn btn-success' onclick='saveRow(this,${JSON.stringify(rowData)}, ${JSON.stringify(tableName)})'><i class='fas fa-check'></i></button>
        <button class='btn btn-secondary' onclick='cancelEdit(this, ${JSON.stringify(tableName)})'>
              <i class='fas fa-times'></i>
        </button>
        `;
    row.querySelector('td:last-child').innerHTML = saveButtonHtml;
    $(tableName).parent().scrollTop(scrollPosition);
}

function addRow(button, tableName, rowIndex) {
    sortByIdDescending(gateCountData)
    let table = $(tableName).DataTable();
    // Add a class to highlight the row
    let scrollPosition = $(tableName).parent().scrollTop();
    rowData = table.row(rowIndex).data()

   // Create an empty row (same number of columns as the table)

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
                <button class='btn btn-success' onclick='saveRow(this,${JSON.stringify(emptyRow)})'><i class='fas fa-check'></i></button>
                <button class='btn btn-secondary' onclick='deleteRow(this, true, ${JSON.stringify(tableName)})'>
                      <i class='fas fa-times'></i>
                </button>
                `;
            editRow.querySelector('td:last-child').innerHTML = saveButtonHtml;
        targetRow.parentNode.insertBefore(newRow, editRow);  // Insert before the target row
        $(tableName).parent().scrollTop(scrollPosition);
}

// Function to remove highlight and save edited values (this can be triggered when you want to save the edits)
function saveRow(button, rowData, tableName) {
    sortByIdDescending(gateCountData)
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
        const input = index === 1 ? cell.querySelector('.form-control') : cell.querySelector('input'); // Get the input element
        if (input) {
        console.log(index)
        console.log(gateCountHeaders)
          const columnName = gateCountHeaders[index]; // Get column name (assuming tableColumns array contains column names)
          // Map the input values to the corresponding properties
          switch (columnName) {
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
      const indexToUpdate = gateCountData.findIndex(item => item["Submission ID"] === Number(submissionId));
      if (indexToUpdate !== -1) {
        // Update the found item with the new rowData values
        gateCountData[indexToUpdate] = { ...gateCountData[indexToUpdate], ...rowData };
        gateCountData[indexToUpdate]["Gate Count - Daily Total"] = gateCountData[indexToUpdate - 1]["Gate Count:"] - gateCountData[indexToUpdate]["Gate Count:"]
        gateCountData[indexToUpdate]["Gate Count - Unique Head Count"] = gateCountData[indexToUpdate]["Gate Count - Daily Total"]/2
        gateCountData[indexToUpdate]["Computer Lab - Daily Total"] = gateCountData[indexToUpdate - 1]["Computer Lab"] - gateCountData[indexToUpdate]["Computer Lab"]
        gateCountData[indexToUpdate]["Computer Lab - Unique Head Count"] = gateCountData[indexToUpdate]["Computer Lab - Daily Total"]/2

        console.log("Updated Data:", gateCountData[indexToUpdate]); // Log updated data for debugging
      } else {
        // Find the position to insert the new element
        for (let i = 0; i < gateCountData.length - 1; i++) {
            if (gateCountData[i]["Submission ID"] > submissionId && gateCountData[i + 1]["Submission ID"] < submissionId) {
                console.log("index = ", i)
                rowData["Gate Count - Daily Total"] = gateCountData[i]["Gate Count:"] - rowData["Gate Count:"]
                rowData["Gate Count - Unique Head Count"] = rowData["Gate Count - Daily Total"]/2
                rowData["Computer Lab - Daily Total"] = gateCountData[i]["Computer Lab"] - rowData["Computer Lab"]
                rowData["Computer Lab - Unique Head Count"] = rowData["Computer Lab - Daily Total"]/2
                gateCountData[i+1]["Gate Count - Daily Total"] = rowData["Gate Count:"] - gateCountData[i+1]["Gate Count:"]
                gateCountData[i+1]["Gate Count - Unique Head Count"] = gateCountData[i+1]["Gate Count - Daily Total"]/2
                gateCountData[i+1]["Computer Lab - Daily Total"] = rowData["Computer Lab"] - gateCountData[i+1]["Computer Lab"]
                gateCountData[i+1]["Computer Lab - Unique Head Count"] = gateCountData[i+1]["Computer Lab - Daily Total"]/2
                gateCountData.splice(i + 1, 0, rowData);  // Insert the new element between i and i+1
                break;
            }
        }
      }

    gateCountTable.clear(); // Clear the current data
    gateCountTable.rows.add(gateCountData); // Add the new data
    gateCountTable.draw(); // Redraw the table with the new data
    $(tableName).parent().scrollTop(scrollPosition);

    calculateTotals();
}

// Function to cancel editing and revert the changes
function cancelEdit(button, tableName) {
 sortByIdDescending(gateCountData)
  let scrollPosition = $(tableName).parent().scrollTop();
  const row = button.closest('tr');
  row.classList.remove('highlighted');

  gateCountTable.clear(); // Clear the current data
  gateCountTable.rows.add(gateCountData); // Add the new data
  gateCountTable.draw(); // Redraw the table with the new data

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

      if (period === "p.m." && hours !== 12) {
        hours += 12; // Add 12 to hours for PM, except for 12 PM
      } else if (period === "a.m." && hours === 12) {
        hours = 0; // Convert 12 AM to 00:00
      }

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
    if (hours > 12) {
      hours -= 12;
    } else if (hours === 0) {
      hours = 12; // Midnight (00:00) is 12 a.m.
    }

    // Format the time in 12-hour format with leading zeros if needed
    const formattedTime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:00 ${period}`;

    // Combine date and formatted time
    return `${date}, ${formattedTime}`;
  }

function calculateTotals() {
    totalGateCount = 0;
    totalComputerLab = 0;
    totalGateCountAverage = 0;
    totalLabAverage = 0
    // Loop through each item and accumulate the cost and profit
    gateCountData.forEach(item => {
      totalGateCount = Number(totalGateCount) + Number(item["Gate Count - Unique Head Count"]);   // Sum the cost
      totalComputerLab = Number(totalComputerLab) + Number(item["Computer Lab - Unique Head Count"]); // Sum the profit
    });

    document.getElementById('gate-count-total').innerHTML = `Total Unique Head Count: ${totalGateCount}`;
    document.getElementById('computer-lab-total').innerHTML = `Total Unique Head Count: ${totalComputerLab}`;

    let totalDays = document.getElementById('total-days').value
    lastYear = document.getElementById('last-year').value

    totalGateCountAverage = (totalGateCount/totalDays).toFixed(2)
    totalLabAverage = (totalComputerLab/totalDays).toFixed(2)

    document.getElementById('gate-count-average').innerHTML = `Average per day: ${totalGateCountAverage}`;
    document.getElementById('computer-lab-average').innerHTML = `Average per day: ${totalLabAverage}`;

    let changePercentage = (((totalGateCount/lastYear)-1)*100).toFixed(2)
    changeText = `${Math.abs(changePercentage)}% ${changePercentage < 0 ? 'decrease' : 'increase'}`;
    document.getElementById('overallCount').innerHTML = `Increase / Decrease: ${changeText}`;

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
        console.log(headerRow)
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

// Function to close the side tab
function closeSideTab() {
  document.querySelector('.side-tab').style.display = 'none';
}

// Function to change the content based on the clicked module
function changeContent(module) {
  let contentArea = document.getElementById("content-area");
  let fileToLoad = "";
  // Define the content for each module
  let content = "";
  switch (module) {
    case 'module1':
      content = "<h2>Module 1 Content</h2><p>This is the content for Module 1.</p>";
      break;
    case 'module2':
      fileToLoad = "modules/gate-count-module.html";
      break;
    case 'module3':
      content = "<h2>Module 3 Content</h2><p>This is the content for Module 3.</p>";
      break;
    case 'module4':
      content = "<h2>Module 4 Content</h2><p>This is the content for Module 4.</p>";
      break;
    case 'module5':
      content = "<h2>Module 5 Content</h2><p>This is the content for Module 5.</p>";
      break;
    default:
      content = "<h2>Welcome! Please select a module.</h2>";
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

}

// Load the 'gate-count-module' content immediately on page load
window.onload = function() {
  changeContent('module2');
};

// Function to trigger the file input dialog
function triggerFileInput() {
  document.getElementById("excel-upload").click();
}
