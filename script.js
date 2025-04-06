// Load the 'gate-count-module' content immediately on page load
window.onload = function() {
  Chart.register(ChartDataLabels);
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
};

// Function to handle the file upload and data conversion
function loadFile(file) {
   // const file = event.target.files[0];
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
            const rawData = displayList(jsonData)

            const splitData = rawData.flatMap(record => splitRecord(record));

            const newData = processRecords(splitData);

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

                  // If the calculated daily total is negative, assign the previous item's gate count
                  const adjustedGateCountDailyTotal = gateCountDailyTotal < 0
                    ? previousItem ? previousItem["Gate Count:"] : 0
                    : gateCountDailyTotal;


                  // Calculate "Gate Count - Unique Head Count" as half of "Gate Count - Daily Total"
                  const gateCountUniqueHeadCount = adjustedGateCountDailyTotal / 2;

                  // Calculate "Computer Lab - Daily Total" as the difference from the previous "Computer Lab"
                  const computerLabDailyTotal = previousItem
                    ? previousItem["Computer Lab"] - item["Computer Lab"]
                    : 0; // For the first item, default to 0

                  const adjustedComputerLabDailyTotal = computerLabDailyTotal < 0
                                      ? previousItem ? previousItem["Computer Lab"] : 0
                                      : computerLabDailyTotal;

                  // Calculate "Computer Lab - Unique Head Count" as half of "Computer Lab - Daily Total"
                  const computerLabUniqueHeadCount = adjustedComputerLabDailyTotal / 2;

                  return {
                    "Submission ID": item["Submission ID"],
                    "Submitted": item["Submitted"],
                    "Gate Count:": item["Gate Count:"],
                    "Gate Count - Daily Total": adjustedGateCountDailyTotal,
                    "Gate Count - Unique Head Count": gateCountUniqueHeadCount,
                    "Computer Lab": item["Computer Lab"],
                    "Computer Lab - Daily Total": adjustedComputerLabDailyTotal,
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

			rovingData.forEach(item => {
               item["Roving Time"] = item["Roving Time"] && item["Roving Time"].trim() !== ""
                   ? correctTime(roundToNearestHalfHour(item["Roving Time"]))
                   : roundToNearestHalfHour(item["Submitted"]) ; // Assign an empty string if Submitted is undefined or an empty string

               //item["Roving Time"] = roundToNearestHalfHour(item["Submitted"])
            });

            console.log(rovingData)

			loanableTechData = newData.filter(item =>
			   item["Type of Inquiry:"] == "Loanable Tech"
			)

			// Helper function to capitalize the first letter of each word
            function capitalizeWords(str) {
                return str.replace(/\b\w/g, (char) => char.toUpperCase());
            }

            distinctiveInquiryData = [
                ...new Set(
                    newData
                        .filter(item =>
                            item["Type of Inquiry:"] !== "" &&
                            item["Type of Inquiry:"] !== "Gate Count" &&
                            item["Type of Inquiry:"] !== "Roving" &&
                            (item["Additional Information:"]?.trim() !== '' || item["Subject(s) of Inquiry:"]?.trim() !== '')
                        )
                        .map(item => {
                            // Choose the non-empty value between "Additional Information:" and "Subject(s) of Inquiry:"
                            const inquiryType = item["Type of Inquiry:"];
                            const inquiry = item["Additional Information:"]?.trim() || item["Subject(s) of Inquiry:"]?.trim();
                            return { "Type": inquiryType, "Inquiry": inquiry };
                        })
                )
            ];

            // Remove case-sensitive duplicates
            distinctiveInquiryData = [
                ...new Set(
                    distinctiveInquiryData.map(item => `${item.Type.toLowerCase()}-${item.Inquiry.toLowerCase()}`)
                )
            ].map(item => {
                // Extract the original objects back from the unique string
                const [type, inquiry] = item.split('-');
                return { "Type": type, "Inquiry": inquiry };
            });

            // Capitalize the first letter of each word in Type and Inquiry fields
            distinctiveInquiryData = distinctiveInquiryData.map(item => ({
                "Type": capitalizeWords(item.Type),
                "Inquiry": capitalizeWords(item.Inquiry)
            }));

            setActiveTab(document.getElementById('dashboard'));
        };
        reader.readAsBinaryString(file);
    }
}

function correctTime(input) {
  // Split the input into date and time parts
  let [datePart, timePart] = input.split(',').map(str => str.trim());

  // Convert to 24-hour format using Date
  let originalDate = new Date(`${datePart} ${timePart.replace(/\./g, '')}`);

  // Get hours and minutes
  let hours = originalDate.getHours();
  let minutes = originalDate.getMinutes();

  // Define the time boundaries
  const lowerBound = new Date(`${datePart} 07:30`);
  const upperBound = new Date(`${datePart} 19:30`);

  if (originalDate < lowerBound || originalDate > upperBound) {
    // If outside the range, add 12 hours (only if it's AM)
    if (hours < 12) {
      originalDate.setHours(hours + 12);
    }
  }

  // Format back to desired string
  let correctedHours = originalDate.getHours() % 12 || 12;
  let correctedMinutes = String(originalDate.getMinutes()).padStart(2, '0');
  let correctedSeconds = String(originalDate.getSeconds()).padStart(2, '0');
  let meridiem = originalDate.getHours() >= 12 ? 'p.m.' : 'a.m.';

  let fixedTime = `${datePart}, ${correctedHours}:${correctedMinutes}:${correctedSeconds} ${meridiem}`

  return fixedTime;
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

    // Select all elements with the class 'my-class'
    var elements = document.getElementsByClassName('table-wrapper');

    // Loop through the elements and set the background color to white
    for (var i = 0; i < elements.length; i++) {
        elements[i].style.backgroundColor = 'white';
    }

   	dataTable = new DataTable(tableName, {
        data: data,
        searching: true,
        pageLength: 100,
        scrollX: false,
        scrollY: 650,
        paging: false,
        columns: tableColumns,
        columnDefs: [{
            "targets": "_all",  // Disable sorting on Name and Country columns
            "orderable": false
        }]
    })

    dataTable.search(currentSearch)

    dataTable.page(currentPage)

    dataTable.draw(false);

    return dataTable
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

    // Convert Excel date to milliseconds (Excel uses a float)
    const jsDate = new Date(epoch.getTime() + excelDate *  86399956.66); // Multiply by 86400000 to convert to milliseconds

    // Get the time zone offset for the date, considering DST
    const timeZoneOffset = jsDate.getTimezoneOffset() / 60; // In hours, positive for behind UTC, negative for ahead of UTC

    // If the offset is -6, adjust for -6 (DST)
    if (timeZoneOffset === 6) jsDate.setHours(jsDate.getHours() - 1); // Daylight Saving Time (UTC-6)

    return jsDate;
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

function convertDateFormat(dateStr) {

    // Remove the comma and trim any extra spaces
    dateStr = dateStr.replace(',', '').trim();

    // Split the string into date and time parts
    let [datePart, timePart, ampm] = dateStr.split(' ');
    // Convert the AM/PM part to uppercase
    ampm = ampm.toUpperCase().replace(/\./g, ''); ;

    // Reassemble the date and time in the desired format
    const newFormat = `${datePart} ${timePart} ${ampm}`;

    return newFormat;
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
		{ header: 'Type of Digital Support Inquiry:', key: 'Type of  Digital Support Inquiry:'},
		{ header: 'Technology Item Type:', key: 'Technology Item Type:'},
		{ header: 'Software/Application Type:', key: 'Software/Application Type:'},
		{ header: "Student's Program", key: "Student's Program"},
		{ header: 'Year of Program', key: 'Year of Program'},
		{ header: 'Was Tech available at the time of request?', key: 'Was Tech available at the time of request?'},
		{ header: 'Subject(s) of Inquiry:', key: 'Subject(s) of Inquiry:'},
		{ header: 'Additional Information:', key: 'Additional Information:'},
    ];

    wsRefStats.columns.forEach(column => {
          column.width = 17; // Set width to ~115px for each column
      });

	const refStatsHeaders = ['Submission ID', 'Submitted', 'Method of Inquiry:', 'Type of Inquiry:', 'Type of Reference:', 'Type of Facilitative Inquiry:',
                'Type of  Digital Support Inquiry:', 'Technology Item Type:', 'Software/Application Type:', "Student's Program", 'Year of Program',
                'Was Tech available at the time of request?', 'Subject(s) of Inquiry:', 'Additional Information:']
    exportRefStats.forEach(item => {
        wsRefStats.addRow(item);
    });

    const wsGateCount = workbook.addWorksheet('Gate Count');
    wsGateCount.columns = [
        { header: 'Submitted', key: 'Submitted', width: 22},
        { header: 'Gate Count', key: 'Gate Count:', width: 22},
        { header: 'Computer Lab', key: 'Computer Lab', width: 22},
        { header: 'Subject(s) of Inquiry', key: 'Subject(s) of Inquiry:', width: 22},
        { header: 'Additional Information', key: 'Additional Information:', width: 22}
    ];
    gateCountData.forEach(item => {
        wsGateCount.addRow(item);
    });

    // Create the second worksheet 'Gate Count Summary'
    const wsGateCountSummary = workbook.addWorksheet('Gate Count Summary');
    wsGateCountSummary.columns = [
        { header: 'Date', key: 'Submitted', width: 22},
        { header: 'Gate Count', key: 'Gate Count:', width: 11},
        { header: 'Daily Total', key: 'Gate Count - Daily Total', width: 11},
        { header: 'Unique Head Count (Daily Total/2)', key: 'Gate Count - Unique Head Count', width: 21},
        { width: 5},
        { header: 'Computer Lab', key: 'Computer Lab', width: 11},
        { header: 'Daily Total', key: 'Computer Lab - Daily Total', width: 11},
        { header: 'Unique Head Count (Daily Total/2)', key: 'Computer Lab - Unique Head Count', width: 21},
        { width: 5},
        { header: 'Subject(s) of Inquiry', key: 'Subject(s) of Inquiry:', width: 25},
        { header: 'Additional Information', key: 'Additional Information:', width: 25}
    ];

    /*wsGateCountSummary.columns.forEach(column => {
              column.width = 22; // Set width to ~115px for each column
          });*/

    gateCountData.forEach(item => {
        wsGateCountSummary.addRow(item);
    });
    let lastRow = wsGateCountSummary.lastRow
    let nextRow = 0
    // Assign the formula to column C starting from row 2
    for (let row = 2; row <= lastRow.number; row++) {
      nextRow = row === lastRow.number ? row + 2 : row + 1
      wsGateCountSummary.getCell(`C${row}`).value = { formula: `IF(B${nextRow}-B${row}<0,B${nextRow},B${nextRow}-B${row})`};
      wsGateCountSummary.getCell(`D${row}`).value = { formula: `C${row}/2`};
      wsGateCountSummary.getCell(`G${row}`).value = { formula: `IF(F${nextRow}-F${row}<0,F${nextRow},F${nextRow}-F${row})`};
      wsGateCountSummary.getCell(`H${row}`).value = { formula: `G${row}/2`};
    }

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

    wsRoving.columns.forEach(column => {
          column.width = 12; // Set width to ~115px for each column
      });
    wsRoving.columns[0].width = 22

    exportRoving.forEach(item => {
        wsRoving.addRow(item);
    });

    lastRow = wsRefStats.lastRow


    let totalRow1 = wsRefStats.insertRow(lastRow.number + 6, ['Type of Inquiry', 'Loanable Tech', '', '', 'Type of Reference Inquiry', 'Citation Help', '', '', 'Loanable Tech', 'Calculator', '']);
        totalRow1.getCell(3).value = { formula: `COUNTIF(C2:C${lastRow.number}, "Loanable Tech")` };
        totalRow1.getCell(7).value = { formula: `COUNTIF(D2:D${lastRow.number}, "Citation Help")` };
        totalRow1.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Calculator")` };
        // Loop through each cell in the inserted row and call customMergedCell(cell) for each
        totalRow1.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow2 = wsRefStats.insertRow(lastRow.number + 7, ['', 'Digital Support', '', '', '', 'Find a Resource (print or online)', '', '', '', 'Camera', '']);
        totalRow2.getCell(3).value = { formula: `COUNTIF(C2:C${lastRow.number}, "Digital Support")` };
        totalRow2.getCell(7).value = { formula: `COUNTIF(D2:D${lastRow.number}, "Find a Resource (print or online)")` };
        totalRow2.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Camera")` };
        totalRow2.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow3 = wsRefStats.insertRow(lastRow.number + 8, ['', 'Basic Reference', '', '', '', 'Database Help', '', '', '', 'Charger, Adapter, etc.', '']);
        totalRow3.getCell(3).value = { formula: `COUNTIF(C2:C${lastRow.number}, "Basic Reference")` };
        totalRow3.getCell(7).value = { formula: `COUNTIF(D2:D${lastRow.number}, "Database Help")` };
        totalRow3.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Charger, Adapter, etc.")` };
        totalRow3.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow4 = wsRefStats.insertRow(lastRow.number + 9, ['', 'Complex Reference', '', '', '', 'Copyright', '', '', '', 'Chromebooks', '']);
        totalRow4.getCell(3).value = { formula: `COUNTIF(C2:C${lastRow.number}, "Complex Reference")` };
        totalRow4.getCell(7).value = { formula: `COUNTIF(D2:D${lastRow.number}, "Copyright")` };
        totalRow4.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Chromebooks")` };
        totalRow4.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow5 = wsRefStats.insertRow(lastRow.number + 10, ['', 'Facilitative', '', '', '', 'Other', '', '', '', 'Chromecast', '']);
        totalRow5.getCell(3).value = { formula: `COUNTIF(C2:C${lastRow.number}, "Facilitative")` };
        totalRow5.getCell(7).value = { formula: `COUNTIF(D2:D${lastRow.number}, "Other")` };
        totalRow5.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Chromecast")` };
        totalRow5.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow6 = wsRefStats.insertRow(lastRow.number + 11, ['', '', '', '', '', '', '', '', '', 'DVD Player', '']);
        totalRow6.getCell(3).value = { formula: `SUM(C${lastRow.number + 6}:C${lastRow.number + 10})` };
        totalRow6.getCell(7).value = { formula: `SUM(G${lastRow.number + 6}:G${lastRow.number + 10})` };
        totalRow6.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "DVD Player")` };
        totalRow6.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow7 = wsRefStats.insertRow(lastRow.number + 12, ['Method of Inquiry', 'Chat', '', '', 'Type of Digital Support Inquiry', 'Document Assistance (e.g. Microsoft Word, Excel, PDF, Google Docs, etc.)', '', '', '', 'Headphones', '']);
        totalRow7.getCell(3).value = { formula: `COUNTIF(B2:B${lastRow.number}, "Chat")` };
        totalRow7.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Document Assistance (e.g. Microsoft Word, Excel, PDF, Google Docs, etc.)")` };
        totalRow7.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Headphones")` };
        totalRow7.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow8 = wsRefStats.insertRow(lastRow.number + 13, ['', 'In Person', '', '', '', 'Internet/Wifi Connectivity', '', '', '', 'Keyboard', '']);
        totalRow8.getCell(3).value = { formula: `COUNTIF(B2:B${lastRow.number}, "In Person")` };
        totalRow8.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Internet/Wifi Connectivity")` };
        totalRow8.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Keyboard")` };
        totalRow8.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow9 = wsRefStats.insertRow(lastRow.number + 14, ['', 'Phone', '', '', '', 'Keyano Account Access (e.g. Webmail, Moodle, or Self-Service)', '', '', '', 'Laptops', '']);
        totalRow9.getCell(3).value = { formula: `COUNTIF(B2:B${lastRow.number}, "Phone")` };
        totalRow9.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Keyano Account Access (e.g. Webmail, Moodle, or Self-Service)")` };
        totalRow9.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Laptops")` };
        totalRow9.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow10 = wsRefStats.insertRow(lastRow.number + 15, ['', 'Email', '', '', '', 'LMS (Moodle. McGraw, MyLAB IT)', '', '', '', 'MFA Token', '']);
        totalRow10.getCell(3).value = { formula: `COUNTIF(B2:B${lastRow.number}, "Email")` };
        totalRow10.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "LMS (Moodle. McGraw, MyLAB IT)")` };
        totalRow10.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "MFA Token")` };
        totalRow10.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow11 = wsRefStats.insertRow(lastRow.number + 16, ['', 'Form Submission', '', '', '', 'Online Navigation (e.g. opening a browser or searching in Google)', '', '', '', 'Power Bank', '']);
        totalRow11.getCell(3).value = { formula: `COUNTIF(B2:B${lastRow.number}, "Form Submission")` };
        totalRow11.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Online Navigation (e.g. opening a browser or searching in Google)")` };
        totalRow11.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Power Bank")` };
        totalRow11.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow12 = wsRefStats.insertRow(lastRow.number + 17, ['', '', '', '', '', 'Print/Scan/Copy Assistance or Troubleshooting', '', '', '', 'Projector', '']);
        totalRow12.getCell(3).value = { formula: `SUM(C${lastRow.number + 12}:C${lastRow.number + 16})` };
        totalRow12.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Print/Scan/Copy Assistance or Troubleshooting")` };
        totalRow12.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Projector")` };
        totalRow12.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow13 = wsRefStats.insertRow(lastRow.number + 18, ['Type of Facilitative Inquiry', 'Interlibrary Loans/Requests/Holds', '', '', '', 'Software (M365, Respondus, Safe Exam, etc.)', '', '', '', 'SAD Light', '']);
        totalRow13.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Interlibrary Loans/Requests/Holds")` };
        totalRow13.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Software (M365, Respondus, Safe Exam, etc.)")` };
        totalRow13.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "SAD Light")` };
        totalRow13.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow14 = wsRefStats.insertRow(lastRow.number + 19, ['', 'General Library Information (e.g. hours, borrowing period, etc.)', '', '', '', 'Other', '', '', '', 'WebCam', '']);
        totalRow14.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "General Library Information (e.g. hours, borrowing period, etc.)")` };
        totalRow14.getCell(7).value = { formula: `COUNTIF(F2:F${lastRow.number}, "Other")` };
        totalRow14.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "WebCam")` };
        totalRow14.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow15 = wsRefStats.insertRow(lastRow.number + 20, ['', 'Library Account (e.g. pin, renewals, fines, etc.)', '', '', '', '', '', '', '', 'Wireless Mouse', '']);
        totalRow15.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Library Account (e.g. pin, renewals, fines, etc.)")` };
        totalRow15.getCell(7).value = { formula: `SUM(G${lastRow.number + 12}:G${lastRow.number + 19})` };
        totalRow15.getCell(11).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Wireless Mouse")` };
        totalRow15.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow16 = wsRefStats.insertRow(lastRow.number + 21, ['', 'Referral/Directional (External - Bookstore, Registrar, Academic Success Centre, etc.)', '', '', 'Laptop Requests:', '', 'Available at time of request:', '', '', '', '']);
        totalRow16.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Referral/Directional (External - Bookstore, Registrar, Academic Success Centre, etc.)")` };
        totalRow16.getCell(6).value = { formula: `COUNTIF(G2:G${lastRow.number}, "Laptops")` };
        totalRow16.getCell(8).value = { formula: `COUNTIFS(G2:G${lastRow.number}, "Laptops", K2:K${lastRow.number}, "Yes")` };
        totalRow16.getCell(11).value = { formula: `SUM(K${lastRow.number + 6}:K${lastRow.number + 20})` };
        totalRow16.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow17 = wsRefStats.insertRow(lastRow.number + 22, ['', 'Referral/Directional (In Library - BAL, Copyright, EdTech, Instruction, etc.)', '', '', '', '', 'Unavailable at time of request:', '', '', '', '']);
        totalRow17.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Referral/Directional (In Library - BAL, Copyright, EdTech, Instruction, etc.)")` };
        totalRow17.getCell(8).value = { formula: `COUNTIFS(G2:G${lastRow.number}, "Laptops", K2:K${lastRow.number}, "No")` };
        totalRow17.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow18 = wsRefStats.insertRow(lastRow.number + 23, ['', 'Community User', '', '', '', '', '', '', '', '', '']);
        totalRow18.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Community User")` };
        totalRow18.getCell(8).value = { formula: `SUM(H${lastRow.number + 21}:H${lastRow.number + 22})` };
        totalRow18.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow19 = wsRefStats.insertRow(lastRow.number + 24, ['', 'Supplies (e.g. stapler, pen, hole punch, etc.)', '', '', '', '', '', '', '', '', '']);
        totalRow19.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Supplies (e.g. stapler, pen, hole punch, etc.)")` };
        totalRow19.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow20 = wsRefStats.insertRow(lastRow.number + 25, ['', 'Study Room', '', '', '', '', '', '', '', '', '']);
        totalRow20.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Study Room")` };
        totalRow20.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow21 = wsRefStats.insertRow(lastRow.number + 26, ['', 'Accessible Format Request', '', '', '', '', '', '', '', '', '']);
        totalRow21.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Accessible Format Request")` };
        totalRow21.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow22 = wsRefStats.insertRow(lastRow.number + 27, ['', 'Reserve Request', '', '', '', '', '', '', '', '', '']);
        totalRow22.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Reserve Request")` };
        totalRow22.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow23 = wsRefStats.insertRow(lastRow.number + 28, ['', 'Scan-on-Demand', '', '', '', '', '', '', '', '', '']);
        totalRow23.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Scan-on-Demand")` };
        totalRow23.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow24 = wsRefStats.insertRow(lastRow.number + 29, ['', 'Other', '', '', '', '', '', '', '', '', '']);
        totalRow24.getCell(3).value = { formula: `COUNTIF(E2:E${lastRow.number}, "Other")` };
        totalRow24.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

    let totalRow25 = wsRefStats.insertRow(lastRow.number + 30, ['', '', '', '', '', '', '', '', '', '', '']);
        totalRow25.getCell(3).value = { formula: `SUM(C${lastRow.number + 18}:C${lastRow.number + 29})` };
        totalRow25.eachCell((cell, colNumber) => {
            customMergedCell(cell);
        });

        wsRefStats.mergeCells(`A${lastRow.number + 6}:A${lastRow.number + 10}`);
        customMergedCell(wsRefStats.getCell(`A${lastRow.number + 6}`))

        wsRefStats.mergeCells(`A${lastRow.number + 12}:A${lastRow.number + 16}`);
        customMergedCell(wsRefStats.getCell(`A${lastRow.number + 12}`))

        wsRefStats.mergeCells(`A${lastRow.number + 18}:A${lastRow.number + 29}`);
        customMergedCell(wsRefStats.getCell(`A${lastRow.number + 18 }`))

        wsRefStats.mergeCells(`E${lastRow.number + 6}:E${lastRow.number + 10}`);
        customMergedCell(wsRefStats.getCell(`E${lastRow.number + 6}`))

        wsRefStats.mergeCells(`E${lastRow.number + 12}:E${lastRow.number + 19}`);
        customMergedCell(wsRefStats.getCell(`E${lastRow.number + 12}`))

        wsRefStats.mergeCells(`E${lastRow.number + 21}:E${lastRow.number + 22}`);
        customMergedCell(wsRefStats.getCell(`E${lastRow.number + 21}`))

        wsRefStats.mergeCells(`F${lastRow.number + 21}:F${lastRow.number + 22}`);
        customMergedCell(wsRefStats.getCell(`F${lastRow.number + 21}`))

        wsRefStats.mergeCells(`I${lastRow.number + 6}:I${lastRow.number + 20}`);
        customMergedCell(wsRefStats.getCell(`I${lastRow.number + 6}`))


    lastRow = wsGateCountSummary.lastRow
    wsGateCountSummary.insertRow(lastRow.number + 2, ['Date', addedGateCount, '', '', '', addedComputerLab]);
    let gateCountTotalRow = wsGateCountSummary.insertRow(lastRow.number + 5, ['','', 'Total', '', '', '', 'Total Lab', totalComputerLab])
        gateCountTotalRow.getCell(4).value = { formula: `SUM(D2:D${lastRow.number})` };
        gateCountTotalRow.getCell(8).value = { formula: `SUM(H2:H${lastRow.number})` };
    let gateCountAverageDay = wsGateCountSummary.insertRow(lastRow.number + 6, ['','', 'Average per day:', parseFloat(totalGateCountAverage), '', '', 'Average per day', parseFloat(totalLabAverage)]);
        gateCountAverageDay.getCell(4).value = { formula: `D${lastRow.number + 5}/${totalDays}` };
        gateCountAverageDay.getCell(4).numFmt = '0.00';
        gateCountAverageDay.getCell(8).value = { formula: `H${lastRow.number + 5}/${totalDays}` };
        gateCountAverageDay.getCell(8).numFmt = '0.00';
    wsGateCountSummary.insertRow(lastRow.number + 9, ['', '', 'Year Over Year Comparison']);
    wsGateCountSummary.insertRow(lastRow.number + 10, ['','', 'Last year', lastYear]);
    let changeRow = wsGateCountSummary.insertRow(lastRow.number + 11, ['','', 'Increase / Decrease:', '']);
        changeRow.getCell(4).value = { formula: `ABS(ROUND(((D${lastRow.number + 5} - D${lastRow.number + 10}) / D${lastRow.number + 10}) * 100, 0)) & "% " & IF(((D${lastRow.number + 5} - D${lastRow.number + 10}) / D${lastRow.number + 10}) > 0, "Increase", IF(((D${lastRow.number + 5} - D${lastRow.number + 10}) / D${lastRow.number + 10}) < 0, "Decrease", "No Change"))` };
        changeRow.getCell(4).alignment = { horizontal: 'right'};
    // Apply styles to the first row (header) for all sheets
    [wsRefStats, wsGateCount, wsGateCountSummary, wsRoving].forEach(sheet => {
        const headerRow = sheet.getRow(1);
        headerRow.eachCell((cell, colNumber) => {
            const value = cell.value ? cell.value.toString() : '';
            cell.font = { bold: true }; // Bold text
            cell.alignment = { horizontal: 'center', vertical: 'middle' }; // Center alignment
            cell.alignment.wrapText = true; // Enable text wrapping
        });

        const numRows = sheet.rowCount; // Get the number of rows in the sheet

         // Loop through all rows, starting from row 2 (skip header)
         for (let rowNum = 2; rowNum <= numRows; rowNum++) {
             const cell = sheet.getCell(rowNum, 1); // Get cell in the first column of the current row

             // Check if the cell value is a string (date in string format)
             //if (cell.value && (cell.value.includes("a.m.") || cell.value.includes("p.m."))) {
             if(cell.value && typeof cell.value === 'string' && (cell.value.includes("a.m.") || cell.value.includes("p.m."))) {
                 // Try to parse the string into a Date object
                const parsedDate = new Date(convertDateFormat(cell.value));
                if (!isNaN(parsedDate.getTime())) { // Ensure it's a valid Date
		     const timeZoneOffset = parsedDate.getTimezoneOffset() / 60; 
		     if (timeZoneOffset === 7) { 
			parsedDate.setHours(parsedDate.getHours() - 1); // Adjust for DST
        	     }
                     parsedDate.setHours(parsedDate.getHours() - 6);
			 
                     cell.value = parsedDate // Set the cell value to the Date object

                     cell.numFmt = 'yyyy-mm-dd h:mm'; // Apply custom date format
                 }
             }
         }

        // Freeze the top row (header)
        sheet.views = [
            {
                state: 'frozen',
                ySplit: 1, // Freeze row 1 (index 0 is the first row)
                topLeftCell: 'A2', // Ensure that the first row is frozen and starting from A2
            }
        ];
    });

    // Enable filter for the header row
    wsRefStats.autoFilter = {
        from: { row: 1, column: 1 }, // Enable autofilter for the whole header row (row 1)
        to: { row: 1, column: wsRefStats.columnCount } // End at the last column
    };

    // Apply auto column width for each worksheet
    [wsRefStats, wsGateCount, wsGateCountSummary, wsRoving].forEach((worksheet) => {
        worksheet.columns.forEach(column => {
            if (column.width > 10) {
                column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
                    if (rowNumber > 1) {  // Exclude the header row from wrap text
                        cell.alignment = cell.alignment || {}; // Ensure alignment object exists
                        //cell.alignment.wrapText = true;  // Enable wrap text for non-header rows
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

    sortByIdDescending(gateCountData)
    sortByIdDescending(rovingData)
    sortByIdDescending(refStats)
}

function customMergedCell(mergedCell) {

    if(mergedCell.value !== '' && !mergedCell.value.formula?.includes("SUM")) {
        mergedCell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    }

    if (mergedCell.value !== '') {
        mergedCell.alignment = {
          horizontal: 'center', // Horizontal alignment
          vertical: 'middle',   // Vertical alignment
        };
    }

    if (mergedCell.value.formula?.includes("SUM")) {
        mergedCell.alignment = {
            vertical: 'top',  // Align text to the top
        };

        mergedCell.font = {
            italic: true,        // Make the text italic
            color: { argb: '808080' },  // Set font color to #808080 (gray)
        };
    }


    if (!mergedCell.value.formula && mergedCell.value !== '') {
        mergedCell.alignment.wrapText = true

        mergedCell.font = {
            bold: true           // Make the text bold
        };
    }

    if (!mergedCell.value.formula && mergedCell.value &&
        (mergedCell.value.toLowerCase().includes("time of request") ||
         mergedCell.value.toLowerCase().includes("laptop request"))) {

        mergedCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E2EFDA' }
        };
        let nextCell = mergedCell.worksheet.getCell(mergedCell.row, mergedCell.col + 1)

            nextCell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E2EFDA' }
          };
    }

    if (mergedCell.value !== '' && !mergedCell.value.formula?.includes("SUM") && !mergedCell.fill) {
            mergedCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'CAEDFB' }
            };
        }

}

function sortByIdAscending(data) {
    return data.sort((a, b) => a["Submission ID"] - b["Submission ID"]);
}

function sortByIdDescending(data) {
    return data.sort((a, b) => b["Submission ID"] - a["Submission ID"]);
}

// Function to toggle the active class when a tab is clicked
function setActiveTab(selectedTab) {
    if(!selectedTab.classList.contains("active")) {
        let contentArea = document.getElementById("content-area");
        let fileToLoad = "";
        // Define the content for each module
        let content = "";
        switch (selectedTab.innerText) {
            case 'Dashboard':
                fileToLoad = "modules/dashboard-module.html";
                tableName = "#distinctive-inquiries-data-table"
                headers = ["Type", "Inquiry"],
                data = distinctiveInquiryData
                break;
            case 'Roving Count':
                fileToLoad = "modules/roving-count-module.html";
                headers = rovingHeaders
                data = rovingData
                tableName = "#roving-data-table"
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
            if(selectedTab.innerText === "Dashboard") loadDashBoard()
            else if (selectedTab.innerText == "Roving Count") initializeRovingCountPage()
        }, 100)
        // Remove active class from all tabs
        const tabs = document.querySelectorAll('.side-tab ul li');
        tabs.forEach(tab => {
            tab.classList.remove('active');
        });

        // Add active class to the clicked tab
        selectedTab.classList.add('active');
    }
}

function loadDashBoard() {
    let chartData = [];
    chartData = filterData("Type of Facilitative Inquiry:", typeOfFacilitativeInquiry.concat("Other"))
    loadCharts("facilitative-chart", chartData)
    chartData = filterData("Type of Reference:", typeOfReference.concat("Other"))
    loadCharts("reference-chart", chartData)
    chartData = filterData("Type of  Digital Support Inquiry:", typeOfDigitalSupportInquiry.concat("Other"))
    loadCharts("digital-support-chart", chartData)
    chartData = filterData("Technology Item Type:", technologyType)
    loadCharts("loanable-tech-chart", chartData)
    chartData = filterData("Technology Item Type:", technologyType)
    availabilityChart("availability-chart", chartData)
    chartData = filterData("Student's Program", programs)
    loadCharts("patron-program-chart", chartData)
}

function filterData(key, items) {
    return items.map(item => {
        let count = refStats.filter(record => {
                      if (key === "Student's Program") {
                        return record["Technology Item Type:"] !== "" && record[key].includes(item);  // Use includes() for other keys
                      } else {
                        return record[key] === item;  // Use strict equality for "abc"
                      }
                    }).length;

        let available = 0;
        let notAvailable = 0;
        if(key === "Technology Item Type:") {
            available = refStats.filter(record => record[key] === item && record["Was Tech available at the time of request?"] === "Yes").length
            notAvailable = refStats.filter(record => record[key] === item && record["Was Tech available at the time of request?"] === "No").length
        }
        //const percentage = totalCount > 0 ? (count / totalCount) * 100 : 0;  // Prevent division by zero
        //(((count-available)/count) * 100).toFixed(1)
        let availablePercentage = ((available/count) * 100).toFixed(1)
        return {
            [item]: [count, availablePercentage, (100-availablePercentage).toFixed(1)]
        };
    });
}

function initializeRovingCountPage() {
    if (document.getElementById("study-room-chart") &&
        document.getElementById("group-table-chart") &&
        document.getElementById("study-carrel-chart") &&
        document.getElementById("computer-lab-chart")) {

        groupRovingData();
        generateTable("study-room-chart", studyRoomAvgHeadCounts);
        generateTable("group-table-chart", groupTablesAvgHeadCounts);
        generateTable("study-carrel-chart", studyCarrelsAvgHeadCounts);
        generateTable("computer-lab-chart", computerLabAvgHeadCounts);
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

function groupRovingData() {
    if (rovingData && rovingData.length > 0) {
       /*rovingData.forEach(item => {
           /*item["Roving Time"] = item["Roving Time"] && item["Roving Time"].trim() !== ""
               ? roundToNearestHalfHour(item["Roving Time"])
               : roundToNearestHalfHour(item["Submitted"]) ; // Assign an empty string if Submitted is undefined or an empty string

           item["Roving Time"] = roundToNearestHalfHour(item["Submitted"])

       });*/
       separateHeadCounts(calculateHeadCountByDay())
   }
}

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
    computerLabAvgHeadCounts = [];
    groupTablesAvgHeadCounts = [];
    studyCarrelsAvgHeadCounts = [];
    studyRoomAvgHeadCounts = [];
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
}

// Function to generate the table
function generateTable(tableName, tableData) {
    let tooltip;
    const table = document.getElementById(tableName);
    if (table) {
        while (table.firstChild) {
          table.removeChild(table.firstChild);
        }
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
}

function splitRecord(record) {
  let splitRecords = [record]; // Start with the original record

  // Iterate over each field in the record
  Object.keys(record).forEach(key => {
    let value = record[key];

    // Check if the value is a string and contains '|'
    if (typeof value === 'string' && value.includes('|')) {
      let values = value.split('|');

      // If there is a split, create new records for each value
      let tempRecords = [];
      splitRecords.forEach((existingRecord, index) => {
        values.forEach((val, i) => {
          let newRecord = { ...existingRecord }; // Clone the existing record
          newRecord[key] = val; // Assign the split value

          // Increment Submission ID by 0.1 for each new record
          if (newRecord.hasOwnProperty('Submission ID')) {
            newRecord['Submission ID'] = parseFloat(existingRecord['Submission ID']) + (i + 1) * 0.1;
          }

          tempRecords.push(newRecord);
        });
      });

      splitRecords = tempRecords; // Replace the old records with the new ones
    }
  });

  return splitRecords;
}

function processRecords(records) {
    return records.map(record => {
        // Function to check and update a field
        function checkAndUpdateField(inquiry, fieldName, validList) {
            let fieldValue = record[fieldName];
            if(inquiry === "Gate Count") {
                if(fieldValue === "") record[fieldName] = 0
            }
            else if (record["Type of Inquiry:"] === inquiry && !validList.includes(fieldValue)) {
                record["Additional Information:"] = (fieldValue ? fieldValue + "; " : "") + (record["Additional Information:"] || "");
                record[fieldName] = "Other";
            }
        }

        // Process all three fields
        checkAndUpdateField("Facilitative", "Type of Facilitative Inquiry:", typeOfFacilitativeInquiry);
        checkAndUpdateField("Basic Reference", "Type of Reference:", typeOfReference);
        checkAndUpdateField("Complex Reference", "Type of Reference:", typeOfReference);
        checkAndUpdateField("Digital Support", "Type of  Digital Support Inquiry:", typeOfDigitalSupportInquiry);
        checkAndUpdateField("Gate Count", "Gate Count:", []);
        checkAndUpdateField("Gate Count", "Computer Lab", []);
        return record;
    });
}

function trimString(item, maxLength) {
    if (item.length > maxLength) {
        return item.slice(0, maxLength) + "...   "; // Truncate to 10 characters
    }
    return item + "      "; // Return the item as is if it's already 10 or fewer characters
}
function shortenListTo10Chars(arr) {
    return arr.map(item => {
        return trimString(item, 15)
    });
}

function loadCharts(chartName, keys) {
    // Enable Chart.js plugin for datalabels
    let mainChart1 = document.getElementById(chartName)
    let ctx = null;

    keys.sort((a, b) => {
          const valueA = Object.values(a)[0][0];
          const valueB = Object.values(b)[0][0];
          return valueB - valueA; // Sort in descending order
        });
    let chartLabels = keys.map(obj => Object.keys(obj)[0])
    let filteredData = keys.map(obj => Object.values(obj)[0][0])
    if(mainChart1) {
        ctx = mainChart1.getContext("2d");

        // Sample data for the courses and percentages
        let data = {
            labels: shortenListTo10Chars(chartLabels),  // Y-axis labels (courses)
            datasets: [{
                label: 'Course Percentage',
                data: filteredData, // X-axis values (percentages)
                backgroundColor: '#006ac3', // Bar color
                borderColor: '#006ac3', // Border color
                borderWidth: 1
            }]
        };

        // Configuration for the chart
        let config = {
            type: 'bar',
            data: data,
            options: {
                responsive: true,
                indexAxis: 'y', // This makes the chart horizontal
                scales: {
                    x: {
                        beginAtZero: true, // Ensures the X-axis starts at 0

                        grid: {
                            display: false // This removes the y-axis grid lines
                        }
                    },
                    y: {
                        beginAtZero: true,
                        ticks: {
                            // Set the course labels to the Y-axis
                            font: {
                                size: 14
                            }
                        }
                    }
                },
                plugins: {
                     legend: {
                        display: false // Hides the legend
                    },
                    tooltip: {

                        callbacks: {
                            label: function(tooltipItem) {
                                // Ensure the raw value is a number and format it to 1 decimal place
                                const value = tooltipItem.raw;
                                const formattedValue = value !== null && !isNaN(value) ? value : value;
                                 if(tooltipItem.dataIndex < keys.length)
                                        return trimString(chartLabels[tooltipItem.dataIndex], 40) + ': ' + formattedValue;
                                 else
                                        return 'Other: ' + formattedValue + '%';
                            },
                            title: function() {
                               return ''; // Clear the title part of the tooltip
                            }
                        } // Remove the title part (optional)

                    },
                    datalabels: {
                        color: (context) => {
                            const value = context.dataset.data[context.dataIndex];
                            return value > 50 ? '#fff' : '#000'; // White font for values > 50%, black for others
                        },
                        font: {
                            weight: "bold",
                            size: 12
                        },
                        formatter: (value) => {

                            return `${value}`; // Display percentage
                        },
                        // Positioning logic for labels based on value
                        anchor: (context) => {
                            const value = context.dataset.data[context.dataIndex];
                            return value > 50 ? 'center' : 'end'; // 'center' for inside, 'end' for outside
                        },
                        align: (context) => {
                            const value = context.dataset.data[context.dataIndex];
                            return value > 50 ? 'center' : 'start'; // 'center' for inside, 'start' for outside
                        },
                        // Adjust label position based on the value (use offset for outside)
                        offset: (context) => {
                            const value = context.dataset.data[context.dataIndex];
                            return value > 50 ? 0 : -40; // No offset inside, offset 10px outside
                        },
                        // Use position to force the label inside or outside based on the value
                        position: (context) => {
                            const value = context.dataset.data[context.dataIndex];
                            return value > 50 ? 'inside' : 'outside'; // 'inside' for > 50%, 'outside' for <= 50%
                        }
                    }

                }
            }
        };

        new Chart(ctx, config);

    } //else loadCharts()
}

function availabilityChart (chartName, keys) {
    var ctx = document.getElementById(chartName).getContext('2d');
    keys.sort((a, b) => {
              const valueA = Object.values(a)[0][0];
              const valueB = Object.values(b)[0][0];
              return valueB - valueA; // Sort in descending order
            });
        let chartLabels = keys.map(obj => Object.keys(obj)[0])
        let filteredData = keys.map(obj => Object.values(obj)[0])
        let availableTech = keys.map(obj => Object.values(obj)[0][1])
        let notAvailableTech = keys.map(obj => Object.values(obj)[0][2])
    var data = {
        labels: chartLabels, // Categories
        datasets: [

            {
                label: 'Available',
                data: availableTech, // Yes percentages for each category
                backgroundColor: '#006ac3', // Color for Yes
                borderColor: '#006ac3',
                borderWidth: 1
            },{
                label: 'Not Available',
                data: notAvailableTech, // No percentages for each category
                backgroundColor: '#505050', // Color for No
                borderColor: '#505050',
                borderWidth: 1
            }
        ]
    };

    var options = {
        indexAxis: 'y',
        responsive: true,
        plugins: {
            legend: {
                display: false // Hides the legend
            },
            tooltip: {
                //position: 'nearest',
                xAlign: 'center', // Align horizontally in the center
                yAlign: 'top', // Align horizontally in the center
                callbacks: {
                    label: function(tooltipItem) {
                        var value = tooltipItem.raw;

                        return tooltipItem.dataset.label + ': ' + value + '%';
                    }
                }
            },
            datalabels: {
                color: '#fff',
                font: {
                    weight: "bold",
                    size: 12
                },
                formatter: (value) => {
                    if (value > 0) {
                        return `${value}%`; // Display percentage if value is greater than 0
                      }
                      return '';
                }
            }
        },
        scales: {
            x: {
                stacked: true, // Stack bars vertically
                beginAtZero: true,
                grid: {
                    display: false // This removes the y-axis grid lines
                }
                //max: 100, // Since it's percentage, the maximum is 100
            },
            y: {
                stacked: true, // Stack bars vertically
            }
        }
    };

    var myChart = new Chart(ctx, {
        type: 'bar',
        data: data,
        options: options
    });

}

// Function to trigger the hidden file input when the "Upload" button is clicked
function triggerFileInput() {
    document.getElementById("file-input").click();
}

// Function to handle the file upload
function handleFileUpload() {
    const fileInput = document.getElementById("file-input");
    const file = fileInput.files[0];

    const fileNameDisplay = document.getElementById("file-name");

    if (file) {
        if (file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || file.type === "application/vnd.ms-excel" || file.type === "text/csv") {
            fileNameDisplay.textContent = `Selected file: ${file.name}`;
        } else {
            fileNameDisplay.textContent = "Please upload a valid Excel file.";
        }
    } else {
        fileNameDisplay.textContent = "No file selected. Please choose a file.";
    }
}

// Function to handle the submit action
function submitFile() {
    const fileInput = document.getElementById("file-input");
    const file = fileInput.files[0];
    const popupContainer = document.getElementById("popup-container");
    const sideTab = document.getElementById("side-tab")
    if (file) {
        // Hide the popup and show the background
        popupContainer.style.display = "none";
        sideTab.style.visibility = "visible";
        loadFile(file);
    } else {
        alert("No file selected. Please upload a file before submitting.");
    }
}

// Event listener for file selection (optional)
document.getElementById("file-input").addEventListener('change', handleFileUpload);

let mainData = [];
let computerLabHeadCount = [];
let groupTablesHeadCount = [];
let studyCarrelsHeadCount = [];
let studyRoomHeadCount = [];
let computerLabAvgHeadCounts = [];
let groupTablesAvgHeadCounts = [];
let studyCarrelsAvgHeadCounts = [];
let studyRoomAvgHeadCounts = [];

let refStatsHeaders = ['Submission ID', 'Submitted', 'Method of Inquiry:', 'Type of Inquiry:', 'Type of Reference:', 'Type of Facilitative Inquiry:',
                'Type of  Digital Support Inquiry:', 'Technology Item Type:', 'Software/Application Type:', "Student's Program", 'Year of Program',
                'Was Tech available at the time of request?', 'Subject(s) of Inquiry:', 'Additional Information:']

let gateCountHeaders = [
    "Submission ID",
    "Submitted", "Gate Count:", "Gate Count - Daily Total",  "Gate Count - Unique Head Count",
    "Computer Lab",  "Computer Lab - Daily Total",  "Computer Lab - Unique Head Count",
     "Subject(s) of Inquiry:", "Additional Information:"]

let methodOfInquiry = [
    "Chat", "In Person", "Phone", "Email", "Form Submission"
]

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

let programs = [
    "Academic Foundations", "Child Care", "Business",
    "College Preparation",
    "Educational Assistant Certificate (EA)", "English For Academic Purposes", "Environmental Technology",
    "General Arts", "General Science", "Governance & Civil Studies", "Health Care Aide", "LINC",
    "Open Studies", "Paramedic", "Practical Nurse", "Power Engineering", "Social Work", "Trades", "UT"
]

let distinctiveInquiryData = [];
let adding = false;
let editing = false;
let tableHeaders;
let tableData;
let gateCountData = [];
let typeOfInquiryData;
let rovingData;
let loanableTechData;
let savedData = [];
let refStats = [];
let addedGateCount = 0;
let addedComputerLab = 0;
let totalGateCount = 0;
let totalComputerLab = 0;
let totalGateCountAverage = 0;
let totalLabAverage = 0;
let lastYear = 0;
let totalDays = 1
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

let activetab;





