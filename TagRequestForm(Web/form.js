
// Enable buttons and tabs
function updateUI() {

    const millSelect = document.getElementById("millSelect"); 
    const millChosen = millSelect.value !== ""; //is .value not empty? If so, then true

    // enables all of them when millSelect.value != ""
    // if millChosen = true -> false -> disabled = false -> button is enabled
    document.getElementById("addRow").disabled = !millChosen;
    document.getElementById("clrAll").disabled = !millChosen;
    document.getElementById("chkAll").disabled = !millChosen;
    document.getElementById("unchkAll").disabled = !millChosen;
    document.getElementById("copyTable").disabled = !millChosen;
    document.getElementById("submitReq").disabled = !millChosen;

    document.getElementById("manualOption").disabled = !millChosen;
    document.getElementById("pasteOption").disabled = !millChosen;
    document.getElementById("pasteTextBox").disabled = !millChosen;
    document.getElementById("importButton").disabled = !millChosen;

}

// Make urgent button red
function toggleUrgent(urgentBtn){
    const label = urgentBtn.closest(".urgent-toggle");
    label.classList.toggle("on", urgentBtn.checked);
}

// switches to paste panel
function switchTab(mode) {
    // mode is either "manual" or "paste"
    const isManual = mode === "manual";

    // Highlight whichever tab is active
    document.getElementById("manualOption").classList.toggle("active", isManual);
    document.getElementById("pasteOption").classList.toggle("active", !isManual);

    // Show the paste panel only when on the paste tab
    document.getElementById("pastePanel").classList.toggle("visible", !isManual);

    // Hide the "Add row" button when pasting (you don't add rows by hand then)
    document.getElementById("addRow").style.display = isManual ? "" : "none";
}

// Downloads the excel file
function openExcelTemplate() {
    const link = document.createElement("a");
    link.href = "../TagReqestTemplate.xlsx";
    link.download = "../TagRequestTemplate.xlsx";
    link.click();
}

function processExcelData() {

    // Uses addRow()

}

function addRow(data = {}) {

    const tableBody = document.getElementById("tbody");
    const row = document.createElement("tr");

    row.appendChild(makeInputCell("tableName", "Search or type table name...", data.tableName));
    row.appendChild(makeInputCell("tagName", "e.g. BladeSpeedRpm", data.tagName));
    row.appendChild(makeInputCell("plcAddress", "192.168.x.x", data.plcAddress));
    row.appendChild(makeInputCell("plcPath", "PLC tag path", data.plcPath));
    row.appendChild(makeInputCell("dataType", "— type —", data.dataType));
    row.appendChild(makeInputCell("units", "— units —", data.units));
    row.appendChild(makeInputCell("transactionFrequency", "e.g. 5 s, 1 min", data.transactionFrequency));
    row.appendChild(makeInputCell("tagDescription", "What does this tag measure?", data.tagDescription));
    row.appendChild(makeInputCell("min", "e.g. 0", data.min));
    row.appendChild(makeInputCell("max", "e.g. 100", data.max));

    tableBody.appendChild(row);

}

// Builds a <td> containing a text input, returns the cell
function makeInputCell(className, placeholder, value) {

    const cell = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-input " + className;
    input.placeholder = placeholder || "";
    input.value = value || "";
    cell.appendChild(input);
    return cell;
    
}






