
/*===================================================================
    UI Updates
  =================================================================== */

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
    if(urgentBtn.checked){
      urgentBtn.parentElement.classList.add("on");
    }
    else {
      urgentBtn.parentElement.classList.remove("on");
    }
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
    link.href = "../TagRequestTemplate.xlsx";
    link.download = "../TagRequestTemplate.xlsx";
    link.click();
}




/*===================================================================
    ADD ROW RUNCTIONS (addrow(), )
  =================================================================== */
function addRow(data = {}) {  //data {} comes from processExcelData

    // Adding data inputs Tag Name -> Max
    const tableBody = document.getElementById("tbody");
    const row = document.createElement("tr");
    row.appendChild(makeInputBoxes("tableName", "Search or type table name...", data.tableName));
    row.appendChild(makeInputBoxes("tagName", "e.g. BladeSpeedRpm", data.tagName));
    row.appendChild(makeInputBoxes("plcAddress", "192.168.x.x", data.plcAddress));
    row.appendChild(makeInputBoxes("plcPath", "PLC tag path", data.plcPath));
    row.appendChild(makeInputBoxes("dataType", "— type —", data.dataType));
    row.appendChild(makeInputBoxes("units", "— units —", data.units));
    row.appendChild(makeInputBoxes("transactionFrequency", "e.g. 5 s, 1 min", data.transactionFrequency));
    row.appendChild(makeInputBoxes("tagDescription", "What does this tag measure?", data.tagDescription));
    row.appendChild(makeInputBoxes("min", "e.g. 0", data.min));
    row.appendChild(makeInputBoxes("max", "e.g. 100", data.max));

    // "Add new table?" box
    const checkBoxCell = document.createElement("td");
    checkBoxCell.className = "checkbox-cell";
    const checkBox = document.createElement("input");
    checkBox.type = "checkbox";
    checkBox.className = "newTable";
    if (data.newTable){
        checkBox.checked = true;
    }
    checkBoxCell.appendChild(checkBox);
    row.appendChild(checkBoxCell);

    // Remove button
    const removeCell = document.createElement("td");
    const removeBtn = document.createElement("button");
    removeBtn.className = "remove-btn";
    removeBtn.textContent = "x";
    removeBtn.onclick = function () {
        row.remove();
    };
    removeCell.appendChild(removeBtn);
    row.appendChild(removeCell);

    // Adding all the table data points <td> to the tableBody
    tableBody.appendChild(row);

    // Gets # of rows and changes UI based on # of rows
    rowsCount();

    // .tageName is the class name for <input>

    //  ======VALIDATION FUNCTIONS=======
    const tagInput = row.querySelector(".tagName");
    tagInput.addEventListener("blur", function() {
        valTag(tagInput);
        chkDups();
    })

    const ipInput = row.querySelector(".plcAddress");
    ipInput.addEventListener("blur", function() {
      valIP(ipInput);
    })
    
    const minInput = row.querySelector(".min");
    minInput.addEventListener("blur", function() {
      valNum(minInput);
    })

    const maxInput = row.querySelector(".max");
    maxInput.addEventListener("blur", function() {
      valNum(maxInput);
    })


}

// Builds <input> and puts it into <td>
function makeInputBoxes(className, placeholder, value) {

    const cell = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.className = "cell-input " + className;
    input.placeholder = placeholder || "";
    input.value = value || "";
    cell.appendChild(input);
    return cell;

}

// Gets called at the end of each addRow() execution
function rowsCount(){

  const rowCount = document.querySelectorAll('#tbody tr').length;
  const noTagsMessage = document.getElementById("emptyState");
  noTagsMessage.style.display = rowCount === 0 ? "block" : "none";

  document.getElementById("rowCounter").textContent =
  rowCount + (rowCount === 1 ? " row" : " rows");

}



/*===================================================================
    INPUT VALIDATIONS (CHECKS WHEN TO ADD WARNING NOTES)
  =================================================================== */


// Adds warning text under data inputs
function addWarning(input, text){
  clearMsg(input);   
                           // clear any old one first
  const msg = document.createElement("div");

  msg.className = "msg warn";
  msg.textContent = "avoid underscores in this tag";

  input.parentElement.appendChild(msg);       // add under the input
  input.classList.add("warn");                // tint input amber
}

// Checks for valid tagName
function valTag(input) {
  clearMsg(input);

  const value = input.value.trim();

  if (!value) return; // if there is no value, we don't need to do anything

  if (value.includes("_")){
    addWarning(input, "Avoid underscores in the tag name")
  }
}

// Checks for valid IP adress
function valIP(input) {

  const IP_PATTERN = /^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?::\d{1,5})?$/;

  clearMsg(input);

  const value = input.value.trim();
  
  if (!value) {
    return;
  }

  if (value.includes(",")) {
    addWarning(input, "Avoid adding , in PLC IP Adresses. Use . instead");
    return;
  }
  
  if (!IP_PATTERN.test(value)){
    addWarning(input, "Avoid adding , in PLC IP Adresses. Use . instead");
    return;
  }
}

// Checks for valid num
function valNumber(input) {
  clrMsg(input);

  const value = input.value.trim();
  if(value.includes(",")) {
    addMsg(input, "Use . instead of ,");
    return;
  }

  if (!/^-?\d*\.?\d*$/.test(value)) {        // regex that checks for valid number
        addMsg(input, label + " must be a number");
    }
}

function chkDups() {
  const rows = document.querySelectorAll("#tbody tr"); // says, look at id toby, and see how many tablerows (tr) there are
  const counts = {};
    
  rows.forEach(function (row) { //creating dict. key = table and value    value = # of occurance
    const tableName = (row.querySelector(".tableName")?.value || "").trim().toLowerCase();
    const tagName = (row.querySelector(".tagName")?.value || "").trim().toLowerCase();

    // Create hashmap
    if (tagName) {
      const key = tableName + "|" + tagName;
      counts[key] = (counts[key] || 0) + 1;  //now we have the occurence value, the scope of this is known outside
    }
  }); 
  
  rows.forEach(function (row) {
    const tagInput = row.querySelector(".tagName");
    const tableName = (row.querySelector(".tableName")?.value || "").trim().toLowerCase();
    const tagName = (row.querySelector(".tagName")?.value || "").trim().toLowerCase();
      
    clearMsg(tagInput); 

    const key = tableName + "|" + tagName;

    if(counts[key] > 1){
      addMsg(tagInput, "Duplicate tags are not allowed")
    }
  });
}



/*===================================================================
  SUPPORT FUNCTIONS
  =================================================================== */

//Adds the warning message 
function addMsg(input, text) {
  const msg = document.createElement("div");
  msg.className = "msg warn"
  msg.textContent = text;
  input.parentElement.appendChild(msg)
}

// Clears the old warning msg
function clearMsg(input){
  const oldMsg = input.parentElement.querySelector(".msg");
  if (oldMsg) oldMsg.remove();   // remove the message if one exists
  input.classList.remove("warn"); // remove amber tint from the input
}




/*===================================================================
  Fetch units from dataparc endpoint
  =================================================================== */
let units = [];
async function getUnits(){
  try {

    //response = fetch(ENDPOINT)
    units = await response.json();
  }

  catch (err) {
    console.error(err)
  }
}



function processExcelData() {

    // Uses addRow()
    // will have data = {tableName:
    //                   tagName:
    //                   plcAddress: }
}



/*============================================================================
  ============================================================================*/
