function updateUI() {

    const millSelect = document.getElementById("millSelect"); 
    const millChosen = millSelect.value !== ""; //is .value not empty? If so, then true

    // enables all of them
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

function switchTab(mode) {
    // mode is either "manual" or "paste"
    const isManual = mode === "manual";

    // Highlight whichever tab is active
    document.getElementById("tabManual").classList.toggle("active", isManual);
    document.getElementById("tabPaste").classList.toggle("active", !isManual);

    // Show the paste panel only when on the paste tab
    document.getElementById("pastePanel").classList.toggle("visible", !isManual);

    // Hide the "Add row" button when pasting (you don't add rows by hand then)
    document.getElementById("addBtn").style.display = isManual ? "" : "none";
}