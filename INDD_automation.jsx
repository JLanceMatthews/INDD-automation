/* InDesign Script to Automate Offering Memorandum Creation */

// Define the property data file (CSV format)
var dataFile = File.openDialog("Select the CSV file with property details");
if (!dataFile) {
    alert("No file selected. Script terminated.");
    exit();
}

// Read the CSV file and parse data
var fileContent = "";
dataFile.open("r");
while (!dataFile.eof) {
    fileContent += dataFile.readln() + "\n";
}
dataFile.close();
var rows = fileContent.split("\n");
var headers = rows[0].split(",");

// Create a function to extract data
function getData(propertyID, columnName) {
    var colIndex = headers.indexOf(columnName);
    if (colIndex === -1) return "";
    for (var i = 1; i < rows.length; i++) {
        var columns = rows[i].split(",");
        if (columns[0] === propertyID) {
            return columns[colIndex];
        }
    }
    return "";
}

// Get active document
var doc = app.activeDocument;

// Fetch property ID from the document
var propertyID = doc.textFrames.itemByName("property_id").contents;

// Auto-fill property details
var propertyName = getData(propertyID, "Property Name");
doc.textFrames.itemByName("property_name").contents = propertyName;
doc.textFrames.itemByName("property_address").contents = getData(propertyID, "Address");
doc.textFrames.itemByName("property_price").contents = getData(propertyID, "Price");
doc.textFrames.itemByName("cap_rate").contents = getData(propertyID, "Cap Rate");
doc.textFrames.itemByName("noi").contents = getData(propertyID, "NOI");

doc.textFrames.itemByName("broker_name").contents = getData(propertyID, "Broker Name");
doc.textFrames.itemByName("broker_phone").contents = getData(propertyID, "Broker Phone");
doc.textFrames.itemByName("broker_email").contents = getData(propertyID, "Broker Email");

// Auto-insert property image
var imagePath = getData(propertyID, "Image Path");
if (imagePath) {
    var imageFrame = doc.rectangles.itemByName("property_image");
    imageFrame.place(File(imagePath));
}

// Apply table styles for financials
var financialTable = doc.stories[0].tables[0];
financialTable.appliedTableStyle = doc.tableStyles.itemByName("FinancialTable");
financialTable.cells.everyItem().appliedCellStyle = doc.cellStyles.itemByName("FinancialCells");

// Export the OM as a PDF with structured naming
var pdfFile = new File("~/Documents/OMs/" + propertyName.replace(/\s+/g, "_") + "_OM.pdf");
doc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, app.pdfExportPresets.itemByName("High Quality Print"));

alert("Offering Memorandum Generated Successfully!");
