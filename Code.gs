/**
 * Loop through docs in target folder and see what's new
 */
function getFileFromTargetFolder() {

    // Get files
    const folderId = "INSERT_FOLDER_ID_HERE";
    const folder = DriveApp.getFolderById(folderId);
    const filesInFolder = folder.getFiles();

    // get data from sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resumeSheet = ss.getSheetByName("Resumes");
    const resumeSheetData = resumeSheet.getDataRange().getValues();
    const resumeSheetHeaders = resumeSheetData[0];
    const currentFileIds = resumeSheetData.map(x => x[resumeSheetHeaders.indexOf("File id")]).slice(1);
    let rowIndex = currentFileIds.length + 1;

    while (filesInFolder.hasNext()) {
        let resume = filesInFolder.next();
        // Check if resume already in database
        if (currentFileIds.indexOf(resume.getId()) == - 1) {
            // If not, extract data and add to the sheet
            let resumeText = extractTextFromPDF(resume.getId(), folderId);
            let resumeExtractedInfo = extractResumeDataFromText(resumeText);
            console.log(resumeExtractedInfo);

            let newRow = [resume.getUrl(),
            resumeExtractedInfo["Name"],
            resumeExtractedInfo["Address"],
            resumeExtractedInfo["Email"],
            resumeExtractedInfo["Phone number"],
            resumeExtractedInfo["Skills"],
            resumeExtractedInfo["Experiences"],
            resumeExtractedInfo["Education"],
            resume.getId(),
            ];
            // Add new row to sheet
            resumeSheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
            rowIndex++;
        }
    }

}

/**
 * EXTRACT TEXT CONTENT FROM PDF 
 * 
 * @param {string} fileId
 * @param {string} parentFolderId
 * @returns {string} pdfContent
 */
function extractTextFromPDF(fileId, parentFolderId) {

    const destFolder = Drive.Files.get(parentFolderId, { "supportsAllDrives": true });
    const newFile = {
        "fileId": fileId,
        "parents": [
            destFolder
        ]
    };
    const args = {
        "resource": {
            "parents": [
                destFolder
            ],
            "name": "temp",
            "mimeType": "application/vnd.google-apps.document",
        },
        "supportsAllDrives": true
    };

    const newTargetDoc = Drive.Files.copy(newFile, fileId, args);
    const newTargetFile = DocumentApp.openById(newTargetDoc.getId());
    const pdfContent = newTargetFile.getBody().getText();
    Drive.Files.remove(newTargetDoc.getId());

    return pdfContent;
}

/**
 * From the extracted pdf content, retrieve the relevant information:
 * File URL	Name	Address	Email	Phone number	Skills	Employers	Education	File id
 */
function extractResumeDataFromText(pdfContent) {

    // Extracting names
    const contentAsArray = pdfContent.split(" ");
    const firstName = contentAsArray.filter(x => FIRST_NAMES.includes(x));
    const lastName = contentAsArray.filter(x => LAST_NAMES.includes(x));
    let name;
    if (firstName.length > 0 && lastName.length > 0) {
        name = firstName[0] + " " + lastName[0];
    } else {
        name = "Unable to parse";
    }

    // Extracting Phone Number
    const phoneNumberRegex = /\b\d{3}-\d{3}-\d{4}\b/g;
    const phoneNumber = (pdfContent.match(phoneNumberRegex)) ? pdfContent.match(phoneNumberRegex)[0] : "Unable to parse";

    // Extracting Email Address
    const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const email = (pdfContent.match(emailRegex)) ? pdfContent.match(emailRegex)[0] : "Unable to parse";

    // Extracting Address
    const addressRegex = /\b\d+\s[A-Za-z\s]+\s[A-Za-z\s]+,\s[A-Za-z\s]+\b/g;
    const address = (pdfContent.match(addressRegex)) ? pdfContent.match(addressRegex)[0] : "Unable to parse";

    // Extracting Skills
    const skillsRegex = /S K I L L S((.|\n)*)E D U C A T I O N/g;
    const skills = (pdfContent.match(skillsRegex)) ? pdfContent.match(skillsRegex)[0]
        .replace(/S K I L L S|E D U C A T I O N|\n/g, '')
        .trim()
        .split('\n')[0] : "Unable to parse";

    // Extracting Education
    const educationRegex = /E D U C A T I O N((.|\n)*)L A N G U A G E/g;
    const education = (pdfContent.match(educationRegex)) ? pdfContent.match(educationRegex)[0]
        .replace(/E D U C A T I O N|L A N G U A G E|\n/g, '')
        .trim()
        .split('\n')[0] : "Unable to parse";

    // Extracting experiences
    const experienceRegex = /E X P E R I E N C E((.|\n)*?)(?=W E B\sC O N T E N T\sM A N A G E R|$)/g;
    const experienceMatch = pdfContent.match(experienceRegex);
    const experienceContent = experienceMatch ? experienceMatch[0] : "Unable to parse";

    const resumeMap = {
        "Name": name,
        "Phone number": phoneNumber,
        "Email": email,
        "Address": address,
        "Skills": skills,
        "Education": education,
        "Experiences": experienceContent
    }

    return resumeMap;
}