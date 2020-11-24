// set up the Notes app
var notesApp = Application('Notes');
notesApp.includeStandardAdditions = true;

// set up the current app
var currentApp = Application.currentApplication();
currentApp.includeStandardAdditions = true;

// ask for notes account
var notesAccount = currentApp.chooseFromList(['iCloud', 'On My Mac'], {
	withPrompt: "Choose Notes Account",
	defaultItems: ['iCloud'],
});
if (notesAccount.length <= 0) displayError("Notes Account not chosen");
var allNotesInAccount = notesApp.accounts.byName(notesAccount).notes;
if (allNotesInAccount.length <= 0) displayError("Notes Account not found");

// ask for notes from the chosen account
var selectedNotes = currentApp.chooseFromList(allNotesInAccount.name(), {
	withPrompt: "Select Notes",
	multipleSelectionsAllowed: true,
});
if (selectedNotes.length <= 0) displayError("No note selected");

// ask for notes from the chosen account
var outputFormat = currentApp.chooseFromList(['Text file (.txt)', 'Hypertext file (.html)'], {
	withPrompt: "Choose output format",
	defaultItems: ['Text file (.txt)'],
}).toString();
if (outputFormat.length <= 0) displayError("No output format");

if (outputFormat === "Text file (.txt)") {
	outputFileFormat = 'plaintext';
	outputFileSuffix = '.txt';
} else if (outputFormat === "Hypertext file (.html)") {
	outputFileFormat = 'body';
	outputFileSuffix = '.html';
}

// ask for folder to save notes in
var savePath = currentApp.chooseFolder({
	withPrompt: "Choose output location",
}).toString();
if (savePath.length <= 0) displayError("No output location specified");

// set up progress
Progress.totalUnitCount = selectedNotes.length;
Progress.completedUnitCount = 0;
Progress.description = "Exporting Notes...";
Progress.additionalDescription = `Exporting notes into ${outputFormat}s.`;

// loop through all notes in the chosen account
for (var i = 0; i < allNotesInAccount.length; i++) {

	// check against selected notes
	if (selectedNotes.includes(allNotesInAccount[i].name())) {
		// display progress
		Progress.additionalDescription = `Exporting Note ${Progress.completedUnitCount + 1} of ${Progress.totalUnitCount}`;

		// name file path
		var noteFilePath = `${savePath}/${allNotesInAccount[i].name()}${outputFileSuffix}`;

		// write note to file
		writeTextToFile(noteFilePath, allNotesInAccount[i][outputFileFormat](), false);

		// increment progress
		Progress.completedUnitCount++;
	}
}

function displayError(errorMessage) {
	currentApp.displayDialog(errorMessage)
}

function writeTextToFile(file, text, overwriteExistingContent=true) {
	try {

		// Convert the file to a string
		var fileString = file.toString();

		// Open the file for writing
		var openedFile = currentApp.openForAccess(Path(fileString), {writePermission: true});

		// Clear the file if content should be overwritten
		if (overwriteExistingContent) {
			currentApp.setEof(openedFile, {to: 0})
		}

		// Write the new content to the file
		currentApp.write(text, {to: openedFile, startingAt: currentApp.getEof(openedFile)});

		// Close the file
		currentApp.closeAccess(openedFile);

		// Return a boolean indicating that writing was successful
		return true;
	} catch (error) {

		try {
			// Close the file
			currentApp.closeAccess(file);
		} catch (error) {
			// Report the error is closing failed
			console.log(`Couldn't close file: ${error}`);
		}

		// Return a boolean indicating that writing was successful
		return false;
	}
}
