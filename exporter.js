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
if (notesAccount.length <= 0) throw new Error("Notes Account not chosen");
var allNotesInAccount = notesApp.accounts.byName(notesAccount).notes;
if (allNotesInAccount.length <= 0) throw new Error("Notes Account not found");

// ask for notes from the chosen account
var selectedNotes = currentApp.chooseFromList(allNotesInAccount.name(), {
	withPrompt: "Select Notes",
	multipleSelectionsAllowed: true,
});
if (selectedNotes.length <= 0) throw new Error("No note selected");

// ask for notes from the chosen account
var outputFormat = currentApp.chooseFromList(['Text file (.txt)', 'Hypertext file (.html)'], {
	withPrompt: "Choose output format",
	defaultItems: ['Text file (.txt)'],
}).toString();
if (outputFormat.length <= 0) throw new Error("No output format");

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

// loop through all notes in the chosen account
for (var i = 0; i < allNotesInAccount.length; i++) {

	// check against selected notes
	if (selectedNotes.includes(allNotesInAccount[i].name())) {

		// name file path
		var noteFilePath = `${savePath}/${allNotesInAccount[i].name()}${outputFileSuffix}`;
		// open file
		var openedNoteFile = currentApp.openForAccess(Path(noteFilePath), {writePermission: true});
		if (!openedNoteFile) continue;
		// write note to file
		currentApp.setEof(openedNoteFile, {to: 0});
		currentApp.write(allNotesInAccount[i][outputFileFormat](), {to: openedNoteFile});
		currentApp.closeAccess(openedNoteFile);
	}
}
