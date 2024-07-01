# UI Tests for Edi

UI automation tests for Edi.

## Built With

- Visual Studio 2022
- .NET 8.0
- NUnit as a test framework
- FlaUI for UI testing
- ExtentReport for automatic report generation

## Running the tests

In Visual Studio go to Test>Test Explorer, then click Run All Tests.

The tests will search for Edit in `$(SolutionDir)$(Configuration)` which is where the `Edi.exe` should be if Edi is built correctly. If you want to run the tests without building Edi from source, then change `APP_PATH` constant in `EdiUITestSession.cs`.

## Tests

### OpenFileWithArgumentsTest
Writes lorem ipsum to the file 'OpenFileWithArgumentsTest.txt'.
Opens the app with the filepath as an argument.
Selects the tab of the file.
Checks that the text editors value is the same as the file.

### OpenFileFromEditorTest
Writes lorem ipsum to the file 'OpenFileFromEditorTest.txt'.
Opens the app, then opens the file with File>Open.
Selects the tab of the file.
Checks that the text editors value is the same as the file.

### OpenAndModifyTest
Writes "Old text" to file 'OpenAndModifyTest.txt'.
Opens the app with the filepath as an argument.
Selects the tab of the file.
Selects all of the text Edit>Select All.
Types "New text" to the text editor.
Waits for 3 seconds so the typing has time to finish.
Checks that the value of the text editor is "New text".
Saves the file with File>Save.
Checks that the saved file is "New text".

### CopyPasteTest
Writes lorem ipsum to the file 'CopyPasteTest.txt'.
Opens the app with the filepath as an argument.
Selects the tab of the file.
Selects all of the text Edit>Select All.
Copies the text with Edit>Copy.
Pastes the cut text 2x with Edit>Paste.
Checks that the value of the text editor is lorem ipsum repeated 2x.

### UndoRedoTest
Writes lorem ipsum to the file 'UndoRedoTest.txt'.
Opens the app with the filepath as an argument.
Selects the tab of the file.
Selects all of the text Edit>Select All.
Cuts the text with Edit>Cut.
Pastes the cut text 3x with Edit>Paste.
Undos 3x with Edit>Undo.
Redos 2x with Edit>Redo.
Undos 1x with Edit>Undo.
Checks that the value of the text editor is lorem ipsum.

### SaveNewFileTest
Opens the app then closes all the tabs so the new 'Untitled.txt' tab can be selectecten unambiguously.
Then create a new file with File>New>Text Document.
Selects the tab with the title 'Untitled.txt'
Types a lorem ipsum text.
Waits for 3 seconds so the typing has time to finish.
Checks that the value of the text editor is the typed lorem ipsum text.
Saves the file as 'SaveNewFileTest.txt' with File>Save As.
Fills out the Save As dialog.
Checks if the saved file has the lorem ipsum text.

### ClosedRecentFileTest
Writes lorem ipsum to file 'ClosedRecentFileTest.txt'.
Opens the app.
Opens the file with File>Open.
Closes the app.
Deletes the file.
Opens the app again.
Check if the app launched without an error window showing up.

> Note: The latest version of Edi does not pass this test.

## Issues with Edi

On my computer Visual Studio couldn't resolve all the reference needed for Edi. Mainly `MWindowLib` and `MWindowInterfacesLib`, and maybe releated to this Visual Studio ignored the build order specified. If it happens to you remove these from the references and readd them manualy, then build each project one-by-one according to the build order specified in the solution.

