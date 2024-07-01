using AventStack.ExtentReports;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.MarkupUtils;
using NUnit.Framework.Interfaces;
using System.Windows.Forms;
using AventStack.ExtentReports.Gherkin.Model;
using Microsoft.CodeAnalysis;
using System.Collections.Generic;

namespace EdiFlaUITest
{
    public class Tests : EdiUITestSession
    {

        /// <summary>
        /// Writes lorem ipsum to the file 'OpenFileWithArgumentsTest.txt'.
        /// Opens the app with the filepath as an argument.
        /// Selects the tab of the file.
        /// Checks that the text editors value is the same as the file.
        /// </summary>
        [Test, Order(0)]
        public void OpenFileWithArgumentsTest()
        {
            WriteToTextFile(currentTestFileName, LOREM_IPSUM);

            OpenApp(currentTestFilePath);
            IgnoreRecentFileErrors();

            var tabItem = GetTabItem(currentTestFilePath);
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            var textValue = textEdit?.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(LOREM_IPSUM));
        }

        /// <summary>
        /// Writes lorem ipsum to the file 'OpenFileFromEditorTest.txt'.
        /// Opens the app, then opens the file with File>Open.
        /// Selects the tab of the file.
        /// Checks that the text editors value is the same as the file.
        /// </summary>
        [Test, Order(1)]
        public void OpenFileFromEditorTest()
        {
            WriteToTextFile(currentTestFileName, LOREM_IPSUM);

            OpenApp();
            IgnoreRecentFileErrors();
            OpenFile(currentTestFileName);

            var tabItem = GetTabItem(currentTestFilePath);
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            var textValue = textEdit?.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(LOREM_IPSUM));
        }

        /// <summary>
        /// Writes "Old text" to file 'OpenAndModifyTest.txt'.
        /// Opens the app with the filepath as an argument.
        /// Selects the tab of the file.
        /// Selects all of the text Edit>Select All.
        /// Types "New text" to the text editor.
        /// Waits for 3 seconds so the typing has time to finish.
        /// Checks that the value of the text editor is "New text".
        /// Saves the file with File>Save.
        /// Checks that the saved file is "New text".
        /// </summary>
        [Test, Order(2)]
        public void OpenAndModifyTest()
        {
            const string OLD_TEXT = "Old text";
            const string NEW_TEXT = "New text";

            WriteToTextFile(currentTestFileName, OLD_TEXT);

            OpenApp(currentTestFilePath);
            IgnoreRecentFileErrors();
            var tabItem = GetTabItem(currentTestFilePath);
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            ClickMenu("Edit", "Select All");

            Keyboard.Type(NEW_TEXT);
            Thread.Sleep(3000); // `Wait.UntilInputIsProcessed()` only waits 100ms, which isn't enough

            var textValue = textEdit?.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(NEW_TEXT));

            ClickMenu("File", "Save");

            var fileContent = ReadFromTextFile(currentTestFileName);
            Assert.That(fileContent, Is.EqualTo(NEW_TEXT));
        }

        /// <summary>
        /// Writes lorem ipsum to the file 'CopyPasteTest.txt'.
        /// Opens the app with the filepath as an argument.
        /// Selects the tab of the file.
        /// Selects all of the text Edit>Select All.
        /// Copies the text with Edit>Copy.
        /// Pastes the cut text 2x with Edit>Paste.
        /// Checks that the value of the text editor is lorem ipsum repeated 2x.
        /// </summary>
        [Test, Order(3)]
        public void CopyPasteTest()
        {
            WriteToTextFile(currentTestFileName, LOREM_IPSUM);

            OpenApp(currentTestFilePath);
            IgnoreRecentFileErrors();
            var tabItem = GetTabItem(currentTestFilePath);
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            ClickMenu("Edit", "Select All");
            ClickMenu("Edit", "Copy");
            ClickMenu("Edit", "Paste");
            ClickMenu("Edit", "Paste");

            var textValue = textEdit?.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(LOREM_IPSUM + LOREM_IPSUM));

            // Just so the unsaved file popup window doesn't appear on exit
            ClickMenu("File", "Save");
        }

        /// <summary>
        /// Writes lorem ipsum to the file 'UndoRedoTest.txt'.
        /// Opens the app with the filepath as an argument.
        /// Selects the tab of the file.
        /// Selects all of the text Edit>Select All.
        /// Cuts the text with Edit>Cut.
        /// Pastes the cut text 3x with Edit>Paste.
        /// Undos 3x with Edit>Undo.
        /// Redos 2x with Edit>Redo.
        /// Undos 1x with Edit>Undo.
        /// Checks that the value of the text editor is lorem ipsum.
        /// </summary>
        [Test, Order(4)]
        public void UndoRedoTest()
        {
            WriteToTextFile(currentTestFileName, LOREM_IPSUM);

            OpenApp(currentTestFilePath);
            IgnoreRecentFileErrors();
            var tabItem = GetTabItem(currentTestFilePath);
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            ClickMenu("Edit", "Select All");
            ClickMenu("Edit", "Cut");

            ClickMenu("Edit", "Paste");
            ClickMenu("Edit", "Paste");
            ClickMenu("Edit", "Paste");

            ClickMenu("Edit", "Undo");
            ClickMenu("Edit", "Undo");
            ClickMenu("Edit", "Undo");

            ClickMenu("Edit", "Redo");
            ClickMenu("Edit", "Redo");

            ClickMenu("Edit", "Undo");

            var textValue = textEdit?.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(LOREM_IPSUM));

            // Just so the unsaved file popup window doesn't appear on exit
            ClickMenu("File", "Save");
        }

        /// <summary>
        /// Steps:
        ///  Opens the app then closes all the tabs so the new 'Untitled.txt' tab can be selectecten unambiguously.
        ///  Then create a new file with File>New>Text Document.
        ///  Selects the tab with the title 'Untitled.txt'
        ///  Types a lorem ipsum text.
        ///  Waits for 3 seconds so the typing has time to finish.
        ///  Checks that the value of the text editor is the typed lorem ipsum text.
        ///  Saves the file as 'SaveNewFileTest.txt' with File>Save As.
        ///  Fills out the Save As dialog.
        ///  Checks if the saved file has the lorem ipsum text.
        /// </summary>
        [Test, Order(5)]
        public void SaveNewFileTest()
        {
            OpenApp();
            IgnoreRecentFileErrors();
            CloseAllTabs();

            ClickMenu("File", "New", "Text Document");

            var tabItem = GetTabItem("Untitled.txt");
            tabItem.Click();

            var textEdit = GetEditFromTabItem(tabItem);
            textEdit.Click();

            Keyboard.Type(LOREM_IPSUM);
            Thread.Sleep(3000); // `Wait.UntilInputIsProcessed()` only waits 100ms, which isn't enough

            var textValue = textEdit.Patterns.Value.PatternOrDefault.Value.Value;
            Assert.That(textValue, Is.EqualTo(LOREM_IPSUM));

            var tabLabel = MustGet("Tab Item Label", () => tabItem.FindFirstDescendant(cf => cf.ByClassName("TextBlock")));
            Assert.That(tabLabel.Name.EndsWith('*'), "Tab label doesn't have the dirty marker '*'");
            testReport.Pass("Tab label has the dirty marker '*'");

            SaveFileAs(currentTestFileName);

            var fileContent = ReadFromTextFile(currentTestFileName);
            Assert.That(fileContent, Is.EqualTo(LOREM_IPSUM));
        }

        /// <summary>
        /// Writes lorem ipsum to file 'ClosedRecentFileTest.txt'.
        /// Opens the app.
        /// Opens the file with File>Open.
        /// Closes the app.
        /// Deletes the file.
        /// Opens the app again.
        /// Check if the app launched without an error window showing up.
        /// 
        /// Note: The latest version of Edi does not pass this test.
        /// </summary>
        [Test, Order(6)]
        public void ClosedRecentFileTest()
        {
            WriteToTextFile(currentTestFileName, LOREM_IPSUM);

            OpenApp();
            IgnoreRecentFileErrors(); // if this happens here it's a side effect external to this test
            OpenFile(currentTestFileName);
            CloseApp();

            File.Delete(currentTestFilePath);

            OpenApp();
            var recentFileErrorPath = RecentFileError();
            if (recentFileErrorPath is not null)
            {
                Assert.Fail($"Edi showed error message about a recent file not being available. File path: '{recentFileErrorPath}'");
            }
            else
            {
                testReport.Pass("Edi did not show error message about recent file(s) not being available.");
            }
        }
    }
}