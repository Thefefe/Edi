using AventStack.ExtentReports.MarkupUtils;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using NUnit.Framework.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

namespace EdiFlaUITest
{
    public class EdiUITestSession
    {
        protected const string LOREM_IPSUM = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";

        // TODO: add project to Edi solution, make path relative
        //protected const string APP_PATH = @"D:\\dev\\projects\\Edi\\Release\\Edi.exe";
        private const string APP_PATH = @"..\..\..\..\Release\Edi.exe";

        private static TimeSpan DEFAULT_TIMEOUT = TimeSpan.FromSeconds(3);

        private static UIA3Automation automation;
        private static FlaUI.Core.Application? app;
        protected static FlaUI.Core.AutomationElements.Window? window;

        private static ExtentReports extent;
        protected static ExtentTest testReport;

        protected static string outputDir;
        protected static string currentTestFileName;
        protected static string currentTestFilePath;

        /// <summary>
        /// Opens the app with the optional arguments.
        /// </summary>
        /// <param name="arguments">Arguments to be passed to the app</param>
        protected static void OpenApp(string? arguments = null)
        {
            Assert.That(app, Is.Null);
            app = FlaUI.Core.Application.Launch(APP_PATH, arguments);
            window = app.GetMainWindow(automation);
        }

        /// <summary>
        /// Closes the app.
        /// </summary>
        protected static void CloseApp()
        {
            Assert.That(app, Is.Not.Null);

            if (app.Close())
            {
                testReport.Log(Status.Pass, "Window closed gracefully.");
            }
            else
            {
                testReport.Log(Status.Warning, "Window couldn't be closed gracefully.");
            }

            app.Dispose();

            app = null;
            window = null;
        }

        /// <summary>
        /// Create a file with the given name at `outputDir`, then writes the provided text to it.
        /// If a file with the same name exis
        /// </summary>
        /// <param name="file">The name of the file</param>
        /// <param name="text">The text to be writen to the file</param>
        protected static void WriteToTextFile(string file, string text)
        {
            File.WriteAllText(outputDir + file, text);
        }

        /// <summary>
        /// Reads a text file from `outputDir`.
        /// </summary>
        /// <param name="file">The name of the file</param>
        /// <returns>The contents of the file</returns>
        protected static string ReadFromTextFile(string file)
        {
            return Retry.WhileException(
                () => File.ReadAllText(outputDir + file),
                TimeSpan.FromSeconds(5)
            ).Result;
        }

        /// <summary>
        /// Save the file in the editor with File>Save As, with the given name int the `outputDir` directory.
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        protected static void SaveFileAs(string fileName)
        {
            ClickMenu("File", "Save As ...");

            var saveDialog = MustGet("Save Dialog Window", () => window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Window).And(cf.ByName("Save As"))));

            var saveFileDirToolbar = MustGet(
                "Save Dialog Directory Toolbar",
                () => saveDialog.FindFirstByXPath("/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar")
            );
            saveFileDirToolbar.Click();

            var saveFileDirEdit = MustGet(
                "Save Dialog Directory TextBox",
                () => saveDialog.FindFirstByXPath("/Pane[2]/Pane[3]/ProgressBar/ComboBox/Edit").AsTextBox()
            );
            saveFileDirEdit.Text = outputDir;
            Keyboard.Press(VirtualKeyShort.ENTER);

            var saveFileNameEdit = MustGet(
                "Save Dialog Filename TextBox",
                () => saveDialog.FindFirstByXPath("/Pane[1]/ComboBox[1]/Edit").AsTextBox()
            );
            saveFileNameEdit.Text = fileName;

            var saveDialogSave = MustGet(
                "Save Dialog Save Button",
                () => saveDialog.FindFirstDescendant(cf => cf.ByControlType(ControlType.Button).And(cf.ByName("Save")))
            );
            saveDialogSave.Click();

            var confirmDialog = TryGet(
                "Save As Override Confirm Dialog",
                () => saveDialog.FindFirstDescendant(cf => cf.ByControlType(ControlType.Window).And(cf.ByName("Save As")))
            );
            if (confirmDialog is not null)
            {
                var confirmOverride = MustGet(
                    "Confirm Override Dialog Confirm Button",
                    () => confirmDialog.FindFirstDescendant(cf => cf.ByControlType(ControlType.Button).And(cf.ByName("Yes")))
                );
                confirmOverride.Click();
            }
        }

        /// <summary>
        /// Opens a file in the editor with File>Open from the `outputDir`.
        /// </summary>
        /// <param name="file">The file to bne </param>
        protected static void OpenFile(string file)
        {
            ClickMenu("File", "Open", "Text files");

            var openDialog = MustGet(
                "Open Dialog Window",
                () => window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Window).And(cf.ByName("Open")))
            );

            var dirToolbar = MustGet(
                "Open Dialog Directory Toolbar",
                () => openDialog.FindFirstByXPath("/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar")
            );
            dirToolbar.Click();

            var dirEdit = MustGet(
                "Open Dialog Directory TextBox",
                () => openDialog.FindFirstByXPath("/Pane[2]/Pane[3]/ProgressBar/ComboBox/Edit").AsTextBox()
            );
            dirEdit.Text = outputDir;
            Keyboard.Press(VirtualKeyShort.ENTER);

            var fileNameEdit = MustGet(
                "Open Dialog Filename TextBox",
                () => openDialog.FindFirstByXPath("/ComboBox[1]/Edit").AsTextBox()
            );
            fileNameEdit.Text = file;

            var openButton = MustGet(
                "Open Dialog Open Button",
                () => openDialog.FindFirstDescendant(
                    cf => cf.ByControlType(ControlType.Button)
                        .And(cf.ByName("Open"))
                        .And(cf.ByAutomationId("1")))
            );
            openButton.Click();
        }

        /// <summary>
        /// Ignores missing recent file errors.
        /// Logs them as warnings.
        /// </summary>
        protected static void IgnoreRecentFileErrors()
        {
            var filePath = RecentFileError();
            while (filePath is not null)
            {
                testReport.Warning($"A recent file couldn't be opened: '{filePath}'");
                filePath = RecentFileError();
            }
        }

        /// <summary>
        /// Handles the window that pops up when a file in the recent files that auto load has been deleted.
        /// </summary>
        /// <returns>The name of the path to the file or null if no such error popped up</returns>
        protected static string? RecentFileError()
        {
            var popupWindow = TryGet("Popup window", () => window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Window)));
            if (popupWindow is null) return null;

            var titleText = TryGet("Error text", () => popupWindow.FindFirstDescendant(cf => cf.ByName("ERROR LOADING FILE")), TimeSpan.FromSeconds(5));
            if (titleText is null) return null;

            var errorText = MustGet(
                "Error Window Text",
                () => popupWindow.FindFirstByXPath("/Custom/Edit")
            ).AsTextBox();

            var text = errorText.Text;
            Assert.That(text.Contains("does not exist or cannot be loaded."));

            var filePath = text.Split("'")[1];

            var button = MustGet(
                "Error Window Confirm Button",
                () => popupWindow.FindFirstDescendant(cf => cf.ByName("Yes"))
            );
            button.Click();

            return filePath;
        }

        /// <summary>
        /// Clicks the given menu path.
        /// </summary>
        /// <param name="main">The main menu in the menu bar</param>
        /// <param name="submenus">The subsequent menus to click</param>
        protected static void ClickMenu(string main, params string[] submenus)
        {
            window.Focus();
            var menuBar = MustGet("Menu Bar", () => window.FindFirstDescendant(cf => cf.Menu()).AsMenu());
            var menuPath = main;
            MenuItem current = menuBar.Items[main];
            Assert.That(current, Is.Not.Null, $"Couldn't find menu item '{main}' from menu bar");
            current.Click();

            foreach (var menuName in submenus)
            {
                menuPath += $">{menuName}";
                current = current.Items[menuName];
                Assert.That(current, Is.Not.Null, $"Couldn't find menu item '{menuPath}'");
                current.Click();
            }

            testReport.Pass($"Clicked menu {menuPath}");
        }

        /// <summary>
        /// Gets the Tab item with the given name.
        /// </summary>
        /// <param name="tabName">The name of the tab</param>
        /// <returns>The tab</returns>
        /// <exception cref="UnreachableException"></exception>
        protected static TabItem GetTabItem(string tabName)
        {
            window.Focus();
            var dockView = MustGet("Dock View", () => window.FindFirstDescendant(cf => cf.ByClassName("AvalonDockView")));
            var tabControls = MustGet("Tab Control", () => dockView.FindAllDescendants(cf => cf.ByClassName("TabControl")));

            foreach (var tabControl in tabControls)
            {
                foreach (var tabItem in tabControl.AsTab().TabItems)
                {
                    if (tabItem.HelpText == tabName)
                    {
                        return tabItem;
                    }
                }
            }

            Assert.Fail($"Couldn't find tab with name '{tabName}'");
            throw new UnreachableException();
        }

        /// <summary>
        /// Closes all open tabs.
        /// </summary>
        protected static void CloseAllTabs()
        {
            window.Focus();
            var dockView = MustGet("Dock View", () => window.FindFirstDescendant(cf => cf.ByClassName("AvalonDockView")));
            var tabControls = MustGet("Tab Control", () => dockView.FindAllDescendants(cf => cf.ByClassName("TabControl")));

            foreach (var tabControl in tabControls)
            {
                foreach (var tabItem in tabControl.AsTab().TabItems)
                {
                    Mouse.Position = tabItem.GetClickablePoint();
                    Mouse.Click(MouseButton.Middle);
                    Wait.UntilInputIsProcessed();
                }
            }
        }

        /// <summary>
        /// Gets the text editor component from the TabItem
        /// </summary>
        /// <param name="tabItem">The TabItem to extract the text edit component from</param>
        /// <returns></returns>
        protected static AutomationElement GetEditFromTabItem(TabItem tabItem)
        {
            window.Focus();
            var ediView = MustGet("Edi View", () => tabItem.FindFirstDescendant(cf => cf.ByClassName("EdiView")));
            return MustGet("Tab Edit", () => ediView.FindFirstDescendant(cf => cf.ByLocalizedControlType("custom")));
        }

        /// <summary>
        /// Calls the given function until it's not null, or the timeout has elapsed.
        /// If the timeout has been reached, throws an exception.
        /// </summary>
        /// <param name="name">The name to log to the report</param>
        /// <param name="func">The function to be called</param>
        /// <param name="timeout">The timeout (Optional)</param>
        /// <returns>The result of the function</returns>
        protected static T MustGet<T>(string name, Func<T> func, TimeSpan? timeout = null) where T : class
        {
            var result = Retry.WhileNull(func, timeout ?? DEFAULT_TIMEOUT).Result;
            Assert.That(result, Is.Not.Null, $"Couldn't find '{name}'");
            testReport.Pass($"Found '{name}'");
            return result;
        }

        /// <summary>
        /// Calls the given function until it's not null, or the timeout has elapsed.
        /// If the timeout has been reached, returns null.
        /// </summary>
        /// <param name="name">The name to log to the report</param>
        /// <param name="func">The function to be called</param>
        /// <param name="timeout">The timeout. (Optional)</param>
        /// <returns>The result of the function, or null if the timeout has been reached</returns>
        protected static T? TryGet<T>(string name, Func<T> func, TimeSpan? timeout = null) where T : class?
        {
            var e = Retry.WhileNull(func, timeout ?? DEFAULT_TIMEOUT).Result;
            if (e is not null)
            {
                testReport.Info($"Found optional item '{name}'");
            }
            else
            {
                testReport.Info($"Couldn't find optional item '{name}'");
            }
            return e;
        }

        [OneTimeSetUp]
        public static void GlobalSetup()
        {
            var r = Directory.GetParent(@"..\..\..\..\")?.FullName ?? ".";
            outputDir = r + @"\UITests\TestOutput\";
            Directory.CreateDirectory(outputDir);

            extent = new ExtentReports();

            var spark = new ExtentSparkReporter(outputDir + "report.html");
            extent.AttachReporter(spark);

            spark.Config.Theme = AventStack.ExtentReports.Reporter.Config.Theme.Dark;
            spark.Config.DocumentTitle = "Edi UI Test Report";
        }

        [SetUp]
        public static void TestSetup()
        {
            var testName = TestContext.CurrentContext.Test.Name;

            currentTestFileName = $"{testName}.txt";
            currentTestFilePath = outputDir + currentTestFileName;

            automation = new UIA3Automation();

            testReport = extent.CreateTest(testName);
        }

        [TearDown]
        public static void TestTeardown()
        {
            switch (TestContext.CurrentContext.Result.Outcome.Status)
            {
                case TestStatus.Failed:
                    {
                        testReport.Fail(MarkupHelper.CreateCodeBlock(TestContext.CurrentContext.Result.Message));
                        break;
                    };
                case TestStatus.Inconclusive:
                case TestStatus.Warning:
                    {
                        testReport.Warning("Warning");
                        break;
                    };
                case TestStatus.Skipped:
                    {
                        testReport.Skip("Skip");
                        break;
                    }
                default:
                    {
                        testReport.Pass("Pass");
                        break;
                    };
            }

            if (app is not null)
            {
                CloseApp();
            }

            automation.Dispose();
        }

        [OneTimeTearDown]
        public static void GlobalTeardown()
        {
            extent.Flush();
        }
    }
}
