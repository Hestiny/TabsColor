using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace TabsColor
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SetTabsColor
    {
        private RunningDocTableEvents _runningDocTableEvents;
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("32555dae-e482-4829-a6b5-c41ea26b0362");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SetTabsColor"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SetTabsColor(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SetTabsColor Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in SetTabsColor's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SetTabsColor(package, commandService);
            Instance._runningDocTableEvents = new RunningDocTableEvents(package);
            Assumes.Present(Instance._runningDocTableEvents);
            if (!Instance._isInitEvent)
                Instance._runningDocTableEvents.AfterSave += Instance.AfterSaveHanel;
            Instance._isInitEvent = true;
        }

        private bool _isInitEvent;
        private void AfterSaveHanel(object sender, Document document)
        {
            DarwTabsColor().Forget();
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            //TODO: 直接切换tabs为正则表达式
            ThreadHelper.ThrowIfNotOnUIThread();
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "SetTabsColor";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            //IDG_VS_FILE_SAVE
        }

        private bool _isWaitingForTabsCommand;

        private async Task DarwTabsColor()
        {
            if (_isWaitingForTabsCommand)
                return;

            _isWaitingForTabsCommand = true;
            //TODO: 获取git修改的文件名
            //设置tabs的着色选项为 正则表达式
            //判断ColorByRegexConfig文件是否存在
            //如果不存在则创建一个默认的
            try
            {
                Shell.ShellE();
            }
            catch (Exception)
            {
                Console.WriteLine("????????????????");
            }
            finally
            {
                _isWaitingForTabsCommand = false;
            }

        }

        public class Shell
        {
            public static void ShellE()
            {
                var p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.Arguments = @"/c git diff --name-only";
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardInput = false;
                p.StartInfo.UseShellExecute = false;
                p.OutputDataReceived += OutputData;
                p.ErrorDataReceived += ErrorData;
                p.Start();
                p.BeginErrorReadLine();
                p.BeginOutputReadLine();
                p.WaitForExit();
            }

            private static void ErrorData(object sender, DataReceivedEventArgs e)
            {
                Console.WriteLine(e.Data);
            }

            private static void OutputData(object sender, DataReceivedEventArgs e)
            {
                Console.WriteLine(e.Data);
            }
        }

        public class RunningDocTableEvents : IVsRunningDocTableEvents3
        {
            #region Members

            private RunningDocumentTable mRunningDocumentTable;
            private DTE mDte;

            public delegate void OnAfterSaveHandler(object sender, Document document);
            public event OnAfterSaveHandler AfterSave;

            public delegate void OnBeforeSaveHandler(object sender, Document document);
            public event OnBeforeSaveHandler BeforeSave;

            #endregion

            #region Constructor

            public RunningDocTableEvents(Package aPackage)
            {
                mDte = (DTE)Package.GetGlobalService(typeof(DTE));
                mRunningDocumentTable = new RunningDocumentTable(aPackage);
                mRunningDocumentTable.Advise(this);
            }

            #endregion

            #region IVsRunningDocTableEvents3 implementation

            public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
            {
                return VSConstants.S_OK;
            }

            public int OnAfterSave(uint docCookie)
            {
                if (null == AfterSave)
                    return VSConstants.S_OK;

                var document = FindDocumentByCookie(docCookie);
                if (null == document)
                    return VSConstants.S_OK;

                AfterSave(this, FindDocumentByCookie(docCookie));
                return VSConstants.S_OK;
            }

            public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
            {
                return VSConstants.S_OK;
            }

            public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
            {
                return VSConstants.S_OK;
            }

            public int OnBeforeSave(uint docCookie)
            {
                if (null == BeforeSave)
                    return VSConstants.S_OK;

                var document = FindDocumentByCookie(docCookie);
                if (null == document)
                    return VSConstants.S_OK;

                BeforeSave(this, FindDocumentByCookie(docCookie));
                return VSConstants.S_OK;
            }

            #endregion

            #region Private Methods

            private Document FindDocumentByCookie(uint docCookie)
            {
                var documentInfo = mRunningDocumentTable.GetDocumentInfo(docCookie);
                return mDte.Documents.Cast<Document>().FirstOrDefault(doc => doc.FullName == documentInfo.Moniker);
            }

            #endregion
        }

    }
}
