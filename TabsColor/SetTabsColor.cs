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
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private DTE2 _dte;
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

            Instance._dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
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
            DarwTabsColor().Forget();
        }

        private bool _isWaitingForTabsCommand;
        private bool _modify = false;
        private bool _untrack = false;
        private List<string> _modifyNames = new List<string>();
        private List<string> _untrackNames = new List<string>();
        private async Task DarwTabsColor()
        {
            if (_isWaitingForTabsCommand)
                return;

            _isWaitingForTabsCommand = true;

            //获取git修改的文件名
            //设置tabs的着色选项为 正则表达式
            //判断ColorByRegexConfig文件是否存在
            //如果不存在则创建一个默认的
            try
            {
                _modify = false;
                _untrack = false;
                _modifyNames.Clear();
                _untrackNames.Clear();
                Shell.ShellE("git diff --name-only", ModifiesOutputData);
                await WaitModifyAsync();
                Shell.ShellE("git ls-files --others --exclude-standard", UntrackOutputData);
                await WaitUntrackAsync();
                SaveColorByRegexConfig();
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

        #region 写入配置文件

        private const string ColorByRegexConfigDefaultContent = @"// 此文件包含按正则表达式对文档选项卡进行着色的规则。每行都包含一个正则表达式，该表达式将针对文件的完整路径进行测试。与正则表达式匹配的所有文件都将共享一种颜色。
// 可以通过右键单击选项卡并选择“设置制表符颜色”来自定义分配给任何文件组的颜色。
// 正则表达式将按照它们在此文件中的显示顺序进行匹配。有关语法，请参阅 https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions。
// 正则表达式匹配为不区分大小写。可以使用捕获组选项(如""(?-i:expression)"")重写此行为。

// 编辑此文件并保存更改以立即查看应用的更改。分析或计算表达式期间遇到的任何错误都将出现在名为 按正则表达式显示颜色 的窗格的输出窗口中。
^.*\.cs$
^.*\.fs$
^.*\.vb$
^.*\.cp?p?$
^.*\.hp?p?$
^.*\.txt$
^.*\.xml$";

        private void SaveColorByRegexConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                string solutionName = _dte.Solution.FullName;
                string fileName = Path.GetFileNameWithoutExtension(solutionName);
                string folderPath = Path.GetDirectoryName(solutionName);
                string cofigPath = $"{folderPath}\\.vs\\{fileName}\\ColorByRegexConfig.txt";
                string newRegex = GetCurRegexByGit();
                if (File.Exists(cofigPath))
                {
                    var content = File.ReadAllText(cofigPath);
                    //查找是否有标记符 有则替换
                    Regex pattern = new Regex("//TabsColorTag\n(.*?)\n//TabsColorTag\n");
                    Match match = pattern.Match(content);
                    if (match.Success)
                    {
                        string extractedContent = match.Value;
                        if (extractedContent == newRegex)
                            return;
                        
                        string newContent = content.Replace(extractedContent, newRegex);
                        Console.WriteLine("Extracted content: " + extractedContent);
                        File.WriteAllText(cofigPath, newContent);
                    }
                    else
                    {
                        //没有则直接添加
                        string newContent = newRegex + content;
                        File.WriteAllText(cofigPath, newContent);
                    }
                }
                else
                {
                    //等用户手动创建 不自动创建
                    return;
                    string newContent = newRegex + ColorByRegexConfigDefaultContent;
                    File.WriteAllText(cofigPath, newContent);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("????????????????");
            }
        }

        /// <summary>
        /// 通过git获取当前的正则表达式
        /// </summary>
        /// <returns></returns>
        private string GetCurRegexByGit()
        {
            string regex = "//TabsColorTag\n^.*(?:#)\\.cs$\n//TabsColorTag\n";
            StringBuilder sb = new StringBuilder();
            foreach (var item in _modifyNames)
            {
                sb.Append(item);
                sb.Append("|");
            }
            foreach (var item in _untrackNames)
            {
                sb.Append(item);
                sb.Append("|");
            }
            if (sb.Length > 0)
            {
                sb.Remove(sb.Length - 1, 1);
            }
            return regex.Replace("#", sb.ToString());

        }

        #endregion

        #region shell命令
        private async Task<bool> WaitModifyAsync()
        {
            while (!_modify)
            {
                await Task.Delay(100);
            }
            return true;
        }

        private async Task<bool> WaitUntrackAsync()
        {
            while (!_untrack)
            {
                await Task.Delay(100);
            }
            return true;
        }

        private static void ModifiesOutputData(object sender, DataReceivedEventArgs e)
        {
            if (Instance._modify)
                return;

            if (e.Data == null)
            {
                Instance._modify = true;
                return;
            }
            Instance._modifyNames.Add(GetFileName(e.Data));
        }

        private static void UntrackOutputData(object sender, DataReceivedEventArgs e)
        {
            if (Instance._untrack)
                return;
            if (e.Data == null)
            {
                Instance._untrack = true;
                return;
            }

            Instance._untrackNames.Add(GetFileName(e.Data));
        }

        private static string GetFileName(string path)
        {
            Regex pattern = new Regex(@"([^/]+)\.cs$");
            Match match = pattern.Match(path);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return path;
        }


        public event DataReceivedEventHandler OutPutInfo;

        public class Shell
        {
            public static System.Diagnostics.Process ShellE(string shell, DataReceivedEventHandler onOutPutInfo)
            {
                var p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "cmd.exe";
                p.StartInfo.Arguments = @"/c " + shell;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardError = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardInput = false;
                p.StartInfo.UseShellExecute = false;
                p.OutputDataReceived += OutputData;
                p.OutputDataReceived += onOutPutInfo;
                p.ErrorDataReceived += ErrorData;
                p.Start();
                p.BeginErrorReadLine();
                p.BeginOutputReadLine();
                p.WaitForExit();
                return p;
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
        #endregion

        /// <summary>
        /// 监听保存事件
        /// </summary>
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
