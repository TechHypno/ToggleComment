using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ToggleComment.Codes;
using ToggleComment.Utils;

namespace ToggleComment
{
    /// <summary>
    /// 拡張機能として登録するコマンドの基底クラスです。
    /// </summary>
    internal abstract class CommandBase
    {
        /// <summary>
        /// コマンドを提供するパッケージです。
        /// </summary>
        private readonly Package _package;

        /// <summary>
        /// コマンドメニューグループのIDです。
        /// </summary>
        public static readonly Guid CommandSet = new Guid("85542055-97d7-4219-a793-8c077b81b25b");

        /// <summary>
        /// サービスプロバイダーを取得します。
        /// </summary>
        protected System.IServiceProvider ServiceProvider
        {
            get { return _package; }
        }

        /// <summary>
        /// コマンドの実行を委譲するインスタンスです。
        /// </summary>
        private readonly IOleCommandTarget _commandTarget;

        /// <summary>
        /// コメントのパターンです。
        /// </summary>
        private readonly IDictionary<string, ICodeCommentPattern[]> _patterns = new Dictionary<string, ICodeCommentPattern[]>();

        private readonly OptionPageGrid _config;
        protected OptionPageGrid Config
        {
            get { return _config; }
        }

        /// <summary>
        /// インスタンスを初期化します。
        /// </summary>
        /// <remarks>
        /// コマンドは .vsct ファイルに定義されている必要があります。
        /// </remarks>
        /// <param name="package">コマンドを提供するパッケージ</param>
        /// <param name="commandId">コマンドのID</param>
        /// <param name="commandSetId">コマンドメニューグループのID</param>
        protected CommandBase(Package package, int commandId, Guid commandSetId)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            _package = package;
            _config = package.GetDialogPage(typeof(OptionPageGrid)) as OptionPageGrid;

            var commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(commandSetId, commandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }

            _commandTarget = (IOleCommandTarget)ServiceProvider.GetService(typeof(SUIHostCommandDispatcher));
        }

        protected abstract ICodeCommentPattern[] CreateCommentPatterns(string language);
        protected abstract void OnExecute(string language, ICodeCommentPattern[] patterns, TextSelection selection);
        /// <summary>
        /// コマンドを実行します。
        /// </summary>c />
        private void Execute(object sender, EventArgs e)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
            if (dte?.ActiveDocument.Object("TextDocument") is TextDocument textDocument)
            {
                var patterns = _patterns.GetOrAdd(textDocument.Language, CreateCommentPatterns);
#if DEBUG
                System.Diagnostics.Debug.WriteLine($"Language: {textDocument.Language}");
#endif
                if (0 < patterns.Length)
                {
                    var selection = textDocument.Selection;
                    OnExecute(textDocument.Language, patterns, selection);
                }
                else if (ExecuteCommand(VSConstants.VSStd2KCmdID.COMMENT_BLOCK) == false)
                {
                    ShowMessageBox(
                        "Toggle Comment is not executable.",
                        $"{textDocument.Language} files is not supported.",
                        OLEMSGICON.OLEMSGICON_INFO);
                }
            }
        }

        /// <summary>
        /// メッセージボックスを表示します。
        /// </summary>
        /// <param name="title">メッセージのタイトル</param>
        /// <param name="message">表示するメッセージ</param>
        /// <param name="icon">表示するアイコン</param>
        protected void ShowMessageBox(string title, string message, OLEMSGICON icon)
        {
            VsShellUtilities.ShowMessageBox(
                ServiceProvider,
                message,
                title,
                icon,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// コマンドを実行した際のコールバックです。
        /// </summary>
        /// <param name="sender">イベントの発行者</param>
        /// <param name="e">イベント引数</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                Execute(sender, e);
            }
            catch (Exception ex)
            {
                ShowMessageBox(
                    $"{GetType().Name} is not executable.",
                    $"{ex.GetType().FullName}: {ex.Message}.",
                    OLEMSGICON.OLEMSGICON_WARNING);
            }
        }

        /// <summary>
        /// 指定のコマンドを実行します。
        /// コマンドが実行できなかった場合は<see langword="false"/>を返します。
        /// </summary>
        protected bool ExecuteCommand(VSConstants.VSStd2KCmdID commandId)
        {
            var grooupId = VSConstants.VSStd2K;
            var result = _commandTarget.Exec(ref grooupId, (uint)commandId, 0, IntPtr.Zero, IntPtr.Zero);

            return result == VSConstants.S_OK;
        }

        /// <summary>
        /// 選択中の行を行選択状態にします。
        /// </summary>
        protected static void SelectLines(TextSelection selection)
        {
            var startPoint = selection.TopPoint.CreateEditPoint();
            startPoint.StartOfLine();

            var endPoint = selection.BottomPoint.CreateEditPoint();
            if (endPoint.AtStartOfLine == false || startPoint.Line == endPoint.Line)
            {
                endPoint.EndOfLine();
            }

            if (selection.Mode == vsSelectionMode.vsSelectionModeBox)
            {
                selection.Mode = vsSelectionMode.vsSelectionModeStream;
            }

            selection.MoveToPoint(startPoint);
            selection.MoveToPoint(endPoint, true);
        }
    }
}
