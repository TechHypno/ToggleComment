using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using ToggleComment.Codes;
using System.ComponentModel;

namespace ToggleComment
{
    #region Configuration
    public partial class OptionPageGrid : DialogPage
    {
        private bool optionMoveCaretDown1C = false;
        private bool optionMoveCaretDown1U = false;

        [Category("Toggle Single Comment")]
        [DisplayName("Move caret down when commenting.")]
        [Description("Place the caret below the selection when commenting.")]
        public bool OptionMoveCaretDown1C
        {
            get { return optionMoveCaretDown1C; }
            set { optionMoveCaretDown1C = value; }
        }

        [Category("Toggle Single Comment")]
        [DisplayName("Move caret down when uncommenting.")]
        [Description("Place the caret below the selection when uncommenting.")]
        public bool OptionMoveCaretDown1U
        {
            get { return optionMoveCaretDown1U; }
            set { optionMoveCaretDown1U = value; }
        }
    }
    #endregion

    /// <summary>
    /// 選択された行のコメントアウト・解除を行うコマンドです。
    /// </summary>
    internal sealed class ToggleSingleComment : CommandBase
    {
        /// <summary>
        /// コマンドのIDです。
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// シングルトンのインスタンスを取得します。
        /// </summary>
        public static ToggleSingleComment Instance { get; private set; }

        /// <summary>
        /// インスタンスを初期化します。
        /// </summary>
        /// <param name="package">コマンドを提供するパッケージ</param>
        private ToggleSingleComment(Package package) : base(package, CommandId, CommandSet){}

        /// <summary>
        /// このコマンドのシングルトンのインスタンスを初期化します。
        /// </summary>
        /// <param name="package">コマンドを提供するパッケージ</param>
        public static void Initialize(Package package)
        {
            Instance = new ToggleSingleComment(package);
        }

        protected override void OnExecute(string language, ICodeCommentPattern[] patterns, TextSelection selection)
        {
            SelectLines(selection);
            var text = selection.Text;
            var isComment = patterns.Any(x => x.IsComment(text));
            if (isComment)
            {
                ExecuteCommand(VSConstants.VSStd2KCmdID.UNCOMMENT_BLOCK);
                if (Config.OptionMoveCaretDown1U)
                    ExecuteCommand(VSConstants.VSStd2KCmdID.DOWN);
            }
            else
            {
                ExecuteCommand(VSConstants.VSStd2KCmdID.COMMENT_BLOCK);
                if (Config.OptionMoveCaretDown1C)
                    ExecuteCommand(VSConstants.VSStd2KCmdID.DOWN);
            }
        }

        /// <summary>
        /// コードのコメントを表すパターンを作成します。
        /// </summary>
        protected override ICodeCommentPattern[] CreateCommentPatterns(string language)
        {
            switch (language)
            {
                case "CSharp":
                case "C/C++":
                case "TypeScript":
                case "JSON":
                    {
                        return new ICodeCommentPattern[] { new LineCommentPattern("//"), new BlockCommentPattern("/*", "*/") };
                    }
                case "XML":
                case "XAML":
                    {
                        return new[] { new BlockCommentPattern("<!--", "-->") };
                    }
                case "HTMLX":
                    {
                        // MEMO : HTML に埋め込まれたCSS, JavaScriptをサポートする
                        return new ICodeCommentPattern[] {
                            new BlockCommentPattern("<!--", "-->"),
                            new BlockCommentPattern("@*", "*@"),
                            new BlockCommentPattern("/*", "*/"),
                            new LineCommentPattern("//")};
                    }
                case "HTML":
                    {
                        // MEMO : VS の UncommentSelection コマンドがブロックコメント <%/* */%> に対応していない
                        return new ICodeCommentPattern[] {
                            new BlockCommentPattern("<!--", "-->"),
                            new BlockCommentPattern("<%--", "--%>"),
                            new BlockCommentPattern("/*", "*/"),
                            new LineCommentPattern("//")};
                    }
                case "JavaScript":
                case "F#":
                    {
                        // MEMO : VS の UncommentSelection コマンドが JavaScript, F# のブロックコメントに対応していない
                        return new[] { new LineCommentPattern("//") };
                    }
                case "CSS":
                    {
                        return new[] { new BlockCommentPattern("/*", "*/") };
                    }
                case "PowerShell":
                    {
                        // MEMO : VS の UncommentSelection コマンドが PowerShell のブロックコメントに対応していない
                        return new[] { new LineCommentPattern("#") };
                    }
                case "Lua":
                case "SQL Server Tools":
                    {
                        return new[] { new LineCommentPattern("--") };
                    }
                case "Basic":
                    {
                        return new[] { new LineCommentPattern("'") };
                    }
                case "Python":
                    {
                        return new[] { new LineCommentPattern("#") };
                    }
                default:
                    {
                        return new ICodeCommentPattern[0];
                    }
            }
        }
    }
}