using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace ToggleComment
{
    /// <summary>
    /// 拡張機能として配置されるパッケージです。
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "2.1", IconResourceID = 400)] // Visual Studio のヘルプ/バージョン情報に表示される情報です。
    [Guid(PackageGuidString)]
    [ProvideOptionPage(typeof(OptionPageGrid),
    "Toggle Comment 2022", "Command Options", 0, 0, true)]
    [ProvideService(typeof(ToggleCommentPackage), IsAsyncQueryable = true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class ToggleCommentPackage : AsyncPackage
    {
        /// <summary>
        /// パッケージのIDです。
        /// </summary>
        public const string PackageGuidString = "9fb26121-a4a7-4f27-81c3-c713a2464345";

        /// <summary>
        /// パッケージを初期化します。
        /// </summary>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            ToggleSingleComment.Initialize(this);
            ToggleDoubleComment.Initialize(this);
        }
    }
}