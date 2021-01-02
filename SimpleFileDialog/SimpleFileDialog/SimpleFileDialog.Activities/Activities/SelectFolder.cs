using System;
using System.Activities;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using SimpleFileDialog.Activities.Properties;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace SimpleFileDialog.Activities
{
    [LocalizedDisplayName(nameof(Resources.SelectFolder_DisplayName))]
    [LocalizedDescription(nameof(Resources.SelectFolder_Description))]
    public class SelectFolder : ContinuableAsyncCodeActivity
    {
        #region Properties

        /// <summary>
        /// If set, continue executing the remaining activities even if the current activity has failed.
        /// </summary>
        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.ContinueOnError_DisplayName))]
        [LocalizedDescription(nameof(Resources.ContinueOnError_Description))]
        public override InArgument<bool> ContinueOnError { get; set; }

        [LocalizedCategory(nameof(Resources.Common_Category))]
        [LocalizedDisplayName(nameof(Resources.Timeout_DisplayName))]
        [LocalizedDescription(nameof(Resources.Timeout_Description))]
        public InArgument<int> TimeoutMS { get; set; }

        [LocalizedDisplayName(nameof(Resources.SelectFolder_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.SelectFolder_Title_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.SelectFolder_InitialFolderPath_DisplayName))]
        [LocalizedDescription(nameof(Resources.SelectFolder_InitialFolderPath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> InitialFolderPath { get; set; }

        [LocalizedDisplayName(nameof(Resources.SelectFolder_SelectedFolderPath_DisplayName))]
        [LocalizedDescription(nameof(Resources.SelectFolder_SelectedFolderPath_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> SelectedFolderPath { get; set; }

        #endregion


        #region Constructors

        public SelectFolder()
        {
        }

        #endregion


        #region Protected Methods

        protected override void CacheMetadata(CodeActivityMetadata metadata)
        {

            base.CacheMetadata(metadata);
        }

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken cancellationToken)
        {
            // Inputs
            var timeout = TimeoutMS.Get(context);

            // Set a timeout on the execution
            var task = ExecuteWithTimeout(context, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            // Outputs
            return (ctx) => {
                SelectedFolderPath.Set(ctx, task.Result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, CancellationToken cancellationToken = default)
        {
            var title = Title.Get(context);
            var defaultpath = InitialFolderPath.Get(context);

            if (!string.IsNullOrEmpty(defaultpath))
            {
                var index = defaultpath.LastIndexOf(@"\");
                var dotindex = defaultpath.IndexOf(@".", index);
                if (dotindex < 0) defaultpath = defaultpath + @"\";
            }

            var dlg = new CommonOpenFileDialog();
            dlg.IsFolderPicker = true;
            if (!string.IsNullOrEmpty(title)) dlg.Title = title;
            if (!string.IsNullOrEmpty(defaultpath)) dlg.InitialDirectory = defaultpath;

            var result = string.Empty;
            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                result = dlg.FileName;
            }
            return await Task.FromResult(result);
        }

        #endregion
    }
}

