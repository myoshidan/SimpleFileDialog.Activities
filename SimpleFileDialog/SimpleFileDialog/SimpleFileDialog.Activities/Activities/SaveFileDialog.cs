using System;
using System.Activities;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.WindowsAPICodePack.Dialogs;
using SimpleFileDialog.Activities.Properties;
using SimpleFileDialog.Enums;
using UiPath.Shared.Activities;
using UiPath.Shared.Activities.Localization;

namespace SimpleFileDialog.Activities
{
    [LocalizedDisplayName(nameof(Resources.SaveFileDialog_DisplayName))]
    [LocalizedDescription(nameof(Resources.SaveFileDialog_Description))]
    public class SaveFileDialog : ContinuableAsyncCodeActivity
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
        public InArgument<int> TimeoutMS { get; set; } = 60000;

        [LocalizedDisplayName(nameof(Resources.SaveFileDialog_Title_DisplayName))]
        [LocalizedDescription(nameof(Resources.SaveFileDialog_Title_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> Title { get; set; }

        [LocalizedDisplayName(nameof(Resources.SaveFileDialog_InitialPath_DisplayName))]
        [LocalizedDescription(nameof(Resources.SaveFileDialog_InitialPath_Description))]
        [LocalizedCategory(nameof(Resources.Input_Category))]
        public InArgument<string> InitialPath { get; set; }

        [LocalizedDisplayName(nameof(Resources.SaveFileDialog_FileType_DisplayName))]
        [LocalizedDescription(nameof(Resources.SaveFileDialog_FileType_Description))]
        [LocalizedCategory(nameof(Resources.Options_Category))]
        public FileTypes FileType { get; set; } = 0;

        [LocalizedDisplayName(nameof(Resources.SaveFileDialog_SelectedFilePath_DisplayName))]
        [LocalizedDescription(nameof(Resources.SaveFileDialog_SelectedFilePath_Description))]
        [LocalizedCategory(nameof(Resources.Output_Category))]
        public OutArgument<string> SelectedFilePath { get; set; }

        #endregion


        #region Constructors

        public SaveFileDialog()
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
            var title = Title.Get(context);
            var initialpath = InitialPath.Get(context);

            // Set a timeout on the execution
            var task = ExecuteWithTimeout(context, cancellationToken);
            if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) throw new TimeoutException(Resources.Timeout_Error);

            // Outputs
            return (ctx) => {
                SelectedFilePath.Set(ctx, task.Result);
            };
        }

        private async Task<string> ExecuteWithTimeout(AsyncCodeActivityContext context, CancellationToken cancellationToken = default)
        {
            var title = Title.Get(context);
            var defaultpath = InitialPath.Get(context);
            var result = string.Empty;
            var dlg = new CommonSaveFileDialog();


            if (!string.IsNullOrEmpty(title)) dlg.Title = title;
            if (!string.IsNullOrEmpty(Path.GetDirectoryName(defaultpath))) dlg.InitialDirectory = Path.GetDirectoryName(defaultpath);
            if (!string.IsNullOrEmpty(Path.GetFileName(defaultpath))) dlg.DefaultFileName = Path.GetFileName(defaultpath);

            if ((int)this.FileType >= 1)
            {
                var filter = new CommonFileDialogFilter();
                filter.DisplayName = "許可されたファイル";

                if (this.FileType.HasFlag(FileTypes.Excel))
                {
                    filter.Extensions.Add("xls");
                    filter.Extensions.Add("xlsx");
                    filter.Extensions.Add("xlsm");
                }
                if (this.FileType.HasFlag(FileTypes.CSV))
                {
                    filter.Extensions.Add("csv");
                }
                if (this.FileType.HasFlag(FileTypes.PDF))
                {
                    filter.Extensions.Add("pdf");
                }
                if (this.FileType.HasFlag(FileTypes.Text))
                {
                    filter.Extensions.Add("txt");
                }
                if (this.FileType.HasFlag(FileTypes.PowerPoint))
                {
                    filter.Extensions.Add("ppt");
                    filter.Extensions.Add("pptx");
                }
                if (this.FileType.HasFlag(FileTypes.Word))
                {
                    filter.Extensions.Add("doc");
                    filter.Extensions.Add("docx");
                }
                dlg.Filters.Add(filter);
            }
            else
            {
                dlg.Filters.Add(new CommonFileDialogFilter("すべてのファイル", "*.*"));
            }

            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                result = dlg.FileName;

            }

            return await Task.FromResult(result);
        }

        #endregion
    }
}

