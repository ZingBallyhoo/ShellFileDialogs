using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ShellFileDialogs.Native;

namespace ShellFileDialogs.Dialogs
{
    public static class FileSaveDialog
    {
	    /// <summary>
	    ///     Shows the file-save dialog. Returns <see langword="null" /> if the user canceled the dialog - otherwise
	    ///     returns the user-selected path.
	    /// </summary>
	    /// <param name="selectedFilterZeroBasedIndex">
	    ///     0-based index of the filter to select. Use <c>-1</c> to indicate no default
	    ///     selection.
	    /// </param>
        public static string? ShowDialog(IntPtr parentHWnd, string? title, string? initialDirectory,
            string? defaultFileName, IReadOnlyCollection<Filter>? filters, int selectedFilterZeroBasedIndex = -1)
        {
            var nfod = new NativeFileSaveDialog();
            try
            {
                return ShowDialogInner(nfod, parentHWnd, title, initialDirectory, defaultFileName, filters);
            }
            finally
            {
                _ = Marshal.ReleaseComObject(nfod);
            }
        }
        
        private static string? ShowDialogInner(IFileSaveDialog dialog, IntPtr parentHWnd, string? title,
            string? initialDirectory, string? defaultFileName, IReadOnlyCollection<Filter>? filters,
            int selectedFilterZeroBasedIndex = -1)
        {
            const FileOpenOptions flags = FileOpenOptions.NoTestFileCreate |
                                          FileOpenOptions.PathMustExist |
                                          FileOpenOptions.ForceFilesystem |
                                          FileOpenOptions.OverwritePrompt;

            dialog.SetOptions(flags);

            if (title != null) dialog.SetTitle(title);

            if (initialDirectory != null)
            {
                var initialDirectoryShellItem = Utility.ParseShellItem2Name(initialDirectory);
                if (initialDirectoryShellItem != null) dialog.SetFolder(initialDirectoryShellItem);
            }

            if (defaultFileName != null) dialog.SetFileName(defaultFileName);

            Utility.SetFilters(dialog, filters, selectedFilterZeroBasedIndex);

            //

            var hr = dialog.Show(parentHWnd);
            if (hr.ValidateDialogShowHResult())
            {
                dialog.GetResult(out var selectedItem);
                if (selectedItem == null) return null;

                return Utility.GetFileNameFromShellItem(selectedItem);
            }

            // User cancelled.
            return null;
        }
    }
}