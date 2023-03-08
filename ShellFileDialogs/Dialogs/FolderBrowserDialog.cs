using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ShellFileDialogs.Native;

namespace ShellFileDialogs.Dialogs
{
    public static class FolderBrowserDialog
    {
	    /// <summary>
	    ///     Shows the folder browser dialog. Returns <see langword="null" /> if the user cancelled the dialog. Otherwise
	    ///     returns the selected path.
	    /// </summary>
        public static string? ShowDialog(IntPtr parentHWnd, string? title, string? initialDirectory)
        {
            var nfod = new NativeFileOpenDialog();
            try
            {
                return ShowDialogInner(nfod, parentHWnd, title, initialDirectory);
            }
            finally
            {
                _ = Marshal.ReleaseComObject(nfod);
            }
        }
        
        private static string? ShowDialogInner(IFileOpenDialog dialog, IntPtr parentHWnd, string? title,
            string? initialDirectory)
        {
            //IFileDialog ifd = dialog;
            var flags =
                FileOpenOptions.NoTestFileCreate |
                FileOpenOptions.PathMustExist |
                FileOpenOptions.PickFolders |
                FileOpenOptions.ForceFilesystem;

            dialog.SetOptions(flags);

            if (title != null) dialog.SetTitle(title);

            if (initialDirectory != null)
            {
                var initialDirectoryShellItem = Utility.ParseShellItem2Name(initialDirectory);
                if (initialDirectoryShellItem != null) dialog.SetFolder(initialDirectoryShellItem);
            }

            //

            var hr = dialog.Show(parentHWnd);
            if (hr.ValidateDialogShowHResult())
            {
                dialog.GetResults(out var resultsArray);
                
                IReadOnlyList<string?> fileNames = Utility.GetFileNames(resultsArray);
                if (fileNames.Count == 0) return null;

                return fileNames[0];
            }

            // User cancelled.
            return null;
        }
    }
}