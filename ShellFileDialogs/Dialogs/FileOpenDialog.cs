using System;
using System.Collections.Generic;
using System.Runtime.Versioning;
using TerraFX.Interop.Windows;

namespace ShellFileDialogs.Dialogs
{
    [SupportedOSPlatform("windows6.0")]
    public static unsafe class FileOpenDialog
    {
        public static string? ShowSingleSelectDialog(IntPtr parentHWND, string? title, string? initialDirectory,
            string? defaultFileName, IReadOnlyList<Filter>? filters, int? selectedFilterZeroBasedIndex)
        {
            var fileNames = ShowDialog(new HWND((void*)parentHWND), title, initialDirectory, defaultFileName, filters,
                selectedFilterZeroBasedIndex, default);
            if (fileNames == null) return null;
            if (fileNames.Count == 0) return null;
            return fileNames[0];
        }
        
        private static IReadOnlyList<string>? ShowDialog(HWND parentHWnd, string? title, string? initialDirectory,
            string? defaultFileName, IReadOnlyList<Filter>? filters, int? selectedFilterZeroBasedIndex,
            FILEOPENDIALOGOPTIONS flags)
        {
            using var nfod = Utility.ActivateClass<IFileOpenDialog>(typeof(TerraFX.Interop.Windows.FileOpenDialog).GUID, typeof(IFileOpenDialog).GUID);
            return ShowDialogInner(nfod, parentHWnd, title, initialDirectory, defaultFileName, filters, selectedFilterZeroBasedIndex ?? -1, flags);
        }
        
        private static IReadOnlyList<string>? ShowDialogInner(IFileOpenDialog* dialog, HWND parentHWnd, string? title,
            string? initialDirectory, string? defaultFileName, IReadOnlyList<Filter>? filters,
            int selectedFilterZeroBasedIndex, FILEOPENDIALOGOPTIONS flags)
        {
            flags = flags |
                    FILEOPENDIALOGOPTIONS.FOS_NOTESTFILECREATE |
                    FILEOPENDIALOGOPTIONS.FOS_FILEMUSTEXIST |
                    FILEOPENDIALOGOPTIONS.FOS_FORCEFILESYSTEM;
            dialog->SetOptions((uint)flags);

            fixed (char* titlePtr = title)
            {
                if (title != null) dialog->SetTitle((ushort*)titlePtr);
            }

            if (initialDirectory != null)
            {
                using var initialDirectoryShellItem = Utility.ParseShellItem2Name(initialDirectory);
                if (initialDirectoryShellItem.Get() != null)
                {
                    dialog->SetFolder((IShellItem*)initialDirectoryShellItem.Get());
                }
            }

            if (defaultFileName != null)
            {
                fixed (char* defaultFileNamePtr = defaultFileName)
                {
                    dialog->SetFileName((ushort*)defaultFileNamePtr);
                }
            }

            Utility.SetFilters((IFileDialog*)dialog, filters, selectedFilterZeroBasedIndex);

            var hr = dialog->Show(parentHWnd);
            if (!Windows.SUCCEEDED(hr))
            {
                // User cancelled.
                return null;
            }
            
            using ComPtr<IShellItemArray> arr = null; 
            dialog->GetResults(arr.GetAddressOf());

            var fileNames = Utility.GetFileNames(arr.Get());
            return fileNames;
        }
    }
}