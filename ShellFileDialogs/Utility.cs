using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using ShellFileDialogs.Native;

namespace ShellFileDialogs
{
    internal static class Utility
    {
        private static readonly Guid _ishellItem2Guid = new Guid(ShellIIDGuid.IShellItem2);

        public static IReadOnlyList<string> GetFileNames(IShellItemArray items)
        {
            var hresult = items.GetCount(out var count);
            if (hresult != HResult.Ok)
                throw new Exception("IShellItemArray.GetCount failed. HResult: " +
                                    hresult); // TODO: Will this ever happen?

            var list = new List<string>((int)count);

            for (var i = 0; i < count; i++)
            {
                var shellItem = GetShellItemAt(items, i);
                var fileName = GetFileNameFromShellItem(shellItem);
                if (fileName != null) list.Add(fileName);
            }

            return list;
        }

        public static IShellItem2? ParseShellItem2Name(string value)
        {
            var ishellItem2GuidCopy = _ishellItem2Guid;

            var hresult =
                ShellNativeMethods.SHCreateItemFromParsingName(value, IntPtr.Zero, ref ishellItem2GuidCopy,
                    out var shellItem);
            if (hresult == HResult.Ok) return shellItem;

            // TODO: Handle HRESULT error codes?
            return null;
        }

        public static string? GetFileNameFromShellItem(IShellItem? item)
        {
            if (item is null) return null;

            var hr = item.GetDisplayName(ShellItemDesignNameOptions.DesktopAbsoluteParsing, out var pszString);
            if (hr == HResult.Ok && pszString != IntPtr.Zero)
            {
                var fileName =
                    Marshal.PtrToStringAuto(
                        pszString)!; // `PtrToStringAuto` won't return `null` if its `ptr` argument is not null, which we check for.
                Marshal.FreeCoTaskMem(pszString);
                return fileName;
            }

            return null;
        }

        public static IShellItem? GetShellItemAt(IShellItemArray array, int i)
        {
            if (array is null) throw new ArgumentNullException(nameof(array));

            var hr = array.GetItemAt((uint)i, out var result);
            if (hr == HResult.Ok) return result;

            return null;
        }

        /// <summary>Sets the file extension filters on <paramref name="dialog" />.</summary>
        /// <param name="dialog">Required. Cannot be <see langword="null" />.</param>
        /// <param name="filters">
        ///     If this is <see langword="null" /> or empty, then this method returns immediately (i.e. it does
        ///     nothing).
        /// </param>
        /// <param name="selectedFilterZeroBasedIndex">
        ///     0-based index of the filter in in <paramref name="filters" /> to use. If
        ///     this value is out-of-range then this method does nothing.
        /// </param>
        public static void SetFilters(IFileDialog dialog, IReadOnlyCollection<Filter>? filters,
            int selectedFilterZeroBasedIndex)
        {
            if (dialog is null) throw new ArgumentNullException(nameof(dialog));

            if (filters == null || filters.Count == 0) return;

            var specs = CreateFilterSpec(filters);
            dialog.SetFileTypes((uint)specs.Length, specs);

            if (selectedFilterZeroBasedIndex > -1 && selectedFilterZeroBasedIndex < filters.Count)
                dialog.SetFileTypeIndex(1 +
                                        (uint)selectedFilterZeroBasedIndex); // In the COM interface (like the other Windows OFD APIs), filter indexes are 1-based, not 0-based.
        }

        public static FilterSpec[] CreateFilterSpec(IReadOnlyCollection<Filter> filters)
        {
            var specs = new FilterSpec[filters.Count];
            var i = 0;
            foreach (var filter in filters)
            {
                specs[i] = filter.ToFilterSpec();
                i++;
            }

            return specs;
        }

        /// <summary>
        ///     Returns <see langword="false" /> if the user cancelled-out of the dialog. Returns <see langword="true" /> if
        ///     the user completed the dialog. All other cases result in a thrown <see cref="Win32Exception" /> or
        ///     <see cref="ExternalException" /> depending on the HRESULT returned from <see cref="IModalWindow.Show(IntPtr)" />.
        /// </summary>
        public static bool ValidateDialogShowHResult(this HResult dialogHResult)
        {
            if (dialogHResult.TryGetWin32ErrorCode(out var win32Code))
            {
                if (win32Code == Win32ErrorCodes.Success)
                    // OK.
                    return true;

                if (win32Code == Win32ErrorCodes.ErrorCancelled)
                    // Cancelled
                    return false;
                // Other Win32 error:

                var msg = string.Format(CultureInfo.CurrentCulture,
                    "Unexpected Win32 error code 0x{0:X2} in HRESULT 0x{1:X4} returned from IModalWindow.Show(...).",
                    (int)win32Code, (int)dialogHResult);
                throw new Win32Exception((int)win32Code, msg);
            }

            if (dialogHResult.IsValidHResult())
            {
                const ushort RPC_E_SERVERFAULT = 0x0105;

                if (dialogHResult.GetFacility() == HResultFacility.Rpc && dialogHResult.GetCode() == RPC_E_SERVERFAULT)
                {
                    // This error happens when calling `IModalWindow.Show` instead of using the `Show` method on a different interface, like `IFileOpenDialog.Show`.
                    var msg = string.Format(CultureInfo.CurrentCulture,
                        "Unexpected RPC HRESULT: 0x{0:X4} (RPC Error {1:X2}) returned from IModalWindow.Show(...). This particular RPC error suggests the dialog was accessed via the wrong COM interface.",
                        (int)dialogHResult, RPC_E_SERVERFAULT);
                    throw new ExternalException(msg, (int)dialogHResult);
                }
                // Fall-through to below:
            }

            // Fall-through to below:
            {
                // Other HRESULT (non-Win32 error):
                // https://stackoverflow.com/questions/11158379/how-can-i-throw-an-exception-with-a-certain-hresult

                var msg = string.Format(CultureInfo.CurrentCulture,
                    "Unexpected HRESULT: 0x{0:X4} returned from IModalWindow.Show(...).", (int)dialogHResult);
                throw new ExternalException(msg, (int)dialogHResult);
            }
        }
    }
}