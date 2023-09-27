using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using TerraFX.Interop.Windows;

namespace ShellFileDialogs
{
    [SupportedOSPlatform("windows6.0")]
    internal static unsafe class Utility
    {
        public static IReadOnlyList<string> GetFileNames(IShellItemArray* items)
        {
            var count = 0u;
            var hresult = items->GetCount(&count);
            Windows.ThrowIfFailed(hresult);

            var list = new List<string>((int)count);

            for (var i = 0; i < count; i++)
            {
                using var shellItem = GetShellItemAt(items, i);
                var fileName = GetFileNameFromShellItem(shellItem);
                if (fileName != null) list.Add(fileName);
            }

            return list;
        }

        public static ComPtr<IShellItem2> ParseShellItem2Name(string value)
        {
            var ishellItem2GuidCopy = typeof(IShellItem2).GUID;

            var pShellItem = (void*)null;
            fixed (char* valuePtr = value)
            {
                var hresult = Windows.SHCreateItemFromParsingName((ushort*)valuePtr, null, &ishellItem2GuidCopy, &pShellItem);
                Windows.ThrowIfFailed(hresult);
            }
            
            return (IShellItem2*)pShellItem;
        }

        public static string? GetFileNameFromShellItem(IShellItem* item)
        {
            if (item == null) return null;

            var stringChars = (ushort*)null;
            var hr = item->GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, &stringChars);
            if (Windows.SUCCEEDED(hr) && stringChars != null)
            {
                return Marshal.PtrToStringUni((IntPtr)stringChars);
            }
            return null;
        }

        public static ComPtr<IShellItem> GetShellItemAt(IShellItemArray* array, int i)
        {
            if (array is null) throw new ArgumentNullException(nameof(array));

            ComPtr<IShellItem> outPtr = null;
            var hr = array->GetItemAt((uint)i, outPtr.GetAddressOf());
            Windows.ThrowIfFailed(hr);
            return outPtr;
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
        public static void SetFilters(IFileDialog* dialog, IReadOnlyList<Filter>? filters,
            int selectedFilterZeroBasedIndex)
        {
            if (dialog is null) throw new ArgumentNullException(nameof(dialog));

            if (filters == null || filters.Count == 0) return;

            // yes, i'm going mental insane
            
            var filterHeadersSize = sizeof(COMDLG_FILTERSPEC) * filters.Count;
            var filterSpecsSize = filterHeadersSize;
            foreach (var filter in filters)
            {
                filterSpecsSize += Encoding.Unicode.GetByteCount(filter.DisplayName) + 2;
                var filterStr = filter.ToFilterSpecString();
                filterSpecsSize += Encoding.Unicode.GetByteCount(filterStr) + 2;
            }
            
            var filterSpecsBuf = stackalloc byte[filterSpecsSize];
            var filterSpecsSpan = new Span<byte>(filterSpecsBuf, filterSpecsSize);
            var headers = MemoryMarshal.Cast<byte, COMDLG_FILTERSPEC>(filterSpecsSpan.Slice(0, filterHeadersSize));
            
            var writeOffset = filterHeadersSize;
            for (var i = 0; i < filters.Count; i++)
            {
                var filter = filters[i];
                
                var nameOffset = writeOffset;
                writeOffset += Encoding.Unicode.GetBytes(filter.DisplayName, filterSpecsSpan.Slice(writeOffset));
                filterSpecsSpan[writeOffset++] = 0;
                filterSpecsSpan[writeOffset++] = 0;
                
                var filterStrOffset = writeOffset;
                writeOffset += Encoding.Unicode.GetBytes(filter.ToFilterSpecString(), filterSpecsSpan.Slice(writeOffset));
                filterSpecsSpan[writeOffset++] = 0;
                filterSpecsSpan[writeOffset++] = 0;
                
                headers[i] = new COMDLG_FILTERSPEC
                {
                    pszName = (ushort*)(filterSpecsBuf + nameOffset),
                    pszSpec = (ushort*)(filterSpecsBuf + filterStrOffset)
                };
            }
            
            dialog->SetFileTypes((uint)headers.Length, (COMDLG_FILTERSPEC*)filterSpecsBuf);

            if (selectedFilterZeroBasedIndex > -1 && selectedFilterZeroBasedIndex < filters.Count)
                dialog->SetFileTypeIndex(1 + (uint)selectedFilterZeroBasedIndex); // In the COM interface (like the other Windows OFD APIs), filter indexes are 1-based, not 0-based.*/
        }
        
        public static ComPtr<TInterface> ActivateClass<TInterface>(Guid clsid, Guid iid) where TInterface : unmanaged, IUnknown.Interface
        {
            Debug.Assert(iid == typeof(TInterface).GUID);
            var ptr = new ComPtr<TInterface>();
            int hr = Windows.CoCreateInstance(&clsid, null, Windows.CLSCTX_SERVER, &iid, (void**)ptr.GetAddressOf());
            Windows.ThrowIfFailed(hr);
            return ptr;
        }
    }
}