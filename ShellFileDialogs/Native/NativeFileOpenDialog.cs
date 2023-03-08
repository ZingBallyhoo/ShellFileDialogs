using System.Runtime.InteropServices;

namespace ShellFileDialogs.Native
{
    [ComImport]
    [Guid(ShellIIDGuid.IFileOpenDialog)]
    [CoClass(typeof(FileOpenDialogRCW))]
    internal interface NativeFileOpenDialog : IFileOpenDialog
    {
    }
}