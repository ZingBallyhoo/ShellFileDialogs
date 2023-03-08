using System.Runtime.InteropServices;

namespace ShellFileDialogs.Native
{
    [ComImport]
    [Guid(ShellIIDGuid.IFileSaveDialog)]
    [CoClass(typeof(FileSaveDialogRCW))]
    internal interface NativeFileSaveDialog : IFileSaveDialog
    {
    }
}