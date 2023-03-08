using System.Runtime.InteropServices;
using ShellFileDialogs.Native;

namespace ShellFileDialogs
{
    [ComImport]
    [ClassInterface(ClassInterfaceType.None)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    [Guid(ShellCLSIDGuid.FileSaveDialog)]
    internal class FileSaveDialogRCW
    {
    }
}