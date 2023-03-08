using System.Runtime.InteropServices;

namespace ShellFileDialogs.Native
{
    [ComImport]
    [ClassInterface(ClassInterfaceType.None)]
    [TypeLibType(TypeLibTypeFlags.FCanCreate)]
    [Guid(ShellCLSIDGuid.FileOpenDialog)]
    internal class FileOpenDialogRCW
    {
    }
}