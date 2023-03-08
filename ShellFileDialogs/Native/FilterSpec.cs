using System.Runtime.InteropServices;

namespace ShellFileDialogs.Native
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    internal struct FilterSpec
    {
        [MarshalAs(UnmanagedType.LPWStr)] internal string Name;
        [MarshalAs(UnmanagedType.LPWStr)] internal string Spec;

        internal FilterSpec(string name, string spec)
        {
            Name = name;
            Spec = spec;
        }
    }
}