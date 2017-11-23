using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [Guid("AA176B85-0E12-4844-8E1A-EEF1DA77F586")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface IUserDictionariesRegistrar
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void RegisterUserDictionary([MarshalAs(UnmanagedType.LPWStr), In] string dictionaryPath, [MarshalAs(UnmanagedType.LPWStr), In] string languageTag);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void UnregisterUserDictionary([MarshalAs(UnmanagedType.LPWStr), In] string dictionaryPath, [MarshalAs(UnmanagedType.LPWStr), In] string languageTag);
    }
}
