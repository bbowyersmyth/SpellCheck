using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [Guid("8E018A9D-2415-4677-BF08-794EA61F94BB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    internal interface ISpellCheckerFactory
    {
        [DispId(1610678272)]
        IEnumString SupportedLanguages
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        int IsSupported([MarshalAs(UnmanagedType.LPWStr), In] string languageTag);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr), In] string languageTag);
    }
}
