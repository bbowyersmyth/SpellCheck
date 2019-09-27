using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("803E3BD4-2828-4410-8290-418D1D73C762")]
    [ComImport]
    internal interface IEnumSpellingError
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        ISpellingError Next();
    }
}
