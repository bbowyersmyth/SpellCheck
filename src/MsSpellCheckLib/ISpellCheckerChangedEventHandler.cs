using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [Guid("0B83A5B0-792F-4EAB-9799-ACF52C5ED08A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    public interface ISpellCheckerChangedEventHandler
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Invoke([MarshalAs(UnmanagedType.Interface), In] ISpellChecker sender);
    }
}
