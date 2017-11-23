using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("E7ED1C71-87F7-4378-A840-C9200DACEE47")]
    [ComImport]
    public interface ISpellChecker2 : ISpellChecker
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Remove([MarshalAs(UnmanagedType.LPWStr), In] string word);
    }
}
