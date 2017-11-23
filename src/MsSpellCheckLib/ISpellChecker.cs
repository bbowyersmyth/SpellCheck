using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B")]
    [ComImport]
    public interface ISpellChecker
    {
        [DispId(1610678272)]
        string languageTag
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [return: MarshalAs(UnmanagedType.LPWStr)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr), In] string text);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr), In] string word);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Add([MarshalAs(UnmanagedType.LPWStr), In] string word);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Ignore([MarshalAs(UnmanagedType.LPWStr), In] string word);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void AutoCorrect([MarshalAs(UnmanagedType.LPWStr), In] string from, [MarshalAs(UnmanagedType.LPWStr), In] string to);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        byte GetOptionValue([MarshalAs(UnmanagedType.LPWStr), In] string optionId);

        [DispId(1610678279)]
        IEnumString OptionIds
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [return: MarshalAs(UnmanagedType.Interface)]
            get;
        }

        [DispId(1610678280)]
        string Id
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [return: MarshalAs(UnmanagedType.LPWStr)]
            get;
        }

        [DispId(1610678281)]
        string LocalizedName
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            [return: MarshalAs(UnmanagedType.LPWStr)]
            get;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        uint add_SpellCheckerChanged([MarshalAs(UnmanagedType.Interface), In] ISpellCheckerChangedEventHandler handler);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void remove_SpellCheckerChanged([In] uint eventCookie);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        IOptionDescription GetOptionDescription([MarshalAs(UnmanagedType.LPWStr), In] string optionId);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSpellingError ComprehensiveCheck([MarshalAs(UnmanagedType.LPWStr), In] string text);
    }
}
