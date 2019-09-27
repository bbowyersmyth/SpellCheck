﻿using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace MsSpellCheckLib
{
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000101-0000-0000-C000-000000000046")]
    [ComImport]
    internal interface IEnumString
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void RemoteNext([In] uint celt, [MarshalAs(UnmanagedType.LPWStr)] out string rgelt, out uint pceltFetched);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Skip([In] uint celt);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Reset();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumString ppenum);
    }
}
