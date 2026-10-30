# VBA Coder COM type-library metadata extractor (Windows PowerShell 5.1).
#
# Reads public type metadata from a .tlb, .olb or .dll with LoadTypeLibEx(REGKIND_NONE).
# It never registers the library, never loads it as an executable, never starts Office and
# never invokes any component. Output is UTF-8 JSON without BOM (see schemas/type-index.schema.json).
#
# This script is intentionally ASCII-only: Windows PowerShell 5.1 decodes BOM-less script files
# with the ANSI code page, so non-ASCII text here would be misread on non-English systems.
#
# Usage:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File export-com.ps1 `
#     -InputPath <tlb|olb|dll> -OutputPath <json> -Bitness <32|64>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [Parameter(Mandatory = $true)][ValidateSet('32', '64')][string]$Bitness
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace VbaCoder
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct TLIBATTR
    {
        public Guid guid;
        public int lcid;
        public int syskind;
        public ushort wMajorVerNum;
        public ushort wMinorVerNum;
        public ushort wLibFlags;
    }

    internal enum TYPEKIND
    {
        TKIND_ENUM = 0,
        TKIND_RECORD = 1,
        TKIND_MODULE = 2,
        TKIND_INTERFACE = 3,
        TKIND_DISPATCH = 4,
        TKIND_COCLASS = 5,
        TKIND_ALIAS = 6,
        TKIND_UNION = 7,
        TKIND_MAX = 8
    }

    internal enum INVOKEKIND
    {
        INVOKE_FUNC = 1,
        INVOKE_PROPERTYGET = 2,
        INVOKE_PROPERTYPUT = 4,
        INVOKE_PROPERTYPUTREF = 8
    }

    internal enum VARKIND
    {
        VAR_PERINSTANCE = 0,
        VAR_STATIC = 1,
        VAR_CONST = 2,
        VAR_DISPATCH = 3
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct TYPEDESC
    {
        public IntPtr unionValue;
        public ushort vt;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PARAMDESC
    {
        public IntPtr pparamdescex;
        public ushort wParamFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct ELEMDESC
    {
        public TYPEDESC tdesc;
        public PARAMDESC paramdesc;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct IDLDESC
    {
        public IntPtr dwReserved;
        public ushort wIDLFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct FUNCDESC
    {
        public int memid;
        public IntPtr lprgscode;
        public IntPtr lprgElemDesc;
        public int funckind;
        public int invkind;
        public int callconv;
        public short cParams;
        public short cParamsOpt;
        public short oVft;
        public short cScodes;
        public ELEMDESC elemdescFunc;
        public ushort wFuncFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct VARDESC
    {
        public int memid;
        public IntPtr lpstrSchema;
        public IntPtr unionValue;
        public ELEMDESC elemdescVar;
        public ushort wVarFlags;
        public int varkind;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct TYPEATTR
    {
        public Guid guid;
        public int lcid;
        public int memidConstructor;
        public int memidDestructor;
        public IntPtr lpstrSchema;
        public int cbSizeInstance;
        public int typekind;
        public ushort cFuncs;
        public ushort cVars;
        public ushort cImplTypes;
        public ushort cbSizeVft;
        public ushort cbAlignment;
        public ushort wTypeFlags;
        public ushort wMajorVerNum;
        public ushort wMinorVerNum;
        public TYPEDESC tdescAlias;
        public IDLDESC idldescType;
    }

    [ComImport]
    [Guid("00020402-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITypeLib
    {
        // ITypeLib::GetTypeInfoCount returns UINT directly: it is one of the few COM methods
        // without an HRESULT. Declaring it as void(out uint) makes the CLR read the count as
        // an HRESULT and leaves the out parameter untouched, yielding zero types.
        [PreserveSig] uint GetTypeInfoCount();
        void GetTypeInfo(uint index, out ITypeInfo ppTInfo);
        void GetTypeInfoType(uint index, out TYPEKIND ptKind);
        void GetTypeInfoOfGuid(ref Guid guid, out ITypeInfo ppTinfo);
        void GetLibAttr(out IntPtr ppTLibAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetDocumentation(int index, [MarshalAs(UnmanagedType.BStr)] out string strName, [MarshalAs(UnmanagedType.BStr)] out string strDocString, out int dwHelpContext, [MarshalAs(UnmanagedType.BStr)] out string strHelpFile);
        [PreserveSig] int IsName(IntPtr szNameBuf, int lHashVal, IntPtr pfName);
        [PreserveSig] int FindName(IntPtr szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound);
        // ReleaseTLibAttr returns void in oleaut32. Without PreserveSig the CLR would treat
        // the leftover return register as an HRESULT and throw a bogus COMException.
        [PreserveSig] void ReleaseTLibAttr(IntPtr pTLibAttr);
    }

    [ComImport]
    [Guid("00020401-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITypeInfo
    {
        void GetTypeAttr(out IntPtr ppTypeAttr);
        void GetTypeComp(out IntPtr ppTComp);
        void GetFuncDesc(uint index, out IntPtr ppFuncDesc);
        void GetVarDesc(uint index, out IntPtr ppVarDesc);
        void GetNames(int memid, IntPtr rgBstrNames, uint cMaxNames, out uint pcNames);
        void GetRefTypeOfImplType(uint index, out uint href);
        void GetImplTypeFlags(uint index, out int pImplTypeFlags);
        void GetIDsOfNames(IntPtr szNames, IntPtr rgMemId, ushort cNames, out ushort pcNames);
        void Invoke(IntPtr pvInstance, int memid, short wFlags, IntPtr pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
        void GetDocumentation(int memid, [MarshalAs(UnmanagedType.BStr)] out string strName, [MarshalAs(UnmanagedType.BStr)] out string strDocString, out int dwHelpContext, [MarshalAs(UnmanagedType.BStr)] out string strHelpFile);
        void GetDllEntry(int memid, int invKind, IntPtr pBstrDllName, IntPtr pBstrName, IntPtr pwOrdinal);
        void GetRefTypeInfo(uint hRefType, out ITypeInfo ppTInfo);
        void AddressOfMember(int memid, int invKind, out IntPtr ppv);
        void CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObj);
        void GetMops(int memid, IntPtr pMops);
        void GetContainingTypeLib(out ITypeLib ppTLB, out uint pIndex);
        [PreserveSig] void ReleaseTypeAttr(IntPtr pTypeAttr);
        [PreserveSig] void ReleaseFuncDesc(IntPtr pFuncDesc);
        [PreserveSig] void ReleaseVarDesc(IntPtr pVarDesc);
    }

    internal static class NativeMethods
    {
        // REGKIND_NONE = 0: load for inspection only, do not touch the registry.
        // For DLL/EXE files oleaut32 loads the image as a data file, so DllMain does not run.
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void LoadTypeLibEx(
            [MarshalAs(UnmanagedType.LPWStr)] string fileName,
            int registrationKind,
            [MarshalAs(UnmanagedType.Interface)] out ITypeLib typeLibrary);
    }

    public sealed class TypeLibraryResult
    {
        public string LibraryName;
        public string LibraryGuid;
        public int Major;
        public int Minor;
        public int Lcid;
        public List<Dictionary<string, object>> Types = new List<Dictionary<string, object>>();
    }

    public static class TypeLibraryReader
    {
        private const int TypeFlagHidden = 0x0040;
        private const int TypeFlagRestricted = 0x0100;
        private const int FuncFlagHidden = 0x0040;
        private const int FuncFlagRestricted = 0x0100;
        private const int VarFlagHidden = 0x0040;
        private const int VarFlagRestricted = 0x0100;

        private const int ImplTypeFlagSource = 0x0002;
        private const int ImplTypeFlagRestricted = 0x0004;
        private const int MaxImplDepth = 32;

        private const int ParamFlagRetval = 0x0008;
        private const int ParamFlagOptional = 0x0010;
        private const int ParamFlagHasDefault = 0x0020;

        private const int VtEmpty = 0;
        private const int VtNull = 1;
        private const int VtVoid = 24;
        private const int VtHresult = 25;
        private const int VtPtr = 26;
        private const int VtSafeArray = 27;
        private const int VtCArray = 28;
        private const int VtUserDefined = 29;
        private const int VtArray = 0x2000;
        private const int VtByRef = 0x4000;
        private const int VtTypeMask = 0x0FFF;

        private sealed class PendingMember
        {
            public Dictionary<string, object> Value;
            public bool IsGetter;

            public PendingMember(Dictionary<string, object> value, bool isGetter)
            {
                Value = value;
                IsGetter = isGetter;
            }
        }

        public static TypeLibraryResult Read(string path)
        {
            ITypeLib library;
            NativeMethods.LoadTypeLibEx(path, 0, out library);
            try
            {
                TypeLibraryResult result = new TypeLibraryResult();

                IntPtr attrPointer;
                library.GetLibAttr(out attrPointer);
                TLIBATTR libraryAttr;
                try
                {
                    libraryAttr = (TLIBATTR)Marshal.PtrToStructure(attrPointer, typeof(TLIBATTR));
                }
                finally
                {
                    library.ReleaseTLibAttr(attrPointer);
                }

                result.LibraryGuid = libraryAttr.guid.ToString("B").ToUpperInvariant();
                result.Major = libraryAttr.wMajorVerNum;
                result.Minor = libraryAttr.wMinorVerNum;
                result.Lcid = libraryAttr.lcid;

                string libraryName;
                string libraryDocumentation;
                int helpContext;
                string helpFile;
                library.GetDocumentation(-1, out libraryName, out libraryDocumentation, out helpContext, out helpFile);
                result.LibraryName = string.IsNullOrEmpty(libraryName)
                    ? Path.GetFileNameWithoutExtension(path)
                    : libraryName;

                uint typeCount;
                typeCount = library.GetTypeInfoCount();
                for (uint index = 0; index < typeCount; index++)
                {
                    TYPEKIND kind;
                    library.GetTypeInfoType(index, out kind);
                    string mappedKind = MapTypeKind(kind);
                    if (mappedKind == null)
                    {
                        continue;
                    }

                    ITypeInfo info;
                    library.GetTypeInfo(index, out info);
                    try
                    {
                        Dictionary<string, object> type = ReadType(info, kind, mappedKind);
                        if (type != null)
                        {
                            result.Types.Add(type);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(info);
                    }
                }

                return result;
            }
            finally
            {
                Marshal.ReleaseComObject(library);
            }
        }

        private static string MapTypeKind(TYPEKIND kind)
        {
            switch (kind)
            {
                case TYPEKIND.TKIND_ENUM:
                    return "enum";
                case TYPEKIND.TKIND_MODULE:
                    return "module";
                case TYPEKIND.TKIND_INTERFACE:
                    return "interface";
                case TYPEKIND.TKIND_DISPATCH:
                    return "interface";
                case TYPEKIND.TKIND_COCLASS:
                    return "class";
                default:
                    // Records, unions and aliases have no PortableTypeIndex representation.
                    return null;
            }
        }

        private static Dictionary<string, object> ReadType(ITypeInfo info, TYPEKIND kind, string mappedKind)
        {
            IntPtr attrPointer;
            info.GetTypeAttr(out attrPointer);
            TYPEATTR attr;
            try
            {
                attr = (TYPEATTR)Marshal.PtrToStructure(attrPointer, typeof(TYPEATTR));
            }
            finally
            {
                info.ReleaseTypeAttr(attrPointer);
            }

            if ((attr.wTypeFlags & (TypeFlagHidden | TypeFlagRestricted)) != 0)
            {
                return null;
            }

            string name;
            string documentation;
            int helpContext;
            string helpFile;
            info.GetDocumentation(-1, out name, out documentation, out helpContext, out helpFile);
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            List<object> members = new List<object>();
            if (kind == TYPEKIND.TKIND_COCLASS)
            {
                // A coclass publishes no members of its own. Flatten its non-source interfaces
                // and expose source interfaces as events so New StdFont completion is not empty.
                ReadCoclassMembers(info, attr.cImplTypes, members);
            }
            else if (kind == TYPEKIND.TKIND_ENUM)
            {
                ReadConstants(info, attr.cVars, members);
            }
            else
            {
                ReadFunctions(info, attr.cFuncs, members, null);
                if (kind == TYPEKIND.TKIND_MODULE)
                {
                    ReadConstants(info, attr.cVars, members);
                }
                else if (kind == TYPEKIND.TKIND_DISPATCH)
                {
                    // Dispinterface properties are published as VAR_DISPATCH variables, not functions.
                    ReadDispatchProperties(info, attr.cVars, members, null);
                }
            }

            Dictionary<string, object> type = new Dictionary<string, object>();
            type["name"] = name;
            type["kind"] = mappedKind;
            if (!string.IsNullOrEmpty(documentation))
            {
                type["documentation"] = documentation;
            }
            type["members"] = members;
            return type;
        }

        private static void ReadFunctions(ITypeInfo info, int count, List<object> members, string forcedKind)
        {
            List<string> order = new List<string>();
            Dictionary<string, PendingMember> map = new Dictionary<string, PendingMember>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < count; index++)
            {
                IntPtr pointer;
                info.GetFuncDesc((uint)index, out pointer);
                try
                {
                    FUNCDESC description = (FUNCDESC)Marshal.PtrToStructure(pointer, typeof(FUNCDESC));
                    if ((description.wFuncFlags & (FuncFlagHidden | FuncFlagRestricted)) != 0)
                    {
                        continue;
                    }

                    string rawName;
                    string documentation;
                    int helpContext;
                    string helpFile;
                    info.GetDocumentation(description.memid, out rawName, out documentation, out helpContext, out helpFile);
                    if (string.IsNullOrEmpty(rawName))
                    {
                        continue;
                    }

                    string name = rawName;
                    string memberKind;
                    bool isGetter = false;
                    if (forcedKind != null)
                    {
                        // Source interfaces are exposed as events; keep the declared name.
                        memberKind = forcedKind;
                    }
                    else if (description.invkind == (int)INVOKEKIND.INVOKE_PROPERTYGET)
                    {
                        memberKind = "property";
                        isGetter = true;
                    }
                    else if (description.invkind == (int)INVOKEKIND.INVOKE_PROPERTYPUT ||
                             description.invkind == (int)INVOKEKIND.INVOKE_PROPERTYPUTREF)
                    {
                        memberKind = "property";
                    }
                    else
                    {
                        // Dispatch interfaces often expose accessors as get_X / put_X functions.
                        memberKind = "method";
                        if (rawName.StartsWith("get_", StringComparison.OrdinalIgnoreCase))
                        {
                            memberKind = "property";
                            name = rawName.Substring(4);
                            isGetter = true;
                        }
                        else if (rawName.StartsWith("putref_", StringComparison.OrdinalIgnoreCase))
                        {
                            memberKind = "property";
                            name = rawName.Substring(7);
                        }
                        else if (rawName.StartsWith("put_", StringComparison.OrdinalIgnoreCase))
                        {
                            memberKind = "property";
                            name = rawName.Substring(4);
                        }
                    }

                    if (name.Length == 0)
                    {
                        continue;
                    }

                    string returnType;
                    List<object> parameters = ReadParameters(info, description, out returnType);

                    Dictionary<string, object> value = new Dictionary<string, object>();
                    value["name"] = name;
                    value["kind"] = memberKind;
                    if (returnType != null)
                    {
                        value["returnType"] = returnType;
                    }
                    if (parameters.Count > 0)
                    {
                        value["parameters"] = parameters;
                    }
                    if (!string.IsNullOrEmpty(documentation))
                    {
                        value["documentation"] = documentation;
                    }

                    string key = memberKind + ":" + name;
                    PendingMember pending;
                    if (!map.TryGetValue(key, out pending))
                    {
                        map[key] = new PendingMember(value, isGetter);
                        order.Add(key);
                    }
                    else if (isGetter && !pending.IsGetter)
                    {
                        pending.Value = value;
                        pending.IsGetter = true;
                    }
                }
                finally
                {
                    info.ReleaseFuncDesc(pointer);
                }
            }

            foreach (string key in order)
            {
                members.Add(map[key].Value);
            }
        }

        private static List<object> ReadParameters(ITypeInfo info, FUNCDESC description, out string returnType)
        {
            returnType = null;
            int count = description.cParams;
            List<object> parameters = new List<object>();
            string[] names = ReadNames(info, description.memid, count);
            int stride = Marshal.SizeOf(typeof(ELEMDESC));

            for (int index = 0; index < count; index++)
            {
                IntPtr elementPointer = new IntPtr(description.lprgElemDesc.ToInt64() + (long)index * stride);
                ELEMDESC element = (ELEMDESC)Marshal.PtrToStructure(elementPointer, typeof(ELEMDESC));
                string typeName = FormatType(info, element.tdesc);

                if ((element.paramdesc.wParamFlags & ParamFlagRetval) != 0)
                {
                    if (returnType == null)
                    {
                        returnType = typeName;
                    }
                    continue;
                }

                string parameterName = index < names.Length ? names[index] : null;
                if (string.IsNullOrEmpty(parameterName))
                {
                    parameterName = "Arg" + (index + 1).ToString(CultureInfo.InvariantCulture);
                }

                Dictionary<string, object> parameter = new Dictionary<string, object>();
                parameter["name"] = parameterName;
                parameter["typeName"] = typeName;
                parameter["optional"] = (element.paramdesc.wParamFlags & (ParamFlagOptional | ParamFlagHasDefault)) != 0;
                parameter["byRef"] = (element.tdesc.vt & VtByRef) != 0;
                parameters.Add(parameter);
            }

            if (returnType == null)
            {
                int returnVt = description.elemdescFunc.tdesc.vt & VtTypeMask;
                if (returnVt != VtEmpty && returnVt != VtNull && returnVt != VtVoid && returnVt != VtHresult)
                {
                    returnType = FormatType(info, description.elemdescFunc.tdesc);
                }
            }

            return parameters;
        }

        private static void ReadConstants(ITypeInfo info, int count, List<object> members)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < count; index++)
            {
                IntPtr pointer;
                info.GetVarDesc((uint)index, out pointer);
                try
                {
                    VARDESC description = (VARDESC)Marshal.PtrToStructure(pointer, typeof(VARDESC));
                    if ((description.wVarFlags & (VarFlagHidden | VarFlagRestricted)) != 0)
                    {
                        continue;
                    }
                    if (description.varkind != (int)VARKIND.VAR_CONST)
                    {
                        continue;
                    }

                    string name;
                    string documentation;
                    int helpContext;
                    string helpFile;
                    info.GetDocumentation(description.memid, out name, out documentation, out helpContext, out helpFile);
                    if (string.IsNullOrEmpty(name) || !seen.Add(name))
                    {
                        continue;
                    }

                    Dictionary<string, object> member = new Dictionary<string, object>();
                    member["name"] = name;
                    member["kind"] = "constant";
                    member["returnType"] = FormatType(info, description.elemdescVar.tdesc);
                    if (!string.IsNullOrEmpty(documentation))
                    {
                        member["documentation"] = documentation;
                    }
                    members.Add(member);
                }
                finally
                {
                    info.ReleaseVarDesc(pointer);
                }
            }
        }

        private static void ReadDispatchProperties(ITypeInfo info, int count, List<object> members, string forcedKind)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < members.Count; index++)
            {
                Dictionary<string, object> existing = members[index] as Dictionary<string, object>;
                if (existing == null || !existing.ContainsKey("name"))
                {
                    continue;
                }
                string existingName = existing["name"] as string;
                if (!string.IsNullOrEmpty(existingName))
                {
                    seen.Add(existingName);
                }
            }

            for (int index = 0; index < count; index++)
            {
                IntPtr pointer;
                info.GetVarDesc((uint)index, out pointer);
                try
                {
                    VARDESC description = (VARDESC)Marshal.PtrToStructure(pointer, typeof(VARDESC));
                    if ((description.wVarFlags & (VarFlagHidden | VarFlagRestricted)) != 0)
                    {
                        continue;
                    }
                    if (description.varkind != (int)VARKIND.VAR_DISPATCH)
                    {
                        continue;
                    }

                    string name;
                    string documentation;
                    int helpContext;
                    string helpFile;
                    info.GetDocumentation(description.memid, out name, out documentation, out helpContext, out helpFile);
                    if (string.IsNullOrEmpty(name) || !seen.Add(name))
                    {
                        continue;
                    }

                    Dictionary<string, object> member = new Dictionary<string, object>();
                    member["name"] = name;
                    member["kind"] = forcedKind == null ? "property" : forcedKind;
                    member["returnType"] = FormatType(info, description.elemdescVar.tdesc);
                    if (!string.IsNullOrEmpty(documentation))
                    {
                        member["documentation"] = documentation;
                    }
                    members.Add(member);
                }
                finally
                {
                    info.ReleaseVarDesc(pointer);
                }
            }
        }

        private static string GetTypeName(ITypeInfo info)
        {
            string name;
            string documentation;
            int helpContext;
            string helpFile;
            try
            {
                info.GetDocumentation(-1, out name, out documentation, out helpContext, out helpFile);
            }
            catch
            {
                return null;
            }
            return name;
        }

        private static bool IsAutomationRoot(string name)
        {
            return string.Equals(name, "IUnknown", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "IDispatch", StringComparison.OrdinalIgnoreCase);
        }

        private static void AddUniqueMembers(List<object> target, List<object> source, HashSet<string> seen)
        {
            for (int index = 0; index < source.Count; index++)
            {
                Dictionary<string, object> member = source[index] as Dictionary<string, object>;
                if (member == null || !member.ContainsKey("name"))
                {
                    continue;
                }
                string name = member["name"] as string;
                if (string.IsNullOrEmpty(name) || !seen.Add(name))
                {
                    continue;
                }
                target.Add(member);
            }
        }

        private static void ReadCoclassMembers(ITypeInfo info, int count, List<object> members)
        {
            // Coordinator of the coclass interfaces. Non-source interfaces contribute their
            // members; source interfaces (event sinks) contribute event members.
            HashSet<string> visitedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> seenMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < count; index++)
            {
                int flags;
                try
                {
                    info.GetImplTypeFlags((uint)index, out flags);
                }
                catch
                {
                    continue;
                }
                if ((flags & ImplTypeFlagRestricted) != 0)
                {
                    continue;
                }

                uint href;
                ITypeInfo referenced;
                try
                {
                    info.GetRefTypeOfImplType((uint)index, out href);
                    info.GetRefTypeInfo(href, out referenced);
                }
                catch
                {
                    continue;
                }

                try
                {
                    bool asEvents = (flags & ImplTypeFlagSource) != 0;
                    AppendInterfaceMembers(referenced, members, visitedTypes, seenMembers, asEvents, 0);
                }
                finally
                {
                    Marshal.ReleaseComObject(referenced);
                }
            }
        }

        private static void AppendInterfaceMembers(
            ITypeInfo info,
            List<object> members,
            HashSet<string> visitedTypes,
            HashSet<string> seenMembers,
            bool asEvents,
            int depth)
        {
            if (depth > MaxImplDepth)
            {
                return;
            }

            string typeName = GetTypeName(info);
            if (string.IsNullOrEmpty(typeName) || IsAutomationRoot(typeName) || !visitedTypes.Add(typeName))
            {
                return;
            }

            IntPtr attrPointer;
            info.GetTypeAttr(out attrPointer);
            TYPEATTR attr;
            try
            {
                attr = (TYPEATTR)Marshal.PtrToStructure(attrPointer, typeof(TYPEATTR));
            }
            finally
            {
                info.ReleaseTypeAttr(attrPointer);
            }

            if ((attr.wTypeFlags & (TypeFlagHidden | TypeFlagRestricted)) != 0)
            {
                return;
            }

            List<object> local = new List<object>();
            if (asEvents)
            {
                ReadFunctions(info, attr.cFuncs, local, "event");
                if (attr.typekind == (int)TYPEKIND.TKIND_DISPATCH)
                {
                    ReadDispatchProperties(info, attr.cVars, local, "event");
                }
            }
            else
            {
                ReadFunctions(info, attr.cFuncs, local, null);
                if (attr.typekind == (int)TYPEKIND.TKIND_DISPATCH)
                {
                    ReadDispatchProperties(info, attr.cVars, local, null);
                }
            }
            AddUniqueMembers(members, local, seenMembers);

            // Follow only interfaces of the same direction: a member interface never inherits
            // events, and an event sink never inherits object members.
            for (int index = 0; index < attr.cImplTypes; index++)
            {
                int flags;
                try
                {
                    info.GetImplTypeFlags((uint)index, out flags);
                }
                catch
                {
                    continue;
                }
                if ((flags & ImplTypeFlagRestricted) != 0)
                {
                    continue;
                }
                bool source = (flags & ImplTypeFlagSource) != 0;
                if (source != asEvents)
                {
                    continue;
                }

                uint href;
                ITypeInfo baseInfo;
                try
                {
                    info.GetRefTypeOfImplType((uint)index, out href);
                    info.GetRefTypeInfo(href, out baseInfo);
                }
                catch
                {
                    continue;
                }

                try
                {
                    AppendInterfaceMembers(baseInfo, members, visitedTypes, seenMembers, asEvents, depth + 1);
                }
                finally
                {
                    Marshal.ReleaseComObject(baseInfo);
                }
            }
        }

        private static string[] ReadNames(ITypeInfo info, int memberId, int count)
        {
            string[] names = new string[count];
            int capacity = count + 1;
            IntPtr buffer = IntPtr.Zero;
            try
            {
                buffer = Marshal.AllocHGlobal(IntPtr.Size * capacity);
                for (int index = 0; index < capacity; index++)
                {
                    Marshal.WriteIntPtr(buffer, index * IntPtr.Size, IntPtr.Zero);
                }

                uint received;
                info.GetNames(memberId, buffer, (uint)capacity, out received);
                for (int index = 0; index < count; index++)
                {
                    IntPtr text = Marshal.ReadIntPtr(buffer, (index + 1) * IntPtr.Size);
                    if (text != IntPtr.Zero)
                    {
                        names[index] = Marshal.PtrToStringBSTR(text);
                        Marshal.FreeBSTR(text);
                    }
                }
            }
            catch
            {
                // Names are optional; persistable parameter names fall back to ArgN.
                return new string[count];
            }
            finally
            {
                if (buffer != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            return names;
        }

        private static string FormatType(ITypeInfo owner, TYPEDESC description)
        {
            int vt = description.vt & VtTypeMask;
            bool isArray = (description.vt & VtArray) != 0;
            string text;

            switch (vt)
            {
                case 0: text = "Variant"; break;
                case 1: text = "Variant"; break;
                case 2: text = "Integer"; break;
                case 3: text = "Long"; break;
                case 4: text = "Single"; break;
                case 5: text = "Double"; break;
                case 6: text = "Currency"; break;
                case 7: text = "Date"; break;
                case 8: text = "String"; break;
                case 9: text = "Object"; break;
                case 10: text = "Long"; break;
                case 11: text = "Boolean"; break;
                case 12: text = "Variant"; break;
                case 13: text = "Object"; break;
                case 14: text = "Variant"; break;
                case 16: text = "Byte"; break;
                case 17: text = "Byte"; break;
                case 18: text = "Integer"; break;
                case 19: text = "Long"; break;
                case 20: text = "LongLong"; break;
                case 21: text = "LongLong"; break;
                case 22: text = "Long"; break;
                case 23: text = "Long"; break;
                case 24: text = "Variant"; break;
                case 25: text = "Long"; break;
                case VtPtr:
                    text = description.unionValue == IntPtr.Zero
                        ? "Variant"
                        : FormatType(owner, (TYPEDESC)Marshal.PtrToStructure(description.unionValue, typeof(TYPEDESC)));
                    break;
                case VtSafeArray:
                case VtCArray:
                    if (description.unionValue == IntPtr.Zero)
                    {
                        text = "Variant";
                    }
                    else
                    {
                        // ARRAYDESC starts with tdescElem, so reusing TYPEDESC is valid.
                        text = FormatType(owner, (TYPEDESC)Marshal.PtrToStructure(description.unionValue, typeof(TYPEDESC))) + "()";
                        isArray = false;
                    }
                    break;
                case VtUserDefined:
                    text = UserDefinedName(owner, description);
                    break;
                case 30:
                case 31:
                    text = "String";
                    break;
                case 37:
                case 38:
                    text = "LongPtr";
                    break;
                default:
                    text = "Variant";
                    break;
            }

            if (isArray)
            {
                text = text + "()";
            }
            return text;
        }

        private static string UserDefinedName(ITypeInfo owner, TYPEDESC description)
        {
            if (description.unionValue == IntPtr.Zero)
            {
                return "Variant";
            }

            ITypeInfo referenced;
            try
            {
                owner.GetRefTypeInfo((uint)description.unionValue.ToInt64(), out referenced);
            }
            catch
            {
                return "Variant";
            }

            try
            {
                string name;
                string documentation;
                int helpContext;
                string helpFile;
                referenced.GetDocumentation(-1, out name, out documentation, out helpContext, out helpFile);
                return string.IsNullOrEmpty(name) ? "Variant" : name;
            }
            catch
            {
                return "Variant";
            }
            finally
            {
                Marshal.ReleaseComObject(referenced);
            }
        }
    }
}
'@

try {
    $resolvedInput = [System.IO.Path]::GetFullPath($InputPath)
    if (-not [System.IO.File]::Exists($resolvedInput)) {
        throw "Input file not found: $resolvedInput"
    }
    $extension = [System.IO.Path]::GetExtension($resolvedInput).ToLowerInvariant()
    if (@('.tlb', '.olb', '.dll') -notcontains $extension) {
        throw "Only .tlb, .olb and .dll are supported: $resolvedInput"
    }

    Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop

    $result = [VbaCoder.TypeLibraryReader]::Read($resolvedInput)
    $types = @($result.Types)

    $document = [ordered]@{
        formatVersion = 1
        library = [ordered]@{
            name    = [string]$result.LibraryName
            guid    = [string]$result.LibraryGuid
            major   = [int]$result.Major
            minor   = [int]$result.Minor
            lcid    = [int]$result.Lcid
            bitness = [int]$Bitness
        }
        types = $types
    }

    $json = $document | ConvertTo-Json -Depth 32
    $directory = [System.IO.Path]::GetDirectoryName($OutputPath)
    if ($directory -and -not [System.IO.Directory]::Exists($directory)) {
        [void][System.IO.Directory]::CreateDirectory($directory)
    }
    $encoding = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($OutputPath, $json + [Environment]::NewLine, $encoding)

    Write-Output ("Extracted {0} {1}.{2}: {3} public types." -f $result.LibraryName, $result.Major, $result.Minor, $types.Count)
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
