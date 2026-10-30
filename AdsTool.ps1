<#
.SYNOPSIS
    Mini TwinCAT ADS test tool: search PLC variables, then read and write them.

.DESCRIPTION
    Needs only the TwinCAT ADS runtime on the PC (the AMS router + TcAdsDll.dll).
    No XAE, no Python, no extra .NET DLLs: it calls TcAdsDll.dll directly, so it
    does not care which TwinCAT.Ads.dll version is (or isn't) installed.

    Symbols are uploaded from the target and struct / FB members and small arrays
    are expanded so you can search for e.g. "motor speed" and find
    MAIN.fbMotor.stStatus.rSpeed.  Any path you type into the Name box is also
    resolved directly by the PLC, so array elements outside the expansion limits
    (MAIN.aBig[4711].x) still work.

    Value formats accepted when writing:
      BOOL      TRUE / FALSE / 1 / 0
      integers  42, -7, 16#FF, 0xFF, 2#1010
      REAL      1.5  (comma also accepted)
      TIME      T#1s500ms, T#2m, or a raw number of ms   (LTIME: raw = ns)
      DATE/DT   D#2026-10-29, DT#2026-10-29-11:34:00
      TOD       TOD#12:30:00.250
      STRING    'text' or text
      ENUM      Running, E_State.Running or the number
      anything  hex:01 02 0A FF  (exactly the variable's size)

.PARAMETER NetId       AMS Net ID or route name of the target. Blank = this PC.
                       The GUI also lists the routes already configured on this PC.
.PARAMETER Port        AMS port. TC3 PLC 1 = 851, TC2 PLC 1 = 801.
.PARAMETER Cli         Run in the console instead of the GUI.
.PARAMETER AdsDll      Full path to TcAdsDll.dll if it isn't found automatically.
.PARAMETER TimeoutMs   ADS timeout in milliseconds.
.PARAMETER MaxDepth    How deep struct / FB members are expanded in the search list.
.PARAMETER MaxArray    Arrays with more elements than this are not expanded in the list.
.PARAMETER Theme       Auto (follow the Windows app setting), Dark or Light.

.EXAMPLE
    .\AdsTool.ps1
.EXAMPLE
    .\AdsTool.ps1 -NetId 5.12.34.56.1.1 -Port 851
.EXAMPLE
    .\AdsTool.ps1 -Cli
#>
[CmdletBinding()]
param(
    [string]$NetId = '',
    [int]$Port = 851,
    [switch]$Cli,
    [string]$AdsDll = '',
    [int]$TimeoutMs = 2000,
    [int]$MaxDepth = 6,
    [int]$MaxArray = 100,
    [ValidateSet('Auto', 'Dark', 'Light')]
    [string]$Theme = 'Auto'
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# ADS interop and symbol handling (C#, compiled on the fly by Add-Type).
# Kept to C# 3 / .NET 3.5 features so it also compiles on old Windows PowerShell.
# ---------------------------------------------------------------------------
$source = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace AdsMini
{
    internal static class Native
    {
        const string Dll = "TcAdsDll.dll";
        [DllImport(Dll)] public static extern int AdsPortOpenEx();
        [DllImport(Dll)] public static extern int AdsPortCloseEx(int port);
        [DllImport(Dll)] public static extern int AdsGetLocalAddressEx(int port, byte[] addr);
        [DllImport(Dll)] public static extern int AdsSyncSetTimeoutEx(int port, int ms);
        [DllImport(Dll)] public static extern int AdsSyncReadReqEx2(int port, byte[] addr, uint group, uint offset, uint length, byte[] data, out uint bytesRead);
        [DllImport(Dll)] public static extern int AdsSyncWriteReqEx(int port, byte[] addr, uint group, uint offset, uint length, byte[] data);
        [DllImport(Dll)] public static extern int AdsSyncReadWriteReqEx2(int port, byte[] addr, uint group, uint offset, uint readLength, byte[] readData, uint writeLength, byte[] writeData, out uint bytesRead);
        [DllImport(Dll)] public static extern int AdsSyncReadStateReqEx(int port, byte[] addr, out ushort adsState, out ushort deviceState);
        [DllImport(Dll, CharSet = CharSet.Ansi)] public static extern int AdsSyncReadDeviceInfoReqEx(int port, byte[] addr, StringBuilder name, byte[] version);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] public static extern IntPtr LoadLibrary(string path);
    }

    public class AdsException : Exception
    {
        public int Code;
        public AdsException(int code, string what) : base(Describe(code, what)) { Code = code; }

        static string Describe(int code, string what)
        {
            string hint;
            switch (code)
            {
                case 0x6: hint = "target port not found. Is the PLC running, and is the AMS port right (TC3 = 851, TC2 = 801)?"; break;
                case 0x7: hint = "target machine not found. Check the AMS Net ID and that an ADS route to it exists."; break;
                case 0x701: hint = "service not supported by the target"; break;
                case 0x702: hint = "invalid index group"; break;
                case 0x703: hint = "invalid index offset"; break;
                case 0x704: hint = "reading/writing not permitted"; break;
                case 0x705: hint = "parameter size not correct"; break;
                case 0x706: hint = "invalid data value"; break;
                case 0x707: hint = "device not ready"; break;
                case 0x708: hint = "device busy"; break;
                case 0x710: hint = "symbol not found"; break;
                case 0x711: hint = "symbol version invalid: the PLC program changed, reload the symbols"; break;
                case 0x712: hint = "device in invalid state"; break;
                case 0x745: hint = "timeout: the target did not answer"; break;
                case 0x748: hint = "ADS port not open"; break;
                default: hint = ""; break;
            }
            return string.Format("{0} failed: ADS error {1} (0x{1:X}){2}", what, code, hint.Length > 0 ? ": " + hint : "");
        }
    }

    public class SymbolInfo
    {
        // Name / Type / Comment / Size are properties (not fields) so WPF lists can bind to them.
        string name = "", type = "", comment = "";
        uint size;
        public string Name { get { return name; } set { name = value ?? ""; } }
        public string Type { get { return type; } set { type = value ?? ""; } }
        public string Comment { get { return comment; } set { comment = value ?? ""; } }
        public uint Size { get { return size; } set { size = value; } }
        public uint IGroup, IOffset, DataType, Flags, Handle;
        public bool IsMember;
        public SymbolInfo Clone() { return (SymbolInfo)MemberwiseClone(); }
        public string Address
        {
            get
            {
                if (Handle != 0) return "by handle 0x" + Handle.ToString("X");
                return string.Format("group 0x{0:X}  offset 0x{1:X}", IGroup, IOffset);
            }
        }
    }

    public class ArrDim { public int LBound; public uint Elements; }

    // A row in the watch list; raises change notifications so the WPF list updates in place.
    public class WatchItem : System.ComponentModel.INotifyPropertyChanged
    {
        string name = "", val = "", type = "";
        bool isError;
        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
        public WatchItem(string name) { this.name = name ?? ""; }
        void Raise(string p)
        {
            System.ComponentModel.PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new System.ComponentModel.PropertyChangedEventArgs(p));
        }
        public string Name { get { return name; } }
        public string Value { get { return val; } set { if (val != value) { val = value ?? ""; Raise("Value"); } } }
        public string Type { get { return type; } set { if (type != value) { type = value ?? ""; Raise("Type"); } } }
        public bool IsError { get { return isError; } set { if (isError != value) { isError = value; Raise("IsError"); } } }
    }

    public class TypeInfo
    {
        public string Name = "", BaseType = "", Comment = "";
        public uint Size, Offs, DataType, Flags;
        public List<ArrDim> Dims = new List<ArrDim>();
        public List<TypeInfo> SubItems = new List<TypeInfo>();
        public Dictionary<long, string> Enum;
    }

    public class ReadResult
    {
        public SymbolInfo Info;
        public byte[] Raw;
        public string Value;
    }

    public class AdsClient : IDisposable
    {
        // ADS data type ids (ADST_*)
        const uint T_INT16 = 2, T_INT32 = 3, T_REAL32 = 4, T_REAL64 = 5, T_INT8 = 16, T_UINT8 = 17, T_UINT16 = 18,
                   T_UINT32 = 19, T_INT64 = 20, T_UINT64 = 21, T_STRING = 30, T_WSTRING = 31, T_BIT = 33, T_BIG = 65;

        public static readonly Encoding Ansi = GetAnsi();
        static readonly DateTime Epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        static readonly CultureInfo IC = CultureInfo.InvariantCulture;
        public static string DllPath = "";

        int port;
        byte[] addr = new byte[8];
        Dictionary<string, TypeInfo> types = new Dictionary<string, TypeInfo>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, SymbolInfo> byName = new Dictionary<string, SymbolInfo>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, SymbolInfo> resolved = new Dictionary<string, SymbolInfo>(StringComparer.OrdinalIgnoreCase);

        public List<SymbolInfo> Symbols = new List<SymbolInfo>();
        public int TopLevelCount;
        public int MaxDepth = 6, MaxArray = 100, MaxTotal = 500000;
        public bool Truncated;
        public string Target = "";

        public bool Connected { get { return port != 0; } }
        public int TypeCount { get { return types.Count; } }

        static Encoding GetAnsi()
        {
            try { return Encoding.GetEncoding(1252); } catch { }
            try { return Encoding.GetEncoding(28591); } catch { }
            return Encoding.ASCII;
        }

        // ---------------------------------------------------------------- DLL / connection

        public static string LoadDll(string explicitPath)
        {
            if (DllPath.Length > 0) return DllPath;
            string bits = IntPtr.Size == 8 ? "Common64" : "Common32";
            List<string> c = new List<string>();
            if (!string.IsNullOrEmpty(explicitPath)) c.Add(explicitPath);
            string tc3 = Environment.GetEnvironmentVariable("TWINCAT3DIR");
            if (!string.IsNullOrEmpty(tc3))
            {
                try { c.Add(Path.Combine(Path.Combine(Path.GetFullPath(Path.Combine(tc3, "..")), bits), "TcAdsDll.dll")); } catch { }
            }
            c.Add(Path.Combine(Path.Combine(@"C:\TwinCAT", bits), "TcAdsDll.dll"));
            foreach (string pf in new string[] { Environment.GetEnvironmentVariable("ProgramFiles(x86)"), Environment.GetEnvironmentVariable("ProgramFiles") })
            {
                if (!string.IsNullOrEmpty(pf))
                    c.Add(Path.Combine(Path.Combine(Path.Combine(Path.Combine(pf, "Beckhoff"), "TwinCAT"), bits), "TcAdsDll.dll"));
            }
            c.Add("TcAdsDll.dll"); // PATH / System32 (TwinCAT 2)

            List<string> tried = new List<string>();
            foreach (string p in c)
            {
                if (p != "TcAdsDll.dll" && !File.Exists(p)) continue;
                if (Native.LoadLibrary(p) != IntPtr.Zero) { DllPath = p; return p; }
                int err = Marshal.GetLastWin32Error();
                tried.Add(p + (err == 193 ? " (wrong 32/64-bit)" : " (error " + err + ")"));
            }
            throw new Exception("Could not load TcAdsDll.dll for this " + (IntPtr.Size == 8 ? "64" : "32") +
                "-bit PowerShell. Is TwinCAT ADS installed? You can point at it with -AdsDll <path>." +
                (tried.Count > 0 ? " Tried: " + string.Join("; ", tried.ToArray()) : ""));
        }

        public void Connect(string netId, int amsPort, int timeoutMs)
        {
            LoadDll(null);
            Disconnect();
            int p = Native.AdsPortOpenEx();
            if (p == 0) throw new Exception("Could not open an ADS port. Is the TwinCAT ADS router / System Service running?");
            port = p;
            try
            {
                byte[] a = new byte[8];
                string id = (netId ?? "").Trim();
                if (id.Length == 0 || id.ToLowerInvariant() == "local")
                    Check(Native.AdsGetLocalAddressEx(port, a), "Get local AMS Net ID");
                else
                    Array.Copy(ParseNetId(id), a, 6);
                a[6] = (byte)(amsPort & 0xFF);
                a[7] = (byte)((amsPort >> 8) & 0xFF);
                addr = a;
                Native.AdsSyncSetTimeoutEx(port, timeoutMs);
                Target = NetIdToString(a) + ":" + amsPort;
                ReadState(); // probe the target so a bad address fails here
            }
            catch
            {
                Disconnect();
                throw;
            }
        }

        public void Disconnect()
        {
            if (port != 0)
            {
                foreach (SymbolInfo s in resolved.Values) ReleaseHandle(s);
                try { Native.AdsPortCloseEx(port); } catch { }
            }
            port = 0;
            resolved.Clear();
        }

        public void Dispose() { Disconnect(); }

        public static string LocalNetId()
        {
            LoadDll(null);
            int p = Native.AdsPortOpenEx();
            if (p == 0) return "";
            try
            {
                byte[] a = new byte[8];
                return Native.AdsGetLocalAddressEx(p, a) == 0 ? NetIdToString(a) : "";
            }
            finally { Native.AdsPortCloseEx(p); }
        }

        public static byte[] ParseNetId(string s)
        {
            string[] parts = s.Trim().Split('.');
            if (parts.Length != 6) throw new FormatException("AMS Net ID must have 6 parts, e.g. 5.12.34.56.1.1");
            byte[] b = new byte[6];
            for (int i = 0; i < 6; i++)
            {
                if (!byte.TryParse(parts[i], NumberStyles.None, IC, out b[i]))
                    throw new FormatException("Bad AMS Net ID part '" + parts[i] + "'");
            }
            return b;
        }

        static string NetIdToString(byte[] a)
        {
            return string.Format("{0}.{1}.{2}.{3}.{4}.{5}", a[0], a[1], a[2], a[3], a[4], a[5]);
        }

        void Check(int err, string what) { if (err != 0) throw new AdsException(err, what); }

        void EnsureConnected() { if (port == 0) throw new InvalidOperationException("Not connected"); }

        // ---------------------------------------------------------------- raw ADS

        public byte[] Read(uint group, uint offset, uint length)
        {
            EnsureConnected();
            byte[] buf = new byte[length];
            uint got;
            Check(Native.AdsSyncReadReqEx2(port, addr, group, offset, length, buf, out got), "Read");
            if (got < length) Array.Resize(ref buf, (int)got);
            return buf;
        }

        public void Write(uint group, uint offset, byte[] data)
        {
            EnsureConnected();
            Check(Native.AdsSyncWriteReqEx(port, addr, group, offset, (uint)data.Length, data), "Write");
        }

        public byte[] ReadWrite(uint group, uint offset, uint readLength, byte[] writeData)
        {
            EnsureConnected();
            byte[] buf = new byte[readLength];
            uint got;
            Check(Native.AdsSyncReadWriteReqEx2(port, addr, group, offset, readLength, buf, (uint)writeData.Length, writeData, out got), "ReadWrite");
            if (got < readLength) Array.Resize(ref buf, (int)got);
            return buf;
        }

        public string ReadState()
        {
            EnsureConnected();
            ushort s, d;
            Check(Native.AdsSyncReadStateReqEx(port, addr, out s, out d), "Read state");
            string[] names = { "Invalid", "Idle", "Reset", "Init", "Start", "Run", "Stop", "SaveCfg", "LoadCfg", "PowerFailure",
                               "PowerGood", "Error", "Shutdown", "Suspend", "Resume", "Config", "Reconfig", "Stopping", "Incompatible", "Exception" };
            return s < names.Length ? names[s] : "State " + s;
        }

        public string DeviceInfo()
        {
            EnsureConnected();
            StringBuilder sb = new StringBuilder(64);
            byte[] v = new byte[4];
            Check(Native.AdsSyncReadDeviceInfoReqEx(port, addr, sb, v), "Read device info");
            return string.Format("{0} {1}.{2}.{3}", sb.ToString().Trim(), v[0], v[1], BitConverter.ToUInt16(v, 2));
        }

        // ---------------------------------------------------------------- symbol upload

        public int LoadSymbols()
        {
            EnsureConnected();
            uint symLen, dtLen = 0;
            try
            {
                byte[] info = Read(0xF00F, 0, 24); // ADSIGRP_SYM_UPLOADINFO2
                symLen = U32(info, 4);
                dtLen = U32(info, 12);
            }
            catch (AdsException)
            {
                byte[] info = Read(0xF00C, 0, 8);  // ADSIGRP_SYM_UPLOADINFO (old runtimes)
                symLen = U32(info, 4);
            }
            byte[] symBuf = symLen > 0 ? Read(0xF00B, 0, symLen) : new byte[0];
            byte[] dtBuf = new byte[0];
            if (dtLen > 0)
            {
                try { dtBuf = Read(0xF00E, 0, dtLen); } catch (AdsException) { }
            }
            BuildFromUpload(symBuf, dtBuf);
            return Symbols.Count;
        }

        public void BuildFromUpload(byte[] symBuf, byte[] dtBuf)
        {
            types = ParseDatatypes(dtBuf);
            List<SymbolInfo> top = ParseSymbols(symBuf);
            top.Sort(delegate(SymbolInfo a, SymbolInfo b) { return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase); });
            TopLevelCount = top.Count;
            Symbols = new List<SymbolInfo>(top.Count * 4);
            byName = new Dictionary<string, SymbolInfo>(StringComparer.OrdinalIgnoreCase);
            resolved.Clear();
            Truncated = false;
            foreach (SymbolInfo s in top)
            {
                if (!Add(s)) break;
                TypeInfo t;
                if (types.TryGetValue(s.Type, out t)) Expand(s, t, 1);
            }
        }

        bool Add(SymbolInfo s)
        {
            if (Symbols.Count >= MaxTotal) { Truncated = true; return false; }
            Symbols.Add(s);
            if (!byName.ContainsKey(s.Name)) byName[s.Name] = s;
            return true;
        }

        TypeInfo FollowAlias(TypeInfo t)
        {
            for (int i = 0; i < 16 && t != null; i++)
            {
                if (t.SubItems.Count > 0 || t.Dims.Count > 0 || t.Enum != null) return t;
                if (string.IsNullOrEmpty(t.BaseType) || string.Equals(t.BaseType, t.Name, StringComparison.OrdinalIgnoreCase)) return t;
                TypeInfo n;
                if (!types.TryGetValue(t.BaseType, out n) || n == t) return t;
                t = n;
            }
            return t;
        }

        static bool IsPointerLike(string typeName, uint flags)
        {
            string u = (typeName ?? "").TrimStart().ToUpperInvariant();
            return u.StartsWith("POINTER TO") || u.StartsWith("REFERENCE TO") || (flags & 0x4) != 0;
        }

        void Expand(SymbolInfo parent, TypeInfo t, int depth)
        {
            if (depth > MaxDepth || Truncated) return;
            t = FollowAlias(t);
            if (t == null) return;
            if (t.Dims.Count > 0) { ExpandArray(parent, t, depth); return; }
            foreach (TypeInfo sub in t.SubItems)
            {
                if ((sub.Flags & 0x40) != 0) continue; // property item, can't be read as data
                SymbolInfo c = new SymbolInfo();
                c.Name = parent.Name + "." + sub.Name;
                c.Type = sub.BaseType;
                c.Size = sub.Size;
                c.DataType = sub.DataType;
                c.Comment = sub.Comment;
                c.IGroup = parent.IGroup;
                c.IOffset = parent.IOffset + sub.Offs;
                c.IsMember = true;
                if (!Add(c)) return;
                if (!IsPointerLike(sub.BaseType, sub.Flags)) Expand(c, sub, depth + 1);
            }
        }

        static string ElementType(TypeInfo t)
        {
            string b = t.BaseType ?? "";
            if (b.Length > 0 && !b.TrimStart().ToUpperInvariant().StartsWith("ARRAY")) return b;
            string src = b.Length > 0 ? b : t.Name;
            int i = src.IndexOf(" OF ", StringComparison.OrdinalIgnoreCase);
            return i >= 0 ? src.Substring(i + 4).Trim() : b;
        }

        void ExpandArray(SymbolInfo parent, TypeInfo t, int depth)
        {
            long count = 1;
            foreach (ArrDim d in t.Dims) count *= d.Elements;
            if (count <= 0 || count > MaxArray) return;
            string et = ElementType(t);
            uint esz = (uint)(t.Size / count);
            TypeInfo eti;
            types.TryGetValue(et, out eti);
            uint edt = eti != null ? eti.DataType : GuessDataType(et.ToUpperInvariant());
            string[] idx = new string[t.Dims.Count];
            for (long k = 0; k < count; k++)
            {
                long r = k;
                for (int i = t.Dims.Count - 1; i >= 0; i--)
                {
                    idx[i] = ((r % t.Dims[i].Elements) + t.Dims[i].LBound).ToString(IC);
                    r /= t.Dims[i].Elements;
                }
                SymbolInfo c = new SymbolInfo();
                c.Name = parent.Name + "[" + string.Join(",", idx) + "]";
                c.Type = et;
                c.Size = esz;
                c.DataType = edt;
                c.IGroup = parent.IGroup;
                c.IOffset = parent.IOffset + (uint)(k * esz);
                c.IsMember = true;
                if (!Add(c)) return;
                if (eti != null && !IsPointerLike(et, 0)) Expand(c, eti, depth + 1);
            }
        }

        // ---------------------------------------------------------------- binary parsing

        static uint U32(byte[] b, int p) { return BitConverter.ToUInt32(b, p); }
        static ushort U16(byte[] b, int p) { return BitConverter.ToUInt16(b, p); }

        static string Str(byte[] b, int p, int len)
        {
            if (len <= 0 || p >= b.Length) return "";
            return Ansi.GetString(b, p, Math.Min(len, b.Length - p));
        }

        static long ReadSigned(byte[] b, int p, int n)
        {
            switch (n)
            {
                case 1: return (sbyte)b[p];
                case 2: return BitConverter.ToInt16(b, p);
                case 4: return BitConverter.ToInt32(b, p);
                case 8: return BitConverter.ToInt64(b, p);
            }
            long v = 0;
            for (int i = Math.Min(n, 8) - 1; i >= 0; i--) v = (v << 8) | b[p + i];
            return v;
        }

        static byte[] Bytes(long v, uint size)
        {
            byte[] all = BitConverter.GetBytes(v);
            byte[] r = new byte[size];
            Array.Copy(all, r, Math.Min(8, (int)size));
            return r;
        }

        public static List<SymbolInfo> ParseSymbols(byte[] b)
        {
            List<SymbolInfo> list = new List<SymbolInfo>();
            int p = 0;
            while (p + 30 <= b.Length)
            {
                uint len = U32(b, p);
                if (len < 30 || p + len > b.Length) break;
                list.Add(ParseSymbolEntry(b, p));
                p += (int)len;
            }
            return list;
        }

        static SymbolInfo ParseSymbolEntry(byte[] b, int p)
        {
            SymbolInfo s = new SymbolInfo();
            s.IGroup = U32(b, p + 4);
            s.IOffset = U32(b, p + 8);
            s.Size = U32(b, p + 12);
            s.DataType = U32(b, p + 16);
            s.Flags = U32(b, p + 20);
            int nl = U16(b, p + 24), tl = U16(b, p + 26), cl = U16(b, p + 28);
            int q = p + 30;
            s.Name = Str(b, q, nl); q += nl + 1;
            s.Type = Str(b, q, tl); q += tl + 1;
            s.Comment = Str(b, q, cl).Trim();
            return s;
        }

        public static Dictionary<string, TypeInfo> ParseDatatypes(byte[] b)
        {
            Dictionary<string, TypeInfo> d = new Dictionary<string, TypeInfo>(StringComparer.OrdinalIgnoreCase);
            int p = 0;
            while (p + 42 <= b.Length)
            {
                uint len = U32(b, p);
                if (len < 42 || p + len > b.Length) break;
                TypeInfo t = ParseDatatype(b, p, p + (int)len);
                if (t.Name.Length > 0) d[t.Name] = t;
                p += (int)len;
            }
            return d;
        }

        static TypeInfo ParseDatatype(byte[] b, int p, int end)
        {
            TypeInfo t = new TypeInfo();
            t.Size = U32(b, p + 16);
            t.Offs = U32(b, p + 20);
            t.DataType = U32(b, p + 24);
            t.Flags = U32(b, p + 28);
            int nl = U16(b, p + 32), tl = U16(b, p + 34), cl = U16(b, p + 36), ad = U16(b, p + 38), si = U16(b, p + 40);
            int q = p + 42;
            t.Name = Str(b, q, nl); q += nl + 1;
            t.BaseType = Str(b, q, tl); q += tl + 1;
            t.Comment = Str(b, q, cl).Trim(); q += cl + 1;
            for (int i = 0; i < ad && q + 8 <= end; i++)
            {
                ArrDim a = new ArrDim();
                a.LBound = BitConverter.ToInt32(b, q);
                a.Elements = U32(b, q + 4);
                t.Dims.Add(a);
                q += 8;
            }
            for (int i = 0; i < si && q + 42 <= end; i++)
            {
                uint l = U32(b, q);
                if (l < 42 || q + l > end) break;
                t.SubItems.Add(ParseDatatype(b, q, q + (int)l));
                q += (int)l;
            }
            // Optional trailing blocks. Only the enum values are used; anything odd just loses them.
            try
            {
                if ((t.Flags & 0x80) != 0) q += 16;                  // type GUID
                if ((t.Flags & 0x200) != 0) q += (int)t.Size;        // copy mask
                if ((t.Flags & 0x800) != 0)                          // method infos
                {
                    int n = U16(b, q); q += 2;
                    for (int i = 0; i < n; i++) q += (int)U32(b, q);
                }
                if ((t.Flags & 0x1000) != 0)                         // attributes
                {
                    int n = U16(b, q); q += 2;
                    for (int i = 0; i < n; i++) q += 2 + b[q] + 1 + b[q + 1] + 1;
                }
                if ((t.Flags & 0x2000) != 0 && t.Size > 0 && t.Size <= 8) // enum infos
                {
                    int n = U16(b, q); q += 2;
                    Dictionary<long, string> e = new Dictionary<long, string>();
                    for (int i = 0; i < n; i++)
                    {
                        int l = b[q]; q += 1;
                        string name = Str(b, q, l); q += l + 1;
                        if (q + (int)t.Size > end) throw new IndexOutOfRangeException();
                        long v = ReadSigned(b, q, (int)t.Size); q += (int)t.Size;
                        if (name.Length > 0 && !e.ContainsKey(v)) e[v] = name;
                    }
                    if (e.Count > 0) t.Enum = e;
                }
            }
            catch { t.Enum = null; }
            return t;
        }

        // ---------------------------------------------------------------- search

        public List<SymbolInfo> Search(string query, bool includeMembers)
        {
            // Every word must match the name or the type. A word with * or ? is a wildcard
            // pattern for the whole name or type (e.g. MAIN.* or ARRAY*); other words match anywhere.
            string[] terms = (query ?? "").Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            Regex[] rx = new Regex[terms.Length];
            for (int i = 0; i < terms.Length; i++)
            {
                if (terms[i].IndexOf('*') >= 0 || terms[i].IndexOf('?') >= 0)
                    rx[i] = new Regex("^" + Regex.Escape(terms[i]).Replace(@"\*", ".*").Replace(@"\?", ".") + "$", RegexOptions.IgnoreCase);
            }
            List<SymbolInfo> r = new List<SymbolInfo>();
            foreach (SymbolInfo s in Symbols)
            {
                if (!includeMembers && s.IsMember) continue;
                bool ok = true;
                for (int i = 0; i < terms.Length && ok; i++)
                {
                    if (rx[i] != null) ok = rx[i].IsMatch(s.Name) || rx[i].IsMatch(s.Type);
                    else ok = s.Name.IndexOf(terms[i], StringComparison.OrdinalIgnoreCase) >= 0
                           || s.Type.IndexOf(terms[i], StringComparison.OrdinalIgnoreCase) >= 0;
                }
                if (ok) r.Add(s);
            }
            return r;
        }

        // ---------------------------------------------------------------- read / write by name

        public SymbolInfo Resolve(string name)
        {
            name = (name ?? "").Trim();
            if (name.Length == 0) throw new ArgumentException("No variable name given");
            SymbolInfo s;
            if (resolved.TryGetValue(name, out s)) return s;
            SymbolInfo cached;
            byName.TryGetValue(name, out cached);
            try
            {
                byte[] r = ReadWrite(0xF009, 0, 0xFFFF, Ansi.GetBytes(name + "\0")); // ADSIGRP_SYM_INFOBYNAMEEX
                s = ParseSymbolEntry(r, 0);
                s.Name = name;
                if (cached != null)
                {
                    if (s.Comment.Length == 0) s.Comment = cached.Comment;
                    s.IsMember = cached.IsMember;
                }
            }
            catch (AdsException)
            {
                if (cached == null) throw;
                // Fall back to a handle plus the type info from the upload.
                s = cached.Clone();
                byte[] h = ReadWrite(0xF003, 0, 4, Ansi.GetBytes(name + "\0")); // ADSIGRP_SYM_HNDBYNAME
                s.Handle = U32(h, 0);
            }
            resolved[name] = s;
            return s;
        }

        void ReleaseHandle(SymbolInfo s)
        {
            if (s.Handle == 0 || port == 0) return;
            try { Write(0xF006, 0, BitConverter.GetBytes(s.Handle)); } catch { }
            s.Handle = 0;
        }

        void Forget(string name)
        {
            SymbolInfo s;
            if (resolved.TryGetValue(name.Trim(), out s)) { ReleaseHandle(s); resolved.Remove(name.Trim()); }
        }

        byte[] ReadRaw(SymbolInfo s)
        {
            return s.Handle != 0 ? Read(0xF005, s.Handle, s.Size) : Read(s.IGroup, s.IOffset, s.Size);
        }

        void WriteRaw(SymbolInfo s, byte[] d)
        {
            if (s.Handle != 0) Write(0xF005, s.Handle, d); else Write(s.IGroup, s.IOffset, d);
        }

        public ReadResult ReadValue(string name)
        {
            SymbolInfo s = Resolve(name);
            byte[] d;
            try { d = ReadRaw(s); }
            catch (AdsException)
            {
                // Program may have been re-downloaded: resolve again once.
                Forget(name);
                s = Resolve(name);
                d = ReadRaw(s);
            }
            ReadResult r = new ReadResult();
            r.Info = s;
            r.Raw = d;
            r.Value = Format(s, d);
            return r;
        }

        public ReadResult WriteValue(string name, string text)
        {
            SymbolInfo s = Resolve(name);
            byte[] d = Parse(s, text);
            try { WriteRaw(s, d); }
            catch (AdsException ex)
            {
                if (ex.Code == 0x704 || ex.Code == 0x706) throw;
                Forget(name);
                s = Resolve(name);
                WriteRaw(s, Parse(s, text));
            }
            return ReadValue(name);
        }

        // Numeric view of a value for charting; NaN when it isn't a number (strings, structs).
        public double NumericValue(SymbolInfo s, byte[] d)
        {
            if (d == null || d.Length == 0) return double.NaN;
            TypeInfo en;
            string bn = BaseName(s.Type, out en);
            if (en != null && d.Length <= 8) return ReadSigned(d, 0, d.Length);
            if (bn == "LTIME" && d.Length >= 8) return BitConverter.ToUInt64(d, 0) / 1e6; // ms
            uint dt = Effective(s, bn);
            if (NeededSize(dt) == 0 || d.Length < NeededSize(dt)) return double.NaN;
            switch (dt)
            {
                case T_BIT: return d[0] != 0 ? 1 : 0;
                case T_INT8: return (sbyte)d[0];
                case T_UINT8: return d[0];
                case T_INT16: return BitConverter.ToInt16(d, 0);
                case T_UINT16: return BitConverter.ToUInt16(d, 0);
                case T_INT32: return BitConverter.ToInt32(d, 0);
                case T_UINT32: return BitConverter.ToUInt32(d, 0);
                case T_INT64: return BitConverter.ToInt64(d, 0);
                case T_UINT64: return BitConverter.ToUInt64(d, 0);
                case T_REAL32: return BitConverter.ToSingle(d, 0);
                case T_REAL64: return BitConverter.ToDouble(d, 0);
            }
            return double.NaN;
        }

        public bool IsBoolType(SymbolInfo s)
        {
            TypeInfo en;
            string bn = BaseName(s.Type, out en);
            return en == null && Effective(s, bn) == T_BIT;
        }

        public SymbolInfo Find(string name)
        {
            SymbolInfo s;
            return byName.TryGetValue((name ?? "").Trim(), out s) ? s : null;
        }

        // ---------------------------------------------------------------- value formatting

        // Follows aliases down to the base type name (upper case); returns the enum type if one is met on the way.
        string BaseName(string typeName, out TypeInfo enumType)
        {
            enumType = null;
            string n = (typeName ?? "").Trim();
            for (int i = 0; i < 16; i++)
            {
                TypeInfo t;
                if (!types.TryGetValue(n, out t)) break;
                if (t.Enum != null) { enumType = t; break; }
                if (t.SubItems.Count > 0 || t.Dims.Count > 0) break;
                if (string.IsNullOrEmpty(t.BaseType) || string.Equals(t.BaseType, n, StringComparison.OrdinalIgnoreCase)) break;
                n = t.BaseType.Trim();
            }
            return n.ToUpperInvariant();
        }

        static uint GuessDataType(string bn)
        {
            switch (bn)
            {
                case "BOOL": case "BIT": return T_BIT;
                case "SINT": return T_INT8;
                case "USINT": case "BYTE": return T_UINT8;
                case "INT": return T_INT16;
                case "UINT": case "WORD": return T_UINT16;
                case "DINT": return T_INT32;
                case "UDINT": case "DWORD": case "TIME": case "TOD": case "TIME_OF_DAY": case "DATE": case "DT": case "DATE_AND_TIME": return T_UINT32;
                case "LINT": return T_INT64;
                case "ULINT": case "LWORD": case "LTIME": return T_UINT64;
                case "REAL": return T_REAL32;
                case "LREAL": return T_REAL64;
            }
            if (bn.StartsWith("WSTRING")) return T_WSTRING;
            if (bn.StartsWith("STRING")) return T_STRING;
            return 0;
        }

        static int NeededSize(uint dt)
        {
            switch (dt)
            {
                case T_BIT: case T_INT8: case T_UINT8: return 1;
                case T_INT16: case T_UINT16: return 2;
                case T_INT32: case T_UINT32: case T_REAL32: return 4;
                case T_INT64: case T_UINT64: case T_REAL64: return 8;
            }
            return 0;
        }

        static uint Effective(SymbolInfo s, string bn)
        {
            uint dt = s.DataType;
            if (dt == 0 || dt == T_BIG)
            {
                uint g = GuessDataType(bn);
                if (g != 0) dt = g;
            }
            return dt;
        }

        static bool IsBitString(string bn) { return bn == "BYTE" || bn == "WORD" || bn == "DWORD" || bn == "LWORD"; }

        static string FormatDuration(ulong ns, bool ltime)
        {
            string prefix = ltime ? "LTIME#" : "T#";
            if (ns == 0) return prefix + "0ms";
            ulong[] div = { 86400000000000UL, 3600000000000UL, 60000000000UL, 1000000000UL, 1000000UL, 1000UL, 1UL };
            string[] unit = { "d", "h", "m", "s", "ms", "us", "ns" };
            StringBuilder sb = new StringBuilder(prefix);
            for (int i = 0; i < div.Length; i++)
            {
                ulong q = ns / div[i];
                ns %= div[i];
                if (q > 0) sb.Append(q.ToString(IC)).Append(unit[i]);
            }
            return sb.ToString();
        }

        static string Hex(byte[] d)
        {
            int n = Math.Min(d.Length, 1024);
            string h = BitConverter.ToString(d, 0, n).Replace("-", " ");
            return n < d.Length ? h + " ... (" + d.Length + " bytes, showing 1024)" : h;
        }

        public string Format(SymbolInfo s, byte[] d)
        {
            if (d == null || d.Length == 0) return "";
            TypeInfo en;
            string bn = BaseName(s.Type, out en);
            if (en != null && d.Length <= 8)
            {
                long v = ReadSigned(d, 0, d.Length);
                string nm;
                return en.Enum.TryGetValue(v, out nm) ? nm + " (" + v.ToString(IC) + ")" : v.ToString(IC);
            }
            if (d.Length >= 4)
            {
                uint u = U32(d, 0);
                switch (bn)
                {
                    case "TIME": return FormatDuration((ulong)u * 1000000UL, false);
                    case "TOD":
                    case "TIME_OF_DAY":
                        return string.Format(IC, "TOD#{0:00}:{1:00}:{2:00}.{3:000}", u / 3600000, u / 60000 % 60, u / 1000 % 60, u % 1000);
                    case "DATE": return "D#" + Epoch.AddSeconds(u).ToString("yyyy-MM-dd", IC);
                    case "DT":
                    case "DATE_AND_TIME": return "DT#" + Epoch.AddSeconds(u).ToString("yyyy-MM-dd-HH:mm:ss", IC);
                }
            }
            if (bn == "LTIME" && d.Length >= 8) return FormatDuration(BitConverter.ToUInt64(d, 0), true);

            uint dt = Effective(s, bn);
            if (d.Length < NeededSize(dt)) return Hex(d);
            switch (dt)
            {
                case T_BIT: return d[0] != 0 ? "TRUE" : "FALSE";
                case T_INT8: return ((sbyte)d[0]).ToString(IC);
                case T_UINT8: return WithHex(bn, d[0], 2);
                case T_INT16: return BitConverter.ToInt16(d, 0).ToString(IC);
                case T_UINT16: return WithHex(bn, BitConverter.ToUInt16(d, 0), 4);
                case T_INT32: return BitConverter.ToInt32(d, 0).ToString(IC);
                case T_UINT32: return WithHex(bn, BitConverter.ToUInt32(d, 0), 8);
                case T_INT64: return BitConverter.ToInt64(d, 0).ToString(IC);
                case T_UINT64: return WithHex(bn, BitConverter.ToUInt64(d, 0), 16);
                case T_REAL32: return BitConverter.ToSingle(d, 0).ToString("R", IC);
                case T_REAL64: return BitConverter.ToDouble(d, 0).ToString("R", IC);
                case T_STRING:
                {
                    int z = Array.IndexOf(d, (byte)0);
                    return "'" + Ansi.GetString(d, 0, z < 0 ? d.Length : z) + "'";
                }
                case T_WSTRING:
                {
                    int z = d.Length & ~1;
                    for (int i = 0; i + 1 < d.Length; i += 2) if (d[i] == 0 && d[i + 1] == 0) { z = i; break; }
                    return "\"" + Encoding.Unicode.GetString(d, 0, z) + "\"";
                }
            }
            return Hex(d);
        }

        static string WithHex(string bn, ulong v, int digits)
        {
            string dec = v.ToString(IC);
            return IsBitString(bn) ? dec + " (16#" + v.ToString("X" + digits, IC) + ")" : dec;
        }

        // ---------------------------------------------------------------- value parsing

        public byte[] Parse(SymbolInfo s, string text)
        {
            string t = (text ?? "").Trim();
            uint size = s.Size;
            if (t.StartsWith("hex:", StringComparison.OrdinalIgnoreCase)) return ParseHex(t.Substring(4), size);

            TypeInfo en;
            string bn = BaseName(s.Type, out en);
            uint dt = Effective(s, bn);

            if (en == null && dt == T_STRING)
            {
                t = Unquote(t, '\'');
                byte[] sb = Ansi.GetBytes(t);
                if (sb.Length > (int)size - 1) throw new FormatException("String too long: max " + (size - 1) + " characters");
                byte[] r = new byte[size];
                Array.Copy(sb, r, sb.Length);
                return r;
            }
            if (en == null && dt == T_WSTRING)
            {
                t = Unquote(t, '"');
                byte[] wb = Encoding.Unicode.GetBytes(t);
                if (wb.Length > (int)size - 2) throw new FormatException("String too long: max " + ((size - 2) / 2) + " characters");
                byte[] r = new byte[size];
                Array.Copy(wb, r, wb.Length);
                return r;
            }

            // Allow pasting a displayed value such as "Running (2)" or "255 (16#FF)".
            int k = t.IndexOf(" (");
            if (k > 0) t = t.Substring(0, k).Trim();
            if (t.Length == 0) throw new FormatException("No value given");

            if (en != null)
            {
                string want = t.IndexOf('.') >= 0 ? t.Substring(t.LastIndexOf('.') + 1) : t;
                foreach (KeyValuePair<long, string> kv in en.Enum)
                    if (string.Equals(kv.Value, want, StringComparison.OrdinalIgnoreCase)) return Bytes(kv.Key, size);
                bool neg; ulong mag;
                ParseInteger(t, out neg, out mag);
                return Bytes(neg ? -(long)mag : (long)mag, size);
            }

            switch (bn)
            {
                case "TIME":
                {
                    ulong ms = ParseDuration(t, false);
                    if (ms > uint.MaxValue) throw new FormatException("TIME too large");
                    return Bytes((long)ms, size);
                }
                case "LTIME": return Bytes(unchecked((long)ParseDuration(t, true)), size);
                case "TOD":
                case "TIME_OF_DAY": return Bytes(ParseTod(t), size);
                case "DATE":
                case "DT":
                case "DATE_AND_TIME": return Bytes(ParseDate(t), size);
            }

            switch (dt)
            {
                case T_BIT:
                {
                    string u = t.ToUpperInvariant();
                    byte v;
                    if (u == "TRUE" || u == "1" || u == "ON") v = 1;
                    else if (u == "FALSE" || u == "0" || u == "OFF") v = 0;
                    else throw new FormatException("BOOL must be TRUE or FALSE");
                    byte[] r = new byte[Math.Max(1u, size)];
                    r[0] = v;
                    return r;
                }
                case T_INT8: case T_INT16: case T_INT32: case T_INT64:
                case T_UINT8: case T_UINT16: case T_UINT32: case T_UINT64:
                    return IntBytes(t, dt, size);
                case T_REAL32:
                    return BitConverter.GetBytes(float.Parse(t.Replace(',', '.'), NumberStyles.Float, IC));
                case T_REAL64:
                    return BitConverter.GetBytes(double.Parse(t.Replace(',', '.'), NumberStyles.Float, IC));
            }
            throw new FormatException("Can't write " + s.Type + " from text. Use hex:01 02 ... with exactly " + size + " bytes.");
        }

        static string Unquote(string t, char q)
        {
            if (t.Length >= 2 && t[0] == q && t[t.Length - 1] == q) return t.Substring(1, t.Length - 2);
            return t;
        }

        static byte[] ParseHex(string h, uint size)
        {
            string c = Regex.Replace(h, @"[\s\-,]|0x", "", RegexOptions.IgnoreCase);
            if (c.Length % 2 != 0 || !Regex.IsMatch(c, "^[0-9A-Fa-f]*$")) throw new FormatException("Bad hex data");
            if (c.Length / 2 != size) throw new FormatException("Hex data must be exactly " + size + " bytes (got " + (c.Length / 2) + ")");
            byte[] r = new byte[size];
            for (int i = 0; i < r.Length; i++) r[i] = byte.Parse(c.Substring(i * 2, 2), NumberStyles.HexNumber, IC);
            return r;
        }

        static void ParseInteger(string t, out bool neg, out ulong mag)
        {
            string s = t.Replace("_", "").Trim();
            neg = false;
            if (s.StartsWith("-")) { neg = true; s = s.Substring(1); }
            else if (s.StartsWith("+")) s = s.Substring(1);
            int b = 10;
            string u = s.ToUpperInvariant();
            if (u.StartsWith("16#")) { b = 16; s = s.Substring(3); }
            else if (u.StartsWith("0X")) { b = 16; s = s.Substring(2); }
            else if (u.StartsWith("8#")) { b = 8; s = s.Substring(2); }
            else if (u.StartsWith("2#")) { b = 2; s = s.Substring(2); }
            if (s.Length == 0) throw new FormatException("Not a number: " + t);
            try
            {
                mag = b == 10 ? ulong.Parse(s, NumberStyles.None, IC) : Convert.ToUInt64(s, b);
            }
            catch (OverflowException) { throw new FormatException("Number too large: " + t); }
            catch (Exception) { throw new FormatException("Not a whole number: " + t); }
        }

        static byte[] IntBytes(string t, uint dt, uint size)
        {
            bool neg; ulong mag;
            ParseInteger(t, out neg, out mag);
            bool signed = dt == T_INT8 || dt == T_INT16 || dt == T_INT32 || dt == T_INT64;
            int bits = NeededSize(dt) * 8;
            if (signed)
            {
                ulong max = bits == 64 ? (ulong)long.MaxValue : (1UL << (bits - 1)) - 1;
                if (mag > (neg ? max + 1 : max)) throw new FormatException(t + " is out of range for a " + bits + "-bit signed value");
                long v = unchecked(neg ? -(long)mag : (long)mag);
                return Bytes(v, size);
            }
            if (neg && mag != 0) throw new FormatException("Negative value not allowed for an unsigned type");
            ulong maxU = bits == 64 ? ulong.MaxValue : (1UL << bits) - 1;
            if (mag > maxU) throw new FormatException(t + " is out of range for a " + bits + "-bit unsigned value");
            return Bytes(unchecked((long)mag), size);
        }

        static string AfterHash(string t, out bool hadPrefix)
        {
            int h = t.IndexOf('#');
            hadPrefix = h >= 0;
            return (h >= 0 ? t.Substring(h + 1) : t).Trim();
        }

        static ulong ParseDuration(string t, bool ltime)
        {
            bool prefixed;
            string s = AfterHash(t, out prefixed).Replace("_", "");
            ulong raw;
            if (!prefixed && ulong.TryParse(s, NumberStyles.None, IC, out raw)) return raw; // raw value: ms for TIME, ns for LTIME
            MatchCollection m = Regex.Matches(s, @"(\d+(?:\.\d+)?)(ms|us|ns|d|h|m|s)", RegexOptions.IgnoreCase);
            int used = 0;
            double ns = 0;
            foreach (Match x in m)
            {
                used += x.Length;
                double v = double.Parse(x.Groups[1].Value, IC);
                switch (x.Groups[2].Value.ToLowerInvariant())
                {
                    case "d": ns += v * 86400e9; break;
                    case "h": ns += v * 3600e9; break;
                    case "m": ns += v * 60e9; break;
                    case "s": ns += v * 1e9; break;
                    case "ms": ns += v * 1e6; break;
                    case "us": ns += v * 1e3; break;
                    case "ns": ns += v; break;
                }
            }
            if (m.Count == 0 || used != s.Length) throw new FormatException("Not a valid time: " + t + " (e.g. T#1s500ms)");
            return ltime ? (ulong)Math.Round(ns) : (ulong)Math.Round(ns / 1e6);
        }

        static long ParseTod(string t)
        {
            bool prefixed;
            string s = AfterHash(t, out prefixed);
            ulong raw;
            if (!prefixed && ulong.TryParse(s, NumberStyles.None, IC, out raw)) return (long)raw;
            TimeSpan ts;
            try { ts = TimeSpan.Parse(s); } catch (Exception) { throw new FormatException("Not a valid time of day: " + t + " (e.g. TOD#12:30:00.250)"); }
            if (ts < TimeSpan.Zero || ts.TotalHours >= 24) throw new FormatException("Time of day must be between 00:00 and 23:59:59.999");
            return (long)ts.TotalMilliseconds;
        }

        static long ParseDate(string t)
        {
            bool prefixed;
            string s = AfterHash(t, out prefixed);
            ulong raw;
            if (!prefixed && ulong.TryParse(s, NumberStyles.None, IC, out raw)) return (long)raw;
            string[] fmts = { "yyyy-MM-dd-HH:mm:ss", "yyyy-MM-dd-HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd" };
            DateTime d;
            if (!DateTime.TryParseExact(s, fmts, IC, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out d))
                throw new FormatException("Not a valid date: " + t + " (e.g. DT#2026-10-29-11:34:00)");
            double sec = (d - Epoch).TotalSeconds;
            if (sec < 0 || sec > uint.MaxValue) throw new FormatException("Date out of range");
            return (long)sec;
        }
    }
}
'@

function Get-ErrMsg($err) {
    $ex = $err
    if ($err -is [System.Management.Automation.ErrorRecord]) { $ex = $err.Exception }
    while ($ex.InnerException) { $ex = $ex.InnerException }
    return $ex.Message
}

function Get-AdsCode($err) {
    $ex = $err
    if ($err -is [System.Management.Automation.ErrorRecord]) { $ex = $err.Exception }
    while ($ex) {
        if ($ex -is [AdsMini.AdsException]) { return $ex.Code }
        $ex = $ex.InnerException
    }
    return 0
}

function New-AdsClient {
    $c = New-Object AdsMini.AdsClient
    $c.MaxDepth = $MaxDepth
    $c.MaxArray = $MaxArray
    return $c
}

# ---------------------------------------------------------------------------
# Routes already configured on this PC
#   TC3: <TwinCAT>\3.1\Target\StaticRoutes.xml  (4026: under ProgramData)
#   TC2: HKLM\SOFTWARE\[WOW6432Node\]Beckhoff\TwinCAT\Remote
# ---------------------------------------------------------------------------
function Get-AdsRoutes {
    $routes = New-Object System.Collections.ArrayList
    $seen = @{}
    $add = {
        param($name, $id, $address, $source)
        if ("$id" -notmatch '^\d{1,3}(\.\d{1,3}){5}$' -or $seen.ContainsKey($id)) { return }
        $seen[$id] = $true
        [void]$routes.Add((New-Object PSObject -Property @{ Name = "$name"; NetId = "$id"; Address = "$address"; Source = $source }))
    }

    $xmlPaths = @()
    if ($env:TWINCAT3DIR) { $xmlPaths += (Join-Path $env:TWINCAT3DIR 'Target\StaticRoutes.xml') }
    $xmlPaths += 'C:\TwinCAT\3.1\Target\StaticRoutes.xml'
    if ($env:ProgramData) { $xmlPaths += (Join-Path $env:ProgramData 'Beckhoff\TwinCAT\3.1\Target\StaticRoutes.xml') }
    $pf86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($pf86) { $xmlPaths += (Join-Path $pf86 'Beckhoff\TwinCAT\3.1\Target\StaticRoutes.xml') }

    foreach ($p in $xmlPaths) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try {
            $doc = New-Object System.Xml.XmlDocument
            $doc.Load($p)
            foreach ($r in $doc.SelectNodes('//RemoteConnections/Route')) {
                $n = $r.SelectSingleNode('Name'); $i = $r.SelectSingleNode('NetId'); $a = $r.SelectSingleNode('Address')
                if ($i) { & $add $(if ($n) { $n.InnerText.Trim() }) $i.InnerText.Trim() $(if ($a) { $a.InnerText.Trim() }) $p }
            }
        }
        catch { }
    }

    foreach ($k in @('HKLM:\SOFTWARE\Beckhoff\TwinCAT\Remote', 'HKLM:\SOFTWARE\WOW6432Node\Beckhoff\TwinCAT\Remote')) {
        try {
            if (-not (Test-Path $k)) { continue }
            foreach ($sub in (Get-ChildItem $k -ErrorAction SilentlyContinue)) {
                try {
                    $v = Get-ItemProperty -LiteralPath $sub.PSPath
                    $bin = $v.'AMS NetId'
                    if ($bin -and $bin.Length -ge 6) {
                        & $add $sub.PSChildName ((0..5 | ForEach-Object { [string]$bin[$_] }) -join '.') $v.Address 'TC2 registry'
                    }
                }
                catch { }
            }
        }
        catch { }
    }
    return ,@($routes | Sort-Object Name)
}

function Format-Route($r) {
    if ($r.Address) { return '{0}   {1}   ({2})' -f $r.Name, $r.NetId, $r.Address }
    return '{0}   {1}' -f $r.Name, $r.NetId
}

# Turns whatever is in the Net ID box into a Net ID ('' = local).
# Accepts a raw Net ID, a dropdown entry, "Local ..." or a route name.
function Resolve-NetIdText($text) {
    $t = ("$text").Trim()
    if (-not $t -or $t -match '^local\b') { return '' }
    $m = [regex]::Match($t, '(?<![\d.])\d{1,3}(\.\d{1,3}){5}(?![\d.])')
    if ($m.Success) { return $m.Value }
    foreach ($r in (Get-AdsRoutes)) { if ($r.Name -eq $t) { return $r.NetId } }
    throw "'$t' is not an AMS Net ID (e.g. 5.12.34.56.1.1) or the name of a route on this PC"
}

# Errors that mean the target is gone, so polling should stop instead of piling up timeouts.
$script:FatalCodes = @(0x6, 0x7, 0x745, 0x748)

# ---------------------------------------------------------------------------
# Console mode
# ---------------------------------------------------------------------------
function Start-Cli {
    $dll = [AdsMini.AdsClient]::LoadDll($AdsDll)
    Write-Host "Mini ADS Tool (console)   TcAdsDll: $dll"
    $c = New-AdsClient
    try {
        $c.Connect((Resolve-NetIdText $NetId), $Port, $TimeoutMs)
        $dev = ''
        try { $dev = $c.DeviceInfo() } catch { }
        Write-Host ("Connected to {0}   state: {1}   {2}" -f $c.Target, $c.ReadState(), $dev)
        $n = $c.LoadSymbols()
        Write-Host ("Loaded {0:N0} entries ({1:N0} symbols, {2:N0} types). Type 'help' for commands." -f $n, $c.TopLevelCount, $c.TypeCount)
        if ($c.Truncated) { Write-Host 'Search list was truncated; reading any path by name still works.' -ForegroundColor Yellow }

        $help = @'
  find <text>          search names (several words = all must match, * and ? wildcards)
  read <name>          read a variable, e.g. read MAIN.stData.rTemp
  write <name> <value> write a value, e.g. write MAIN.bStart TRUE   (hex:01 02 for raw bytes)
  info <name>          show type, size and address
  watch <name> [ms]    read repeatedly until a key is pressed (default 500 ms)
  routes               list the ADS routes configured on this PC
  state                show the ADS state
  reload               upload the symbols again (after a PLC download)
  quit
'@
        while ($true) {
            $line = Read-Host 'ads'
            if ($null -eq $line) { break }
            $line = $line.Trim()
            if ($line.Length -eq 0) { continue }
            try {
                if ($line -match '^(q|quit|exit)$') { break }
                elseif ($line -match '^(f|find|s|search)\s+(.+)$') {
                    $res = $c.Search($matches[2], $true)
                    $show = [Math]::Min($res.Count, 100)
                    for ($i = 0; $i -lt $show; $i++) { Write-Host ('  {0,-60} {1}' -f $res[$i].Name, $res[$i].Type) }
                    if ($res.Count -gt $show) { Write-Host ('  ... {0:N0} more, narrow the search' -f ($res.Count - $show)) -ForegroundColor DarkGray }
                    elseif ($res.Count -eq 0) { Write-Host '  no matches' -ForegroundColor DarkGray }
                }
                elseif ($line -match '^(r|read)\s+(\S+)$') {
                    $r = $c.ReadValue($matches[2])
                    Write-Host ('  {0} = {1}' -f $r.Info.Name, $r.Value) -ForegroundColor Green
                }
                elseif ($line -match '^(w|write)\s+(\S+)\s+(.+)$') {
                    $r = $c.WriteValue($matches[2], $matches[3])
                    Write-Host ('  {0} := {1}' -f $r.Info.Name, $r.Value) -ForegroundColor Green
                }
                elseif ($line -match '^(i|info)\s+(\S+)$') {
                    $s = $c.Resolve($matches[2])
                    Write-Host ('  {0}  : {1}  ({2} bytes)  {3}' -f $s.Name, $s.Type, $s.Size, $s.Address)
                    if ($s.Comment) { Write-Host ('  // ' + $s.Comment) -ForegroundColor DarkGray }
                }
                elseif ($line -match '^watch\s+(\S+)(?:\s+(\d+))?$') {
                    $name = $matches[1]
                    $ms = 500
                    if ($matches[2]) { $ms = [int]$matches[2] }
                    Write-Host '  (press any key to stop)' -ForegroundColor DarkGray
                    while (-not [Console]::KeyAvailable) {
                        $r = $c.ReadValue($name)
                        Write-Host ("`r  {0}  {1} = {2}          " -f (Get-Date -Format 'HH:mm:ss.fff'), $name, $r.Value) -NoNewline
                        Start-Sleep -Milliseconds $ms
                    }
                    [void][Console]::ReadKey($true)
                    Write-Host ''
                }
                elseif ($line -match '^state$') { Write-Host ('  ' + $c.ReadState()) }
                elseif ($line -match '^routes$') {
                    Write-Host ('  {0,-28} {1,-20} {2}' -f 'Local (this PC)', [AdsMini.AdsClient]::LocalNetId(), '')
                    foreach ($r in (Get-AdsRoutes)) { Write-Host ('  {0,-28} {1,-20} {2}' -f $r.Name, $r.NetId, $r.Address) }
                    Write-Host '  (restart with -NetId <Net ID or route name> to use one)' -ForegroundColor DarkGray
                }
                elseif ($line -match '^reload$') {
                    $n = $c.LoadSymbols()
                    Write-Host ('  Loaded {0:N0} entries' -f $n)
                }
                else { Write-Host $help }
            }
            catch { Write-Host ('  ' + (Get-ErrMsg $_)) -ForegroundColor Red }
        }
    }
    finally { $c.Dispose() }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
# GUI-only C#: the 10 s chart (a WPF element drawn in OnRender) and the dark title bar call.
# Uses only WPF, which is part of Windows / .NET Framework.
$uiSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace AdsMini
{
    public static class Win
    {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        public static void DarkTitleBar(IntPtr hwnd, bool dark)
        {
            int v = dark ? 1 : 0;
            try
            {
                // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Windows 10 20H1+ / 11), 19 on older Windows 10 builds
                if (DwmSetWindowAttribute(hwnd, 20, ref v, 4) != 0) DwmSetWindowAttribute(hwnd, 19, ref v, 4);
            }
            catch { }
        }
    }

    public class SparkView : FrameworkElement
    {
        public double WindowMs = 10000;
        public bool IsBool;
        public bool Running;
        public string EmptyText = "";
        readonly List<double> ts = new List<double>();
        readonly List<double> vs = new List<double>();
        readonly Stopwatch sw = Stopwatch.StartNew();
        static readonly CultureInfo IC = CultureInfo.InvariantCulture;
        readonly Typeface face = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        Brush lineBrush, textBrush, fillBrush, tipBack, tipText;
        Pen linePen, gridPen, cursorPen;
        double mouseX = -1;

        public SparkView()
        {
            ClipToBounds = true;
            SetColors(Colors.DodgerBlue, Colors.Gray, Colors.Gray, Colors.Black, Colors.White);
            MouseMove += delegate(object s, MouseEventArgs e) { mouseX = e.GetPosition(this).X; InvalidateVisual(); };
            MouseLeave += delegate(object s, MouseEventArgs e) { mouseX = -1; InvalidateVisual(); };
        }

        static Brush Solid(Color c)
        {
            SolidColorBrush b = new SolidColorBrush(c);
            b.Freeze();
            return b;
        }

        public void SetColors(Color line, Color grid, Color text, Color tipBg, Color tipFg)
        {
            lineBrush = Solid(line);
            textBrush = Solid(text);
            tipBack = Solid(tipBg);
            tipText = Solid(tipFg);
            LinearGradientBrush fill = new LinearGradientBrush(Color.FromArgb(0x60, line.R, line.G, line.B), Color.FromArgb(0x00, line.R, line.G, line.B), 90.0);
            fill.Freeze();
            fillBrush = fill;
            linePen = new Pen(lineBrush, 2.0);
            linePen.LineJoin = PenLineJoin.Round;
            linePen.Freeze();
            gridPen = new Pen(Solid(grid), 1.0);
            gridPen.DashStyle = new DashStyle(new double[] { 2, 3 }, 0);
            gridPen.Freeze();
            cursorPen = new Pen(textBrush, 1.0);
            cursorPen.Freeze();
            InvalidateVisual();
        }

        public int Count { get { return vs.Count; } }

        public void Clear()
        {
            ts.Clear();
            vs.Clear();
            InvalidateVisual();
        }

        public void Add(double v)
        {
            double now = sw.Elapsed.TotalMilliseconds;
            ts.Add(now);
            vs.Add(v);
            int drop = 0; // keep one point just before the window so the line enters from the left edge
            while (drop + 1 < ts.Count && ts[drop + 1] <= now - WindowMs) drop++;
            if (drop > 0) { ts.RemoveRange(0, drop); vs.RemoveRange(0, drop); }
            InvalidateVisual();
        }

        public void Refresh() { InvalidateVisual(); }

        static string Num(double v)
        {
            double a = Math.Abs(v);
            if (a >= 1e7 || (a > 0 && a < 1e-3)) return v.ToString("G4", IC);
            return Math.Round(v, 3).ToString(IC);
        }

        FormattedText Text(string s, double size, Brush b)
        {
#pragma warning disable 618
            return new FormattedText(s, IC, FlowDirection.LeftToRight, face, size, b);
#pragma warning restore 618
        }

        protected override void OnRender(DrawingContext dc)
        {
            double w = ActualWidth, h = ActualHeight;
            dc.DrawRectangle(Brushes.Transparent, null, new Rect(0, 0, w, h)); // makes the whole area hoverable
            if (w < 60 || h < 40) return;
            if (vs.Count == 0)
            {
                FormattedText e = Text(EmptyText, 12, textBrush);
                dc.DrawText(e, new Point(Math.Max(0, (w - e.Width) / 2), (h - e.Height) / 2));
                return;
            }

            double now = Running ? sw.Elapsed.TotalMilliseconds : ts[ts.Count - 1];
            double t0 = now - WindowMs;
            double mn = double.MaxValue, mx = double.MinValue;
            foreach (double v in vs) { if (v < mn) mn = v; if (v > mx) mx = v; }
            string sMax, sMin;
            if (IsBool)
            {
                mn = -0.15; mx = 1.15;
                sMax = "TRUE"; sMin = "FALSE";
            }
            else
            {
                if (mx - mn < 1e-12) { double pad = Math.Abs(mx) > 0 ? Math.Abs(mx) * 0.1 : 1; mn -= pad; mx += pad; }
                double p2 = (mx - mn) * 0.12;
                mn -= p2; mx += p2;
                sMax = Num(mx); sMin = Num(mn);
            }

            FormattedText fMax = Text(sMax, 11, textBrush), fMin = Text(sMin, 11, textBrush);
            double L = Math.Max(fMax.Width, fMin.Width) + 16, R = 12, T = 10, B = 26;
            double pw = w - L - R, ph = h - T - B;
            if (pw < 20 || ph < 10) return;

            double yTop = IsBool ? T + ph - (1 - mn) / (mx - mn) * ph : T;
            double yBot = IsBool ? T + ph - (0 - mn) / (mx - mn) * ph : T + ph;
            dc.DrawLine(gridPen, new Point(L, Math.Round(yTop) + 0.5), new Point(L + pw, Math.Round(yTop) + 0.5));
            dc.DrawLine(gridPen, new Point(L, Math.Round((yTop + yBot) / 2) + 0.5), new Point(L + pw, Math.Round((yTop + yBot) / 2) + 0.5));
            dc.DrawLine(gridPen, new Point(L, Math.Round(yBot) + 0.5), new Point(L + pw, Math.Round(yBot) + 0.5));
            for (int s = 0; s <= 10; s += 2)
            {
                double x = Math.Round(L + pw * s / 10.0) + 0.5;
                if (s > 0 && s < 10) dc.DrawLine(gridPen, new Point(x, T), new Point(x, T + ph));
                FormattedText ft = Text(s == 10 ? "now" : "-" + (10 - s) + "s", 11, textBrush);
                dc.DrawText(ft, new Point(Math.Max(0, Math.Min(x - ft.Width / 2, w - ft.Width)), T + ph + 7));
            }
            dc.DrawText(fMax, new Point(L - 8 - fMax.Width, yTop - fMax.Height / 2));
            dc.DrawText(fMin, new Point(L - 8 - fMin.Width, yBot - fMin.Height / 2));

            List<Point> pts = new List<Point>(vs.Count * 2 + 2);
            for (int i = 0; i < vs.Count; i++)
            {
                double x = L + (ts[i] - t0) / WindowMs * pw;
                if (x < L) x = L;
                double y = T + ph - (vs[i] - mn) / (mx - mn) * ph;
                if (IsBool && pts.Count > 0) pts.Add(new Point(x, pts[pts.Count - 1].Y)); // step
                pts.Add(new Point(x, y));
            }
            if (Running) pts.Add(new Point(L + pw, pts[pts.Count - 1].Y)); // hold the last value up to "now"

            dc.PushClip(new RectangleGeometry(new Rect(L, T - 4, pw + 4, ph + 8)));
            if (pts.Count > 1)
            {
                StreamGeometry fill = new StreamGeometry();
                using (StreamGeometryContext g = fill.Open())
                {
                    g.BeginFigure(new Point(pts[0].X, yBot), true, true);
                    g.PolyLineTo(pts, false, false);
                    g.LineTo(new Point(pts[pts.Count - 1].X, yBot), false, false);
                }
                fill.Freeze();
                dc.DrawGeometry(fillBrush, null, fill);

                StreamGeometry line = new StreamGeometry();
                using (StreamGeometryContext g = line.Open())
                {
                    g.BeginFigure(pts[0], false, false);
                    g.PolyLineTo(pts.GetRange(1, pts.Count - 1), true, true);
                }
                line.Freeze();
                dc.DrawGeometry(null, linePen, line);
            }
            dc.DrawEllipse(lineBrush, null, pts[pts.Count - 1], 3.5, 3.5);
            dc.Pop();

            // Hover readout: value at the time under the mouse.
            if (mouseX >= L && mouseX <= L + pw)
            {
                double tMouse = t0 + (mouseX - L) / pw * WindowMs;
                int best = -1;
                for (int i = 0; i < vs.Count; i++) { if (ts[i] <= tMouse) best = i; else break; }
                if (best < 0) best = 0;
                double y = T + ph - (vs[best] - mn) / (mx - mn) * ph;
                dc.DrawLine(cursorPen, new Point(Math.Round(mouseX) + 0.5, T), new Point(Math.Round(mouseX) + 0.5, T + ph));
                dc.DrawEllipse(tipBack, linePen, new Point(mouseX, y), 4, 4);
                string label = (IsBool ? (vs[best] != 0 ? "TRUE" : "FALSE") : Num(vs[best])) + "    " + ((now - tMouse) / 1000.0).ToString("0.0", IC) + " s ago";
                FormattedText ft = Text(label, 12, tipText);
                double bw = ft.Width + 16, bh = ft.Height + 8;
                double bx = mouseX + 12;
                if (bx + bw > w) bx = mouseX - 12 - bw;
                double by = Math.Max(2, Math.Min(y - bh - 10, h - bh - 2));
                dc.DrawRoundedRectangle(tipBack, null, new Rect(bx, by, bw, bh), 6, 6);
                dc.DrawText(ft, new Point(bx + 8, by + 4));
            }
        }
    }
}
'@

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mini ADS Tool" Width="1300" Height="840" MinWidth="980" MinHeight="620"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
  <Window.Resources>
    <!-- Theme colours. These are replaced at runtime when switching dark / light. -->
    <SolidColorBrush x:Key="Bg" Color="#16171B"/>
    <SolidColorBrush x:Key="Surface" Color="#1F2026"/>
    <SolidColorBrush x:Key="Surface2" Color="#2A2B33"/>
    <SolidColorBrush x:Key="Hover" Color="#33343D"/>
    <SolidColorBrush x:Key="Border" Color="#30313A"/>
    <SolidColorBrush x:Key="Text" Color="#E7E8EC"/>
    <SolidColorBrush x:Key="Dim" Color="#9A9CA8"/>
    <SolidColorBrush x:Key="Accent" Color="#4F8CFF"/>
    <SolidColorBrush x:Key="AccentText" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Select" Color="#2D4A80"/>
    <SolidColorBrush x:Key="SelectText" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="Error" Color="#FF7A7A"/>
    <SolidColorBrush x:Key="Ok" Color="#4ADE80"/>
    <SolidColorBrush x:Key="Warn" Color="#FBBF24"/>
    <SolidColorBrush x:Key="Comment" Color="#86C48A"/>
    <SolidColorBrush x:Key="Overlay" Color="#14FFFFFF"/>
    <SolidColorBrush x:Key="TipBg" Color="#E7E8EC"/>
    <SolidColorBrush x:Key="TipFg" Color="#16171B"/>

    <!-- Text helpers -->
    <Style x:Key="Icon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
    <Style x:Key="IconOnly" TargetType="TextBlock" BasedOn="{StaticResource Icon}">
      <Setter Property="Margin" Value="0"/>
    </Style>
    <Style x:Key="Caption" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="{DynamicResource Dim}"/>
    </Style>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="14"/>
    </Style>
    <Style x:Key="Pill" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource Surface2}"/>
      <Setter Property="CornerRadius" Value="10"/>
      <Setter Property="Padding" Value="9,3"/>
      <Setter Property="Margin" Value="0,0,6,6"/>
    </Style>

    <!-- Buttons: rounded, with a hover overlay so the same template works for normal and accent. -->
    <Style TargetType="Button">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Background" Value="{DynamicResource Surface2}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Grid>
              <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6"/>
              <Border x:Name="ov" Background="{DynamicResource Overlay}" CornerRadius="6" Opacity="0"/>
              <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="ov" Property="Opacity" Value="1"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.8"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="{DynamicResource Accent}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Accent}"/>
      <Setter Property="Foreground" Value="{DynamicResource AccentText}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <!-- Text boxes -->
    <Style TargetType="TextBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Background" Value="{DynamicResource Surface2}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,0"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="CaretBrush" Value="{DynamicResource Text}"/>
      <Setter Property="SelectionBrush" Value="{DynamicResource Accent}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False"
                            VerticalAlignment="{TemplateBinding VerticalContentAlignment}"
                            HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Dim}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Check boxes look like toggle switches -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Grid Width="36" Height="20" VerticalAlignment="Center">
                <Border x:Name="track" CornerRadius="10" Background="{DynamicResource Surface2}"
                        BorderBrush="{DynamicResource Dim}" BorderThickness="1"/>
                <Ellipse x:Name="knob" Width="12" Height="12" Fill="{DynamicResource Dim}"
                         HorizontalAlignment="Left" Margin="4,0,0,0"/>
              </Grid>
              <ContentPresenter Margin="8,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="track" Property="Background" Value="{DynamicResource Accent}"/>
                <Setter TargetName="track" Property="BorderBrush" Value="{DynamicResource Accent}"/>
                <Setter TargetName="knob" Property="Fill" Value="{DynamicResource AccentText}"/>
                <Setter TargetName="knob" Property="HorizontalAlignment" Value="Right"/>
                <Setter TargetName="knob" Property="Margin" Value="0,0,4,0"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="track" Property="Opacity" Value="0.85"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Slim scroll bars -->
    <Style x:Key="ThumbStyle" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border CornerRadius="4" Background="{DynamicResource Dim}" Opacity="0.45" Margin="2"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Grid Background="Transparent">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb>
                  <Thumb Style="{StaticResource ThumbStyle}"/>
                </Track.Thumb>
              </Track>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="10"/>
          <Setter Property="MinHeight" Value="10"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Grid Background="Transparent">
                  <Track x:Name="PART_Track" IsDirectionReversed="False">
                    <Track.Thumb>
                      <Thumb Style="{StaticResource ThumbStyle}"/>
                    </Track.Thumb>
                  </Track>
                </Grid>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <!-- List views -->
    <Style TargetType="ListView">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
      <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
    </Style>
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="Foreground" Value="{DynamicResource Dim}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GridViewColumnHeader">
            <Grid>
              <Border Background="{TemplateBinding Background}" BorderBrush="{DynamicResource Border}"
                      BorderThickness="0,0,0,1" Padding="8,7">
                <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
              </Border>
              <Thumb x:Name="PART_HeaderGripper" HorizontalAlignment="Right" Width="9" Margin="0,0,-4,0" Cursor="SizeWE">
                <Thumb.Template>
                  <ControlTemplate TargetType="Thumb">
                    <Border Background="Transparent">
                      <Rectangle Width="1" Margin="0,8" Fill="{DynamicResource Border}" HorizontalAlignment="Center"/>
                    </Border>
                  </ControlTemplate>
                </Thumb.Template>
              </Thumb>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ListViewItem">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListViewItem">
            <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="0,4" Margin="0,1">
              <GridViewRowPresenter Columns="{TemplateBinding GridView.ColumnCollection}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource Hover}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource Select}"/>
                <Setter Property="Foreground" Value="{DynamicResource SelectText}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="WatchItemStyle" TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
      <Style.Triggers>
        <DataTrigger Binding="{Binding IsError}" Value="True">
          <Setter Property="Foreground" Value="{DynamicResource Error}"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <!-- Editable combo box (target picker) -->
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Padding" Value="10,7"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}" Margin="0,1">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource Hover}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource Select}"/>
                <Setter Property="Foreground" Value="{DynamicResource SelectText}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="MinHeight" Value="34"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton x:Name="toggle" Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="tb" Background="{DynamicResource Surface2}" BorderBrush="{DynamicResource Border}"
                            BorderThickness="1" CornerRadius="6">
                      <TextBlock Text="&#xE70D;" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="10"
                                 Foreground="{DynamicResource Dim}" HorizontalAlignment="Right"
                                 VerticalAlignment="Center" Margin="0,0,11,0"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="tb" Property="BorderBrush" Value="{DynamicResource Dim}"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter x:Name="content" IsHitTestVisible="False" Margin="10,0,30,0" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <TextBox x:Name="PART_EditableTextBox" Visibility="Hidden" Margin="1,1,28,1"
                       Background="Transparent" BorderThickness="0"
                       IsReadOnly="{TemplateBinding IsReadOnly}"/>
              <Popup x:Name="PART_Popup" Placement="Bottom" AllowsTransparency="True" Focusable="False"
                     PopupAnimation="Fade" IsOpen="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}}">
                <Border Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}" BorderThickness="1"
                        CornerRadius="8" Margin="0,4,0,0" Padding="4"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}"
                        MaxHeight="{TemplateBinding MaxDropDownHeight}">
                  <ScrollViewer>
                    <ItemsPresenter/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEditable" Value="True">
                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                <Setter TargetName="content" Property="Visibility" Value="Hidden"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ToolTip">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Border}"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- ===== Header ===== -->
    <Border Grid.Row="0" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}"
            BorderThickness="0,0,0,1" Padding="16,10">
      <DockPanel LastChildFill="False">
        <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="34" Height="34" CornerRadius="9" Background="{DynamicResource Accent}" Margin="0,0,10,0">
            <TextBlock Text="ADS" FontWeight="Bold" FontSize="11" Foreground="{DynamicResource AccentText}"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center" Margin="0,0,22,0">
            <TextBlock Text="Mini ADS Tool" FontWeight="SemiBold" FontSize="15"/>
            <TextBlock Text="TwinCAT ADS test tool" Style="{StaticResource Caption}" FontWeight="Normal"/>
          </StackPanel>
          <ComboBox x:Name="cbNet" IsEditable="True" Width="340" MaxDropDownHeight="420" VerticalAlignment="Center"
                    ToolTip="Pick a route (connects straight away) or type an AMS Net ID and press Enter"/>
          <TextBlock Text="Port" Foreground="{DynamicResource Dim}" VerticalAlignment="Center" Margin="12,0,6,0"/>
          <TextBox x:Name="tbPort" Width="64" ToolTip="TC3 PLC = 851, TC2 PLC = 801"/>
          <Button x:Name="btnConnect" Style="{StaticResource AccentButton}" Margin="12,0,0,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource Icon}" Text="&#xE71B;"/>
              <TextBlock Text="Connect" VerticalAlignment="Center"/>
            </StackPanel>
          </Button>
          <Button x:Name="btnDisconnect" Margin="8,0,0,0" IsEnabled="False">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource Icon}" Text="&#xE711;"/>
              <TextBlock Text="Disconnect" VerticalAlignment="Center"/>
            </StackPanel>
          </Button>
          <Button x:Name="btnReload" Margin="8,0,0,0" Padding="10,6" IsEnabled="False" ToolTip="Reload symbols (after a PLC download)">
            <TextBlock Style="{StaticResource IconOnly}" Text="&#xE72C;"/>
          </Button>
        </StackPanel>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="chip" CornerRadius="14" Background="{DynamicResource Surface2}" Padding="12,6" Margin="0,0,10,0">
            <StackPanel Orientation="Horizontal">
              <Ellipse x:Name="dot" Width="8" Height="8" Fill="{DynamicResource Dim}" VerticalAlignment="Center" Margin="0,0,8,0"/>
              <TextBlock x:Name="lblState" Text="Not connected" VerticalAlignment="Center"/>
            </StackPanel>
          </Border>
          <Button x:Name="btnTheme" Padding="10,6" ToolTip="Switch light / dark">
            <TextBlock x:Name="themeIcon" Style="{StaticResource IconOnly}" Text="&#xE706;"/>
          </Button>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- ===== Content ===== -->
    <Grid Grid.Row="1" Margin="12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="5*" MinWidth="320"/>
        <ColumnDefinition Width="12"/>
        <ColumnDefinition Width="6*" MinWidth="420"/>
      </Grid.ColumnDefinitions>

      <!-- Symbols -->
      <Border Grid.Column="0" Style="{StaticResource Card}">
        <DockPanel>
          <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid>
              <TextBox x:Name="tbSearch" Padding="32,0,8,0"/>
              <TextBlock Style="{StaticResource IconOnly}" Text="&#xE721;" Foreground="{DynamicResource Dim}"
                         HorizontalAlignment="Left" Margin="12,0,0,0" IsHitTestVisible="False"/>
              <TextBlock x:Name="searchHint" Text="Search name or type, e.g.  motor speed   REAL   MAIN.*"
                         Foreground="{DynamicResource Dim}" Margin="34,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
            </Grid>
            <CheckBox x:Name="cbMembers" Grid.Column="1" Content="Members" IsChecked="True" Margin="14,0,2,0"
                      ToolTip="Include struct / FB members and array elements"/>
          </Grid>
          <TextBlock x:Name="lblCount" DockPanel.Dock="Bottom" Style="{StaticResource Caption}" FontWeight="Normal" Margin="4,8,0,0"/>
          <ListView x:Name="lvSym" SelectionMode="Single">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="Name" Width="330" DisplayMemberBinding="{Binding Name}"/>
                <GridViewColumn Header="Type" Width="150" DisplayMemberBinding="{Binding Type}"/>
                <GridViewColumn Header="Size" Width="56" DisplayMemberBinding="{Binding Size}"/>
              </GridView>
            </ListView.View>
          </ListView>
        </DockPanel>
      </Border>

      <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent"/>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="3*" MinHeight="380"/>
          <RowDefinition Height="12"/>
          <RowDefinition Height="2*" MinHeight="150"/>
        </Grid.RowDefinitions>

        <!-- Variable -->
        <Border Grid.Row="0" Style="{StaticResource Card}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <Grid Grid.Row="0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="tbName" FontFamily="Cascadia Mono, Consolas" ToolTip="Any path works, e.g. MAIN.aData[7].rValue (Enter reads it)"/>
              <Button x:Name="btnRead" Grid.Column="1" Margin="8,0,0,0">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Style="{StaticResource Icon}" Text="&#xE896;"/>
                  <TextBlock Text="Read" VerticalAlignment="Center"/>
                </StackPanel>
              </Button>
              <Button x:Name="btnWatch" Grid.Column="2" Margin="8,0,0,0" ToolTip="Add to watch list">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Style="{StaticResource Icon}" Text="&#xE710;"/>
                  <TextBlock Text="Watch" VerticalAlignment="Center"/>
                </StackPanel>
              </Button>
            </Grid>

            <Grid Grid.Row="1" Margin="2,16,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <StackPanel>
                <TextBlock Text="VALUE" Style="{StaticResource Caption}"/>
                <TextBlock x:Name="lblValue" Text="-" FontFamily="Cascadia Mono, Consolas" FontSize="28"
                           TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
              </StackPanel>
              <CheckBox x:Name="cbLive" Grid.Column="1" Content="Live" VerticalAlignment="Bottom" Margin="0,0,0,6"
                        ToolTip="Read every 100 ms and chart the last 10 s"/>
            </Grid>

            <WrapPanel Grid.Row="2" Margin="0,10,0,0">
              <Border x:Name="pillType" Style="{StaticResource Pill}" Visibility="Collapsed">
                <TextBlock x:Name="lblType" FontSize="12"/>
              </Border>
              <Border x:Name="pillAddr" Style="{StaticResource Pill}" Visibility="Collapsed">
                <TextBlock x:Name="lblAddr" FontSize="12" Foreground="{DynamicResource Dim}"/>
              </Border>
              <Border x:Name="pillComment" Style="{StaticResource Pill}" Visibility="Collapsed">
                <TextBlock x:Name="lblComment" FontSize="12" Foreground="{DynamicResource Comment}"
                           TextTrimming="CharacterEllipsis" MaxWidth="520"/>
              </Border>
            </WrapPanel>

            <Border x:Name="chartHost" Grid.Row="3" Margin="0,4,0,12" CornerRadius="8"
                    Background="{DynamicResource Surface2}" MinHeight="120" Padding="4"/>

            <Grid Grid.Row="4">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="tbNew" FontFamily="Cascadia Mono, Consolas"
                       ToolTip="TRUE, 42, 16#FF, 1.5, T#2s, 'text', E_State.Run, hex:01 02 (Enter writes)"/>
              <Button x:Name="btnWrite" Grid.Column="1" Style="{StaticResource AccentButton}" Margin="8,0,0,0">
                <StackPanel Orientation="Horizontal">
                  <TextBlock Style="{StaticResource Icon}" Text="&#xE898;"/>
                  <TextBlock Text="Write" VerticalAlignment="Center"/>
                </StackPanel>
              </Button>
              <CheckBox x:Name="cbConfirm" Grid.Column="2" Content="Confirm" IsChecked="True" Margin="14,0,0,0"
                        ToolTip="Ask before writing"/>
            </Grid>
            <TextBlock Grid.Row="5" Style="{StaticResource Caption}" FontWeight="Normal" Margin="2,8,0,0"
                       Text="Values: TRUE, 42, 16#FF, 1.5, T#2s, 'text', E_State.Run, hex:01 02"/>
          </Grid>
        </Border>

        <GridSplitter Grid.Row="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent"/>

        <!-- Watch list -->
        <Border Grid.Row="2" Style="{StaticResource Card}">
          <DockPanel>
            <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
              <TextBlock Text="Watch list" FontWeight="SemiBold" FontSize="14" VerticalAlignment="Center"/>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <CheckBox x:Name="cbWatchLive" Content="Live" Margin="0,0,14,0" ToolTip="Refresh every 500 ms"/>
                <Button x:Name="btnRemove" Padding="10,4" ToolTip="Remove selected">
                  <TextBlock Style="{StaticResource IconOnly}" Text="&#xE74D;"/>
                </Button>
                <Button x:Name="btnClear" Content="Clear" Margin="6,0,0,0" Padding="12,4"/>
              </StackPanel>
            </DockPanel>
            <ListView x:Name="lvWatch" SelectionMode="Extended" ItemContainerStyle="{StaticResource WatchItemStyle}">
              <ListView.View>
                <GridView>
                  <GridViewColumn Header="Name" Width="250" DisplayMemberBinding="{Binding Name}"/>
                  <GridViewColumn Header="Value" Width="200" DisplayMemberBinding="{Binding Value}"/>
                  <GridViewColumn Header="Type" Width="120" DisplayMemberBinding="{Binding Type}"/>
                </GridView>
              </ListView.View>
            </ListView>
          </DockPanel>
        </Border>
      </Grid>
    </Grid>

    <!-- ===== Status ===== -->
    <Border Grid.Row="2" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Border}"
            BorderThickness="0,1,0,0" Padding="16,6">
      <TextBlock x:Name="lblStatus" Foreground="{DynamicResource Dim}" FontSize="12" TextTrimming="CharacterEllipsis"/>
    </Border>
  </Grid>
</Window>
'@

function Start-Gui {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'The GUI needs an STA thread. Start it with: powershell.exe -STA -File AdsTool.ps1 (AdsTool.bat does this).'
    }
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    if (-not ('AdsMini.SparkView' -as [type])) {
        $refs = @([System.Windows.FrameworkElement].Assembly.Location, [System.Windows.Media.Visual].Assembly.Location,
                  [System.Windows.DependencyObject].Assembly.Location, [System.Xaml.XamlReader].Assembly.Location)
        if ($PSVersionTable.PSEdition -eq 'Core') {
            $refs += @('System.Runtime', 'System.Collections', 'System.Runtime.InteropServices', 'System.ObjectModel', 'netstandard')
        }
        else { $refs += @([System.Diagnostics.Stopwatch].Assembly.Location) }
        Add-Type -TypeDefinition $uiSource -ReferencedAssemblies $refs -Language CSharp -IgnoreWarnings
    }

    $dll = [AdsMini.AdsClient]::LoadDll($AdsDll)

    $startDark = $Theme -eq 'Dark'
    if ($Theme -eq 'Auto') {
        try {
            $v = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction Stop
            $startDark = ($v.AppsUseLightTheme -eq 0)
        }
        catch { }
    }

    $window = [System.Windows.Markup.XamlReader]::Parse($xaml)
    foreach ($m in [regex]::Matches($xaml, 'x:Name="(\w+)"')) {
        $el = $window.FindName($m.Groups[1].Value)
        if ($el) { Set-Variable -Name $m.Groups[1].Value -Value $el }
    }

    $spark = New-Object AdsMini.SparkView
    $spark.EmptyText = 'Turn on Live to chart the last 10 s of a number or BOOL'
    $chartHost.Child = $spark

    # Shared mutable state lives in a hashtable so event handlers can change it.
    $S = @{
        Client    = (New-AdsClient)
        Results   = (New-Object 'System.Collections.Generic.List[AdsMini.SymbolInfo]')
        Watch     = (New-Object 'System.Collections.ObjectModel.ObservableCollection[AdsMini.WatchItem]')
        Busy      = $false
        LocalId   = $null
        ChartName = ''
        Tick      = 0
        Dark      = $startDark
        Hwnd      = [IntPtr]::Zero
        Loading   = $false
        SelBefore = -1
    }
    $lvWatch.ItemsSource = $S.Watch

    $Palettes = @{
        Dark  = @{ Bg = '#16171B'; Surface = '#1F2026'; Surface2 = '#2A2B33'; Hover = '#33343D'; Border = '#30313A'
                   Text = '#E7E8EC'; Dim = '#9A9CA8'; Accent = '#4F8CFF'; AccentText = '#FFFFFF'; Select = '#2D4A80'
                   SelectText = '#FFFFFF'; Error = '#FF7A7A'; Ok = '#4ADE80'; Warn = '#FBBF24'; Comment = '#86C48A'
                   Overlay = '#14FFFFFF'; TipBg = '#E7E8EC'; TipFg = '#16171B' }
        Light = @{ Bg = '#F3F4F7'; Surface = '#FFFFFF'; Surface2 = '#F1F2F5'; Hover = '#E9EAF0'; Border = '#E0E2E8'
                   Text = '#1D1F24'; Dim = '#6B6F7B'; Accent = '#2563EB'; AccentText = '#FFFFFF'; Select = '#DCE7FF'
                   SelectText = '#10203F'; Error = '#C62828'; Ok = '#16A34A'; Warn = '#D97706'; Comment = '#2E7D32'
                   Overlay = '#0D000000'; TipBg = '#1D1F24'; TipFg = '#FFFFFF' }
    }

    function Get-Color($hex) { [System.Windows.Media.Color][System.Windows.Media.ColorConverter]::ConvertFromString($hex) }

    function Set-Theme($dark) {
        $S.Dark = $dark
        $p = $Palettes[$(if ($dark) { 'Dark' } else { 'Light' })]
        foreach ($k in $p.Keys) {
            $b = New-Object System.Windows.Media.SolidColorBrush (Get-Color $p[$k])
            $b.Freeze()
            # set_Item, not [$k] = : PowerShell's indexer stores a PSObject wrapper
            # round the brush, which WPF rejects when a DynamicResource wants a Brush.
            $window.Resources.set_Item($k, $b)
        }
        $spark.SetColors((Get-Color $p.Accent), (Get-Color $p.Border), (Get-Color $p.Dim), (Get-Color $p.TipBg), (Get-Color $p.TipFg))
        if ($dark) { $themeIcon.Text = [string][char]0xE706 } else { $themeIcon.Text = [string][char]0xE708 }   # sun / moon
        if ($S.Hwnd -ne [IntPtr]::Zero) { [AdsMini.Win]::DarkTitleBar($S.Hwnd, $dark) }
    }

    function Set-Status($msg, [switch]$Err) {
        $lblStatus.Text = (Get-Date -Format 'HH:mm:ss') + '    ' + $msg
        $lblStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($Err) { 'Error' } else { 'Dim' }))
    }

    # Lets WPF repaint before a long blocking call (e.g. the symbol upload).
    function Update-UiNow {
        $window.Dispatcher.Invoke([action] { }, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    function Set-StateChip($text, $brushKey) {
        $lblState.Text = $text
        $dot.SetResourceReference([System.Windows.Shapes.Shape]::FillProperty, $brushKey)
    }

    function Set-Connected($on) {
        $btnDisconnect.IsEnabled = $on
        $btnReload.IsEnabled = $on
    }

    # Refills the Net ID dropdown, keeping whatever is typed in the box.
    function Update-RouteList {
        $S.Loading = $true
        try {
            $keep = $cbNet.Text
            if ($null -eq $S.LocalId) {
                $S.LocalId = ''
                try { $S.LocalId = [AdsMini.AdsClient]::LocalNetId() } catch { }
            }
            $routes = Get-AdsRoutes
            $cbNet.Items.Clear()
            [void]$cbNet.Items.Add(('Local (this PC)   ' + $S.LocalId).Trim())
            foreach ($r in $routes) { [void]$cbNet.Items.Add((Format-Route $r)) }
            $cbNet.Text = $keep
            return @($routes).Count
        }
        finally { $S.Loading = $false }
    }

    function Update-State {
        try {
            $dev = ''
            try { $dev = '   ' + $S.Client.DeviceInfo() } catch { }
            $st = $S.Client.ReadState()
            $key = 'Warn'
            if ($st -eq 'Run') { $key = 'Ok' }
            Set-StateChip ('{0}   {1}{2}' -f $S.Client.Target, $st, $dev) $key
        }
        catch { Set-StateChip ($S.Client.Target + '   ' + (Get-ErrMsg $_)) 'Error' }
    }

    function Update-Search {
        $S.Results = $S.Client.Search($tbSearch.Text, [bool]$cbMembers.IsChecked)
        $lvSym.ItemsSource = $S.Results
        if ($S.Results.Count -gt 0) { $lvSym.ScrollIntoView($S.Results[0]) }
        $lblCount.Text = '{0:N0} of {1:N0} symbols' -f $S.Results.Count, $S.Client.Symbols.Count
    }

    function Update-Symbols {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            Set-Status 'Uploading symbols...'
            Update-UiNow
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $n = $S.Client.LoadSymbols()
            $msg = 'Loaded {0:N0} entries ({1:N0} symbols, {2:N0} types) in {3} ms' -f $n, $S.Client.TopLevelCount, $S.Client.TypeCount, $sw.ElapsedMilliseconds
            if ($S.Client.Truncated) { $msg += '. List truncated; typing a full path in Name still works.' }
            Set-Status $msg
            Update-Search
        }
        finally { $window.Cursor = $null }
    }

    function Connect-Target {
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            $p = 0
            if (-not [int]::TryParse($tbPort.Text.Trim(), [ref]$p)) { throw 'Port must be a number (TC3 PLC = 851, TC2 PLC = 801)' }
            $cbLive.IsChecked = $false
            Set-Status 'Connecting...'
            Update-UiNow
            $S.Client.Connect((Resolve-NetIdText $cbNet.Text), $p, $TimeoutMs)
            Set-Connected $true
            Update-State
            Update-Symbols
            [void]$tbSearch.Focus()
        }
        catch {
            $S.Client.Disconnect()
            Set-Connected $false
            Set-StateChip 'Not connected' 'Dim'
            Set-Status (Get-ErrMsg $_) -Err
            [void][System.Windows.MessageBox]::Show($window, (Get-ErrMsg $_), 'Connect failed', 'OK', 'Error')
        }
        finally { $window.Cursor = $null }
    }

    function Show-Read($r) {
        $i = $r.Info
        $lblType.Text = '{0}   {1} bytes' -f $i.Type, $i.Size
        $lblAddr.Text = $i.Address
        $lblComment.Text = ($i.Comment -replace '\s+', ' ')
        $pillType.Visibility = 'Visible'
        $pillAddr.Visibility = 'Visible'
        $pillComment.Visibility = $(if ($lblComment.Text) { 'Visible' } else { 'Collapsed' })
        if ($lblValue.Text -ne $r.Value) { $lblValue.Text = $r.Value }

        # chart: restart when the variable changes, then add this sample if it's numeric
        if ($S.ChartName -ne $i.Name) {
            $spark.Clear()
            $S.ChartName = $i.Name
            $spark.IsBool = $S.Client.IsBoolType($i)
        }
        $v = $S.Client.NumericValue($i, $r.Raw)
        if (-not [double]::IsNaN($v)) { $spark.Add($v) }
    }

    function Read-Current([switch]$Quiet) {
        $name = $tbName.Text.Trim()
        if (-not $name) { return }
        if (-not $S.Client.Connected) { if (-not $Quiet) { Set-Status 'Not connected' -Err }; return }
        try {
            $r = $S.Client.ReadValue($name)
            Show-Read $r
            if (-not $Quiet) { Set-Status ('Read {0} = {1}' -f $name, $r.Value) }
        }
        catch {
            $lblValue.Text = '-'
            $cbLive.IsChecked = $false
            $spark.Running = $false
            Set-Status (Get-ErrMsg $_) -Err
        }
    }

    function Write-Current {
        $name = $tbName.Text.Trim()
        $val = $tbNew.Text
        if (-not $name) { return }
        if (-not $S.Client.Connected) { Set-Status 'Not connected' -Err; return }
        if ($cbConfirm.IsChecked) {
            $ans = [System.Windows.MessageBox]::Show($window, "Write`n`n    $val`n`nto $name ?", 'Confirm write', 'YesNo', 'Question')
            if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }
        try {
            $r = $S.Client.WriteValue($name, $val)
            Show-Read $r
            Set-Status ('Wrote {0}, read back {1}' -f $name, $r.Value)
        }
        catch { Set-Status (Get-ErrMsg $_) -Err }
    }

    function Add-Watch($name) {
        $name = ("$name").Trim()
        if (-not $name) { return }
        foreach ($w in $S.Watch) { if ($w.Name -eq $name) { return } }
        $S.Watch.Add((New-Object AdsMini.WatchItem($name)))
        Update-Watch
    }

    function Update-Watch {
        if (-not $S.Client.Connected) { return }
        foreach ($w in @($S.Watch)) {
            try {
                $r = $S.Client.ReadValue($w.Name)
                $w.Value = $r.Value
                $w.Type = $r.Info.Type
                $w.IsError = $false
            }
            catch {
                $w.Value = Get-ErrMsg $_
                $w.IsError = $true
                if ($script:FatalCodes -contains (Get-AdsCode $_)) {
                    $cbWatchLive.IsChecked = $false
                    Set-Status (Get-ErrMsg $_) -Err
                    break
                }
            }
        }
    }

    # Keep the Name column filling the list width.
    function Fit-NameColumn($lv, $index) {
        $gv = $lv.View
        $other = 0
        for ($i = 0; $i -lt $gv.Columns.Count; $i++) { if ($i -ne $index) { $other += $gv.Columns[$i].ActualWidth } }
        $w = $lv.ActualWidth - $other - 24
        if ($w -gt 120) { $gv.Columns[$index].Width = $w }
    }

    # --- events -----------------------------------------------------------------
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $searchTimer = New-Object System.Windows.Threading.DispatcherTimer
    $searchTimer.Interval = [TimeSpan]::FromMilliseconds(250)

    $btnConnect.Add_Click({ Connect-Target })
    $btnDisconnect.Add_Click({
        $cbLive.IsChecked = $false
        $spark.Running = $false
        $cbWatchLive.IsChecked = $false
        $S.Client.Disconnect()
        $S.Client.Symbols.Clear()
        Update-Search
        Set-Connected $false
        Set-StateChip 'Not connected' 'Dim'
        Set-Status 'Disconnected'
    })
    $btnReload.Add_Click({ try { Update-State; Update-Symbols } catch { Set-Status (Get-ErrMsg $_) -Err } })
    $btnTheme.Add_Click({ Set-Theme (-not $S.Dark) })

    # Picking a route from the list connects straight away; typing still needs Enter / Connect.
    $cbNet.Add_DropDownOpened({
        [void](Update-RouteList)
        $S.SelBefore = $cbNet.SelectedIndex
    })
    $cbNet.Add_DropDownClosed({
        if ($cbNet.SelectedIndex -ge 0 -and $cbNet.SelectedIndex -ne $S.SelBefore) {
            [void]$window.Dispatcher.BeginInvoke([action] { Connect-Target })
        }
    })
    $cbNet.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter -and -not $cbNet.IsDropDownOpen) { $e.Handled = $true; Connect-Target }
    })
    $tbPort.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) { $e.Handled = $true; Connect-Target }
    })

    $tbSearch.Add_TextChanged({
        $searchHint.Visibility = $(if ($tbSearch.Text) { 'Collapsed' } else { 'Visible' })
        $searchTimer.Stop()
        $searchTimer.Start()
    })
    $searchTimer.Add_Tick({ $searchTimer.Stop(); Update-Search })
    $cbMembers.Add_Click({ Update-Search })

    $lvSym.Add_SelectionChanged({
        $sym = $lvSym.SelectedItem
        if ($sym) {
            $tbName.Text = $sym.Name
            Read-Current
        }
    })
    $lvSym.Add_MouseDoubleClick({ if ($lvSym.SelectedItem) { Add-Watch $lvSym.SelectedItem.Name } })
    $lvSym.Add_SizeChanged({ Fit-NameColumn $lvSym 0 })
    $lvWatch.Add_SizeChanged({ Fit-NameColumn $lvWatch 1 })

    $tbName.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) { $e.Handled = $true; Read-Current }
    })
    $tbNew.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Enter) { $e.Handled = $true; Write-Current }
    })
    $btnRead.Add_Click({ Read-Current })
    $btnWrite.Add_Click({ Write-Current })
    $btnWatch.Add_Click({ Add-Watch $tbName.Text })
    $cbLive.Add_Click({
        $spark.Running = [bool]$cbLive.IsChecked
        $spark.Refresh()
    })

    $lvWatch.Add_SelectionChanged({
        $w = $lvWatch.SelectedItem
        if ($w) {
            $tbName.Text = $w.Name
            Read-Current -Quiet
        }
    })
    $btnRemove.Add_Click({
        foreach ($w in @($lvWatch.SelectedItems)) { [void]$S.Watch.Remove($w) }
    })
    $btnClear.Add_Click({ $S.Watch.Clear() })

    # 100 ms tick: the Live variable (and chart) every tick, the watch list every 5th.
    $timer.Add_Tick({
        if ($S.Busy -or -not $S.Client.Connected) { return }
        $S.Busy = $true
        try {
            $S.Tick++
            if ($cbLive.IsChecked) { Read-Current -Quiet }
            if ($cbWatchLive.IsChecked -and ($S.Tick % 5 -eq 0)) { Update-Watch }
        }
        finally { $S.Busy = $false }
    })

    $window.Add_SourceInitialized({
        $S.Hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
        [AdsMini.Win]::DarkTitleBar($S.Hwnd, $S.Dark)
    })
    $window.Add_Loaded({
        $tbPort.Text = [string]$Port
        $n = Update-RouteList
        $S.Loading = $true
        if ($NetId) { $cbNet.Text = $NetId } else { $cbNet.SelectedIndex = 0 }
        $S.Loading = $false
        Update-Search
        Set-Status ('Ready. {0} route(s) found. TcAdsDll: {1}' -f $n, $dll)
        [void]$btnConnect.Focus()
        $timer.Start()
    })
    $window.Add_Closing({
        $timer.Stop()
        $searchTimer.Stop()
        $S.Client.Dispose()
    })

    Set-Theme $startDark
    [void]$window.ShowDialog()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    if (-not ('AdsMini.AdsClient' -as [type])) {
        Add-Type -TypeDefinition $source -Language CSharp
    }
    if ($Cli) { Start-Cli } else { Start-Gui }
}
catch {
    $msg = Get-ErrMsg $_
    if ($Cli) {
        Write-Host $msg -ForegroundColor Red
    }
    else {
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [void][System.Windows.Forms.MessageBox]::Show($msg, 'Mini ADS Tool', 'OK', 'Error')
        }
        catch { Write-Host $msg -ForegroundColor Red }
    }
    exit 1
}
