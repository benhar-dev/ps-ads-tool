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
        public string Name = "", Type = "", Comment = "";
        public uint IGroup, IOffset, Size, DataType, Flags, Handle;
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
            string q = (query ?? "").Trim();
            List<SymbolInfo> r = new List<SymbolInfo>();
            Regex rx = null;
            string[] terms = null;
            if (q.IndexOf('*') >= 0 || q.IndexOf('?') >= 0)
                rx = new Regex("^" + Regex.Escape(q).Replace(@"\*", ".*").Replace(@"\?", ".") + "$", RegexOptions.IgnoreCase);
            else
                terms = q.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (SymbolInfo s in Symbols)
            {
                if (!includeMembers && s.IsMember) continue;
                if (rx != null)
                {
                    if (!rx.IsMatch(s.Name)) continue;
                }
                else
                {
                    bool ok = true;
                    foreach (string t in terms)
                        if (s.Name.IndexOf(t, StringComparison.OrdinalIgnoreCase) < 0) { ok = false; break; }
                    if (!ok) continue;
                }
                r.Add(s);
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
# GUI-only helpers: the 10 s chart (drawn with GDI+) and the Windows calls for
# the dark title bar / dark scrollbars. Only needs System.Drawing, which is part of Windows.
$uiSource = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Runtime.InteropServices;

namespace AdsMini
{
    public static class Win
    {
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);

        public static void DarkTitleBar(IntPtr hwnd, bool dark)
        {
            int v = dark ? 1 : 0;
            try
            {
                // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Windows 10 20H1+ / 11), 19 on older Windows 10 builds
                if (DwmSetWindowAttribute(hwnd, 20, ref v, 4) != 0) DwmSetWindowAttribute(hwnd, 19, ref v, 4);
                SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0, 0x0027); // NOMOVE|NOSIZE|NOZORDER|FRAMECHANGED
            }
            catch { }
        }

        public static void Theme(IntPtr hwnd, string name)
        {
            try { SetWindowTheme(hwnd, string.IsNullOrEmpty(name) ? null : name, null); } catch { }
        }
    }

    public class Spark
    {
        public double WindowMs = 10000;
        public bool IsBool;
        public bool Running;
        readonly List<double> ts = new List<double>();
        readonly List<double> vs = new List<double>();
        readonly Stopwatch sw = Stopwatch.StartNew();
        static readonly CultureInfo IC = CultureInfo.InvariantCulture;

        public int Count { get { return vs.Count; } }

        public void Clear() { ts.Clear(); vs.Clear(); }

        public void Add(double v)
        {
            double now = sw.Elapsed.TotalMilliseconds;
            ts.Add(now);
            vs.Add(v);
            int drop = 0; // keep one point just before the window so the line enters from the left edge
            while (drop + 1 < ts.Count && ts[drop + 1] <= now - WindowMs) drop++;
            if (drop > 0) { ts.RemoveRange(0, drop); vs.RemoveRange(0, drop); }
        }

        static string Num(double v)
        {
            double a = Math.Abs(v);
            if (a >= 1e7 || (a > 0 && a < 1e-3)) return v.ToString("G4", IC);
            return Math.Round(v, 3).ToString(IC);
        }

        public void Paint(Graphics g, int w, int h, Font font, Color back, Color grid, Color line, Color text, string emptyText)
        {
            g.Clear(back);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (Brush tb = new SolidBrush(text))
            using (Pen gp = new Pen(grid))
            {
                if (vs.Count == 0)
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    g.DrawString(emptyText, font, tb, new RectangleF(0, 0, w, h), sf);
                    return;
                }
                double now = Running ? sw.Elapsed.TotalMilliseconds : ts[ts.Count - 1];
                double t0 = now - WindowMs;
                double mn = double.MaxValue, mx = double.MinValue;
                foreach (double v in vs) { if (v < mn) mn = v; if (v > mx) mx = v; }
                if (IsBool) { mn = 0; mx = 1; }
                else if (mx - mn < 1e-12)
                {
                    double pad = Math.Abs(mx) > 0 ? Math.Abs(mx) * 0.1 : 1;
                    mn -= pad; mx += pad;
                }
                string sMax = IsBool ? "TRUE" : Num(mx);
                string sMin = IsBool ? "FALSE" : Num(mn);
                if (!IsBool) { double pad = (mx - mn) * 0.08; mn -= pad; mx += pad; }

                float fh = font.GetHeight(g);
                float lw = Math.Max(g.MeasureString(sMax, font).Width, g.MeasureString(sMin, font).Width);
                float L = lw + 8, R = 10, T = fh / 2 + 4, B = fh + 6;
                float pw = w - L - R, ph = h - T - B;
                if (pw < 20 || ph < 10) return;

                for (int i = 0; i <= 2; i++) { float y = T + ph * i / 2f; g.DrawLine(gp, L, y, L + pw, y); }
                for (int s = 0; s <= 10; s += 2)
                {
                    float x = L + pw * s / 10f;
                    g.DrawLine(gp, x, T, x, T + ph);
                    string lab = s == 10 ? "now" : "-" + (10 - s) + "s";
                    SizeF sz = g.MeasureString(lab, font);
                    g.DrawString(lab, font, tb, Math.Max(0, Math.Min(x - sz.Width / 2, w - sz.Width)), T + ph + 2);
                }
                // value labels sit on the padded range's 8% lines, i.e. the real min / max
                float yMax = T + ph * 0.08f / 1.16f, yMin = T + ph - ph * 0.08f / 1.16f;
                if (IsBool) { yMax = T; yMin = T + ph; }
                g.DrawString(sMax, font, tb, L - 4 - g.MeasureString(sMax, font).Width, yMax - fh / 2);
                g.DrawString(sMin, font, tb, L - 4 - g.MeasureString(sMin, font).Width, yMin - fh / 2);

                List<PointF> pts = new List<PointF>();
                for (int i = 0; i < vs.Count; i++)
                {
                    float x = (float)(L + (ts[i] - t0) / WindowMs * pw);
                    if (x < L) x = L;
                    float y = (float)(T + ph - (vs[i] - mn) / (mx - mn) * ph);
                    if (IsBool && pts.Count > 0) pts.Add(new PointF(x, pts[pts.Count - 1].Y)); // step
                    pts.Add(new PointF(x, y));
                }
                if (Running) pts.Add(new PointF(L + pw, pts[pts.Count - 1].Y)); // hold last value up to "now"

                g.SetClip(new RectangleF(L, T - 3, pw + 4, ph + 6));
                using (Pen lp = new Pen(line, 2f))
                {
                    lp.LineJoin = LineJoin.Round;
                    if (pts.Count > 1) g.DrawLines(lp, pts.ToArray());
                }
                PointF last = pts[pts.Count - 1];
                using (Brush lb = new SolidBrush(line)) g.FillEllipse(lb, last.X - 3, last.Y - 3, 6, 6);
                g.ResetClip();
            }
        }
    }
}
'@

function Start-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    if (-not ('AdsMini.Spark' -as [type])) {
        $refs = @('System.Drawing')
        if ($PSVersionTable.PSEdition -eq 'Core') { $refs = @('System.Drawing.Common', 'System.Drawing.Primitives', 'System.Collections', 'System.Runtime', 'System.Runtime.Extensions', 'System.Runtime.InteropServices', 'System.Globalization', 'netstandard') } # PowerShell 7
        Add-Type -TypeDefinition $uiSource -ReferencedAssemblies $refs -Language CSharp
    }
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $dll = [AdsMini.AdsClient]::LoadDll($AdsDll)

    $startDark = $Theme -eq 'Dark'
    if ($Theme -eq 'Auto') {
        try {
            $v = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction Stop
            $startDark = ($v.AppsUseLightTheme -eq 0)
        }
        catch { }
    }

    # Shared mutable state lives in a hashtable so event handlers can change it.
    $S = @{
        Client    = (New-AdsClient)
        Results   = (New-Object 'System.Collections.Generic.List[AdsMini.SymbolInfo]')
        Busy      = $false
        LocalId   = $null
        Spark     = (New-Object AdsMini.Spark)
        ChartName = ''
        Tick      = 0
        Dark      = $startDark
        Pal       = $null
    }

    function C($hex) { [System.Drawing.ColorTranslator]::FromHtml($hex) }
    function Get-Palette($dark) {
        if ($dark) {
            return @{ Back = (C '#1F1F1F'); Input = (C '#2A2A2A'); Header = (C '#2D2D2D'); Button = (C '#333333'); Hover = (C '#3E3E3E')
                      Border = (C '#454545'); Text = (C '#E4E4E4'); Dim = (C '#9D9D9D'); Grid = (C '#3A3A3A'); Accent = (C '#4CA6FF')
                      Comment = (C '#7FB96A'); Error = (C '#F28B82') }
        }
        return @{ Back = [System.Drawing.SystemColors]::Control; Input = [System.Drawing.SystemColors]::Window; Header = (C '#F3F3F3')
                  Button = [System.Drawing.SystemColors]::Control; Hover = (C '#E5F1FB'); Border = (C '#C8C8C8')
                  Text = [System.Drawing.SystemColors]::ControlText; Dim = (C '#6E6E6E'); Grid = (C '#E3E3E3'); Accent = (C '#0067C0')
                  Comment = (C '#2E7D32'); Error = (C '#B3261E') }
    }
    $S.Pal = Get-Palette $startDark

    $fontUi = New-Object System.Drawing.Font('Segoe UI', 9)
    $fontSmall = New-Object System.Drawing.Font('Segoe UI', 8)
    $fontMono = New-Object System.Drawing.Font('Consolas', 10)

    function New-Pad($l, $t, $r, $b) { New-Object System.Windows.Forms.Padding($l, $t, $r, $b) }
    function New-Label($text, $tag) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.AutoSize = $true; $l.Margin = (New-Pad 3 8 3 3); $l.Tag = $tag
        return $l
    }
    function New-Button($text, $width) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text; $b.Width = $width; $b.Height = 28
        return $b
    }
    function New-Check($text, $checked) {
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Text = $text; $c.AutoSize = $true; $c.Checked = $checked; $c.Margin = (New-Pad 8 7 3 3)
        return $c
    }
    function New-TextBox($mono) {
        $t = New-Object System.Windows.Forms.TextBox
        if ($mono) { $t.Font = $fontMono }
        return $t
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mini ADS Tool'
    $form.Size = New-Object System.Drawing.Size(1200, 780)
    $form.MinimumSize = New-Object System.Drawing.Size(900, 560)
    $form.StartPosition = 'CenterScreen'
    $form.Font = $fontUi

    # --- connection bar -----------------------------------------------------
    $top = New-Object System.Windows.Forms.FlowLayoutPanel
    $top.Dock = 'Top'; $top.AutoSize = $true; $top.WrapContents = $false; $top.Padding = (New-Pad 6 6 6 4)
    $cbNet = New-Object System.Windows.Forms.ComboBox
    $cbNet.DropDownStyle = 'DropDown'   # pick a route or type a Net ID
    $cbNet.Width = 330; $cbNet.DropDownWidth = 460; $cbNet.MaxDropDownItems = 20; $cbNet.Margin = (New-Pad 3 5 3 3)
    $tbPort = New-TextBox $false; $tbPort.Width = 55; $tbPort.Text = [string]$Port; $tbPort.Margin = (New-Pad 3 5 3 3)
    $btnConnect = New-Button 'Connect' 90
    $btnDisconnect = New-Button 'Disconnect' 90; $btnDisconnect.Enabled = $false
    $btnReload = New-Button 'Reload symbols' 115; $btnReload.Enabled = $false
    $cbDark = New-Check 'Dark' $startDark
    $lblState = New-Label 'Not connected'; $lblState.Margin = (New-Pad 14 8 3 3)
    foreach ($ctl in @((New-Label 'AMS Net ID'), $cbNet, (New-Label 'Port'), $tbPort, $btnConnect, $btnDisconnect, $btnReload, $cbDark, $lblState)) {
        [void]$top.Controls.Add($ctl)
    }

    # --- status line ----------------------------------------------------------
    $status = New-Object System.Windows.Forms.Label
    $status.Dock = 'Bottom'; $status.AutoSize = $false; $status.Height = 24; $status.TextAlign = 'MiddleLeft'
    $status.Padding = (New-Pad 6 0 6 0); $status.AutoEllipsis = $true

    function Set-Status($msg, [switch]$Err) {
        $status.Text = (Get-Date -Format 'HH:mm:ss') + '   ' + $msg
        $status.Tag = $(if ($Err) { 'error' } else { '' })
        if ($Err) { $status.ForeColor = $S.Pal.Error } else { $status.ForeColor = $S.Pal.Text }
    }

    # --- list views: owner-drawn headers so they follow the theme ---------------
    function New-ListView {
        $lv = New-Object System.Windows.Forms.ListView
        $lv.View = 'Details'; $lv.FullRowSelect = $true; $lv.HideSelection = $false; $lv.Dock = 'Fill'
        $lv.BorderStyle = 'None'; $lv.OwnerDraw = $true
        $lv.add_DrawColumnHeader({
            param($sender, $e)
            $P = $S.Pal
            $g = $e.Graphics
            $br = New-Object System.Drawing.SolidBrush($P.Header)
            $g.FillRectangle($br, $e.Bounds)
            $br.Dispose()
            $pen = New-Object System.Drawing.Pen($P.Border)
            $g.DrawLine($pen, $e.Bounds.Left, $e.Bounds.Bottom - 1, $e.Bounds.Right, $e.Bounds.Bottom - 1)
            $g.DrawLine($pen, $e.Bounds.Right - 1, $e.Bounds.Top + 4, $e.Bounds.Right - 1, $e.Bounds.Bottom - 5)
            $pen.Dispose()
            $r = New-Object System.Drawing.Rectangle(($e.Bounds.X + 6), $e.Bounds.Y, ($e.Bounds.Width - 10), $e.Bounds.Height)
            [System.Windows.Forms.TextRenderer]::DrawText($g, $e.Header.Text, $sender.Font, $r, $P.Dim,
                [System.Windows.Forms.TextFormatFlags]'VerticalCenter, Left, EndEllipsis, SingleLine')
        })
        $lv.add_DrawItem({ param($sender, $e) $e.DrawDefault = $true })
        $lv.add_DrawSubItem({ param($sender, $e) $e.DrawDefault = $true })
        return $lv
    }
    # Stretch one column so the columns exactly fill the width (no unthemed strip on the right).
    function Fit-Column($lv, $stretch) {
        $sum = 0
        for ($i = 0; $i -lt $lv.Columns.Count; $i++) { if ($i -ne $stretch) { $sum += $lv.Columns[$i].Width } }
        $w = $lv.ClientSize.Width - $sum - 1
        if ($w -lt 80) { $w = 80 }
        if ($lv.Columns[$stretch].Width -ne $w) { $lv.Columns[$stretch].Width = $w }
    }

    # --- left: search ---------------------------------------------------------
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = 'Fill'; $split.SplitterWidth = 5

    $searchBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $searchBar.Dock = 'Top'; $searchBar.AutoSize = $true; $searchBar.WrapContents = $false
    $tbSearch = New-TextBox $false; $tbSearch.Width = 260; $tbSearch.Margin = (New-Pad 3 5 3 5)
    $cbMembers = New-Check 'Struct / FB members' $true
    $lblCount = New-Label '' 'dim'
    foreach ($ctl in @((New-Label 'Search'), $tbSearch, $cbMembers, $lblCount)) { [void]$searchBar.Controls.Add($ctl) }

    $lvSym = New-ListView
    $lvSym.MultiSelect = $false; $lvSym.VirtualMode = $true
    [void]$lvSym.Columns.Add('Name', 360)
    [void]$lvSym.Columns.Add('Type', 170)
    [void]$lvSym.Columns.Add('Size', 55)

    $split.Panel1.Controls.Add($lvSym)
    $split.Panel1.Controls.Add($searchBar)
    $lvSym.BringToFront()

    # --- right top: variable details (grid layout, scales with DPI) -----------
    $split2 = New-Object System.Windows.Forms.SplitContainer
    $split2.Dock = 'Fill'; $split2.Orientation = 'Horizontal'; $split2.SplitterWidth = 5
    $split.Panel2.Controls.Add($split2)

    $grid = New-Object System.Windows.Forms.TableLayoutPanel
    $grid.Dock = 'Fill'; $grid.ColumnCount = 3; $grid.Padding = (New-Pad 8 6 8 6)
    [void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
    [void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))

    $lblVar = New-Label 'Variable'; $lblVar.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $tbName = New-TextBox $false
    $btnRead = New-Button 'Read' 110
    $lblType = New-Label ''
    $lblAddr = New-Label '' 'dim'
    $lblComment = New-Label '' 'comment'; $lblComment.AutoEllipsis = $true
    $tbValue = New-TextBox $true; $tbValue.ReadOnly = $true
    $cbLive = New-Check 'Live + chart' $false; $cbLive.Anchor = 'Left'
    $tbNew = New-TextBox $true
    $btnWrite = New-Button 'Write' 110
    $cbConfirm = New-Check 'Ask before writing' $true; $cbConfirm.Margin = (New-Pad 0 6 3 3)
    $btnWatch = New-Button 'Add to watch' 110
    $chart = New-Object System.Windows.Forms.PictureBox   # double-buffered, painted by AdsMini.Spark
    $chart.Dock = 'Fill'; $chart.MinimumSize = New-Object System.Drawing.Size(100, 110); $chart.Margin = (New-Pad 3 8 3 3)
    $lblTip = New-Label 'Any path works, e.g. MAIN.aData[7].rValue.   Values: TRUE, 42, 16#FF, 1.5, T#2s, ''text'', E_State.Run, hex:01 02' 'dim'
    $lblTip.AutoEllipsis = $true

    $AnchorLR = [System.Windows.Forms.AnchorStyles]'Left, Right'
    function Add-Cell($ctl, $col, $row, $span) {
        $grid.Controls.Add($ctl, $col, $row)
        if ($span -gt 1) { $grid.SetColumnSpan($ctl, $span) }
    }
    function Add-Row($label, $ctl, $extra) {
        $row = $grid.RowCount
        $grid.RowCount = $row + 1
        [void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
        if ($label) {
            $l = New-Label $label
            $l.Anchor = 'Left'; $l.Margin = (New-Pad 3 3 10 3)
            Add-Cell $l 0 $row 1
        }
        if ($ctl) {
            if ($ctl -is [System.Windows.Forms.TextBox] -or $ctl -is [System.Windows.Forms.Label]) { $ctl.Anchor = $AnchorLR }
            if ($ctl -is [System.Windows.Forms.Label]) { $ctl.AutoSize = $false; $ctl.Height = 22; $ctl.TextAlign = 'MiddleLeft' }
            if ($ctl -is [System.Windows.Forms.TextBox]) { $ctl.Margin = (New-Pad 3 4 3 4) }
            Add-Cell $ctl 1 $row $(if ($extra) { 1 } else { 2 })
        }
        if ($extra) { Add-Cell $extra 2 $row 1 }
    }
    $grid.RowCount = 0
    Add-Row $null $lblVar $null
    Add-Row 'Name' $tbName $btnRead
    Add-Row 'Type' $lblType $null
    Add-Row 'Address' $lblAddr $null
    Add-Row 'Comment' $lblComment $null
    Add-Row 'Value' $tbValue $cbLive
    Add-Row 'New value' $tbNew $btnWrite
    Add-Row $null $cbConfirm $btnWatch
    $chartRow = $grid.RowCount
    $grid.RowCount = $chartRow + 2
    [void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
    [void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
    Add-Cell $chart 0 $chartRow 3
    $lblTip.AutoSize = $false; $lblTip.Height = 22; $lblTip.Anchor = $AnchorLR
    Add-Cell $lblTip 0 ($chartRow + 1) 3
    $split2.Panel1.Controls.Add($grid)

    # --- right bottom: watch list ----------------------------------------------
    $watchBar = New-Object System.Windows.Forms.FlowLayoutPanel
    $watchBar.Dock = 'Top'; $watchBar.AutoSize = $true; $watchBar.WrapContents = $false; $watchBar.Padding = (New-Pad 6 2 6 2)
    $cbWatchLive = New-Check 'Live update (500 ms)' $false
    $btnRemove = New-Button 'Remove' 80
    $btnClear = New-Button 'Clear' 80
    $lblWatch = New-Label 'Watch list'; $lblWatch.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    foreach ($ctl in @($lblWatch, $cbWatchLive, $btnRemove, $btnClear)) { [void]$watchBar.Controls.Add($ctl) }

    $lvWatch = New-ListView
    [void]$lvWatch.Columns.Add('Name', 280)
    [void]$lvWatch.Columns.Add('Value', 180)
    [void]$lvWatch.Columns.Add('Type', 110)
    $split2.Panel2.Controls.Add($lvWatch)
    $split2.Panel2.Controls.Add($watchBar)
    $lvWatch.BringToFront()

    $form.Controls.Add($split)
    $form.Controls.Add($top)
    $form.Controls.Add($status)
    $split.BringToFront()

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 100
    $searchTimer = New-Object System.Windows.Forms.Timer
    $searchTimer.Interval = 250

    # --- theme ------------------------------------------------------------------
    function Set-ThemeOn($c) {
        $P = $S.Pal
        if ($c -is [System.Windows.Forms.TextBox] -or $c -is [System.Windows.Forms.ComboBox] -or $c -is [System.Windows.Forms.ListView]) {
            $c.BackColor = $P.Input; $c.ForeColor = $P.Text
            if ($c -is [System.Windows.Forms.TextBox]) { $c.BorderStyle = $(if ($S.Dark) { 'FixedSingle' } else { 'Fixed3D' }) }
            if ($c -is [System.Windows.Forms.ComboBox]) { $c.FlatStyle = $(if ($S.Dark) { 'Flat' } else { 'Standard' }) }
        }
        elseif ($c -is [System.Windows.Forms.Button]) {
            $c.ForeColor = $P.Text
            if ($S.Dark) {
                $c.FlatStyle = 'Flat'; $c.BackColor = $P.Button
                $c.FlatAppearance.BorderColor = $P.Border; $c.FlatAppearance.MouseOverBackColor = $P.Hover
            }
            else { $c.FlatStyle = 'Standard'; $c.BackColor = $P.Button; $c.UseVisualStyleBackColor = $true }
        }
        elseif ($c -is [System.Windows.Forms.PictureBox]) { $c.BackColor = $P.Input }
        else {
            $c.BackColor = $P.Back
            switch ("$($c.Tag)") {
                'dim' { $c.ForeColor = $P.Dim }
                'comment' { $c.ForeColor = $P.Comment }
                'error' { $c.ForeColor = $P.Error }
                default { $c.ForeColor = $P.Text }
            }
        }
        foreach ($child in $c.Controls) { Set-ThemeOn $child }
    }

    function Set-Theme($dark) {
        $S.Dark = $dark
        $S.Pal = Get-Palette $dark
        $form.SuspendLayout()
        Set-ThemeOn $form
        $split.BackColor = $S.Pal.Border; $split2.BackColor = $S.Pal.Border   # splitter bars
        $split.Panel1.BackColor = $S.Pal.Back; $split.Panel2.BackColor = $S.Pal.Back
        $split2.Panel1.BackColor = $S.Pal.Back; $split2.Panel2.BackColor = $S.Pal.Back
        $form.ResumeLayout()
        [AdsMini.Win]::DarkTitleBar($form.Handle, $dark)
        foreach ($lv in @($lvSym, $lvWatch)) { [AdsMini.Win]::Theme($lv.Handle, $(if ($dark) { 'DarkMode_Explorer' } else { 'Explorer' })) }
        [AdsMini.Win]::Theme($cbNet.Handle, $(if ($dark) { 'DarkMode_CFD' } else { $null }))
        foreach ($it in $lvWatch.Items) { if ("$($it.Tag)" -ne 'error') { $it.ForeColor = $S.Pal.Text } else { $it.ForeColor = $S.Pal.Error } }
        $form.Refresh()
    }

    # --- behaviour ------------------------------------------------------------
    function Set-Connected($on) {
        $btnDisconnect.Enabled = $on
        $btnReload.Enabled = $on
    }

    # Refills the Net ID dropdown, keeping whatever is typed in the box.
    function Update-RouteList {
        $keep = $cbNet.Text
        if ($null -eq $S.LocalId) {
            $S.LocalId = ''
            try { $S.LocalId = [AdsMini.AdsClient]::LocalNetId() } catch { }
        }
        $routes = Get-AdsRoutes
        $cbNet.BeginUpdate()
        $cbNet.Items.Clear()
        [void]$cbNet.Items.Add(('Local (this PC)   ' + $S.LocalId).Trim())
        foreach ($r in $routes) { [void]$cbNet.Items.Add((Format-Route $r)) }
        $cbNet.EndUpdate()
        $cbNet.Text = $keep
        return @($routes).Count
    }

    function Update-State {
        try {
            $dev = ''
            try { $dev = '   ' + $S.Client.DeviceInfo() } catch { }
            $lblState.Text = '{0}   {1}{2}' -f $S.Client.Target, $S.Client.ReadState(), $dev
        }
        catch { $lblState.Text = $S.Client.Target + '   ' + (Get-ErrMsg $_) }
    }

    function Update-Search {
        $lvSym.SelectedIndices.Clear()
        $S.Results = $S.Client.Search($tbSearch.Text, $cbMembers.Checked)
        $lvSym.VirtualListSize = $S.Results.Count
        $lvSym.Invalidate()
        $lblCount.Text = '{0:N0} of {1:N0}' -f $S.Results.Count, $S.Client.Symbols.Count
        Fit-Column $lvSym 0
    }

    function Update-Symbols {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Set-Status 'Uploading symbols...'
            [System.Windows.Forms.Application]::DoEvents()
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $n = $S.Client.LoadSymbols()
            $msg = 'Loaded {0:N0} entries ({1:N0} symbols, {2:N0} types) in {3} ms' -f $n, $S.Client.TopLevelCount, $S.Client.TypeCount, $sw.ElapsedMilliseconds
            if ($S.Client.Truncated) { $msg += '. List truncated; typing a full path in Name still works.' }
            Set-Status $msg
            Update-Search
        }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }

    function Show-Read($r) {
        $i = $r.Info
        $lblType.Text = '{0}   ({1} bytes)' -f $i.Type, $i.Size
        $lblAddr.Text = $i.Address
        $lblComment.Text = ($i.Comment -replace '\s+', ' ')
        if ($tbValue.Text -ne $r.Value) { $tbValue.Text = $r.Value }

        # chart: restart when the variable changes, then add this sample if it's numeric
        if ($S.ChartName -ne $i.Name) {
            $S.Spark.Clear()
            $S.ChartName = $i.Name
            $S.Spark.IsBool = $S.Client.IsBoolType($i)
        }
        $v = $S.Client.NumericValue($i, $r.Raw)
        if (-not [double]::IsNaN($v)) { $S.Spark.Add($v) }
        $chart.Invalidate()
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
            $tbValue.Text = ''
            $cbLive.Checked = $false
            Set-Status (Get-ErrMsg $_) -Err
        }
    }

    function Write-Current {
        $name = $tbName.Text.Trim()
        $val = $tbNew.Text
        if (-not $name) { return }
        if (-not $S.Client.Connected) { Set-Status 'Not connected' -Err; return }
        if ($cbConfirm.Checked) {
            $ans = [System.Windows.Forms.MessageBox]::Show("Write`n`n    $val`n`nto $name ?", 'Confirm write', 'YesNo', 'Question')
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
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
        foreach ($it in $lvWatch.Items) { if ($it.Text -eq $name) { return } }
        $it = New-Object System.Windows.Forms.ListViewItem($name)
        [void]$it.SubItems.Add('')
        [void]$it.SubItems.Add('')
        [void]$lvWatch.Items.Add($it)
        Update-Watch
    }

    function Update-Watch {
        if (-not $S.Client.Connected) { return }
        foreach ($it in $lvWatch.Items) {
            try {
                $r = $S.Client.ReadValue($it.Text)
                if ($it.SubItems[1].Text -ne $r.Value) { $it.SubItems[1].Text = $r.Value }
                if ($it.SubItems[2].Text -ne $r.Info.Type) { $it.SubItems[2].Text = $r.Info.Type }
                if ("$($it.Tag)" -eq 'error') { $it.Tag = ''; $it.ForeColor = $S.Pal.Text }
            }
            catch {
                $it.SubItems[1].Text = Get-ErrMsg $_
                $it.Tag = 'error'
                $it.ForeColor = $S.Pal.Error
                if ($script:FatalCodes -contains (Get-AdsCode $_)) {
                    $cbWatchLive.Checked = $false
                    Set-Status (Get-ErrMsg $_) -Err
                    break
                }
            }
        }
    }

    $btnConnect.add_Click({
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = 0
            if (-not [int]::TryParse($tbPort.Text.Trim(), [ref]$p)) { throw 'Port must be a number (TC3 PLC = 851, TC2 PLC = 801)' }
            $cbLive.Checked = $false
            $S.Client.Connect((Resolve-NetIdText $cbNet.Text), $p, $TimeoutMs)
            Set-Connected $true
            Update-State
            Update-Symbols
            [void]$tbSearch.Focus()
        }
        catch {
            $S.Client.Disconnect()
            Set-Connected $false
            $lblState.Text = 'Not connected'
            Set-Status (Get-ErrMsg $_) -Err
            [void][System.Windows.Forms.MessageBox]::Show((Get-ErrMsg $_), 'Connect failed', 'OK', 'Error')
        }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    # Re-read the routes each time the list opens, so ones added meanwhile show up.
    $cbNet.add_DropDown({ [void](Update-RouteList) })
    $cbNet.add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $e.SuppressKeyPress = $true; $btnConnect.PerformClick() }
    })

    $btnDisconnect.add_Click({
        $cbLive.Checked = $false
        $cbWatchLive.Checked = $false
        $S.Client.Disconnect()
        $S.Client.Symbols.Clear()
        Update-Search
        Set-Connected $false
        $lblState.Text = 'Not connected'
        Set-Status 'Disconnected'
    })

    $btnReload.add_Click({
        try { Update-State; Update-Symbols } catch { Set-Status (Get-ErrMsg $_) -Err }
    })

    $cbDark.add_CheckedChanged({ Set-Theme $cbDark.Checked })

    $tbSearch.add_TextChanged({ $searchTimer.Stop(); $searchTimer.Start() })
    $searchTimer.add_Tick({ $searchTimer.Stop(); Update-Search })
    $cbMembers.add_CheckedChanged({ Update-Search })

    $lvSym.add_RetrieveVirtualItem({
        param($sender, $e)
        $sym = $S.Results[$e.ItemIndex]
        $item = New-Object System.Windows.Forms.ListViewItem($sym.Name)
        [void]$item.SubItems.Add($sym.Type)
        [void]$item.SubItems.Add([string]$sym.Size)
        $e.Item = $item
    })
    $lvSym.add_SelectedIndexChanged({
        if ($lvSym.SelectedIndices.Count -gt 0) {
            $sym = $S.Results[$lvSym.SelectedIndices[0]]
            $tbName.Text = $sym.Name
            Read-Current
        }
    })
    $lvSym.add_ItemActivate({
        if ($lvSym.SelectedIndices.Count -gt 0) { Add-Watch $S.Results[$lvSym.SelectedIndices[0]].Name }
    })
    $lvSym.add_Resize({ Fit-Column $lvSym 0 })
    $lvWatch.add_Resize({ Fit-Column $lvWatch 1 })

    $tbName.add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $e.SuppressKeyPress = $true; Read-Current }
    })
    $tbNew.add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $e.SuppressKeyPress = $true; Write-Current }
    })
    $btnRead.add_Click({ Read-Current })
    $btnWrite.add_Click({ Write-Current })
    $btnWatch.add_Click({ Add-Watch $tbName.Text })
    $cbLive.add_CheckedChanged({ $chart.Invalidate() })

    $chart.add_Paint({
        param($sender, $e)
        $P = $S.Pal
        $S.Spark.Running = $cbLive.Checked
        $S.Spark.Paint($e.Graphics, $chart.ClientSize.Width, $chart.ClientSize.Height, $fontSmall,
            $P.Input, $P.Grid, $P.Accent, $P.Dim, 'Tick "Live + chart" to plot the last 10 s of a number or BOOL')
    })
    $chart.add_Resize({ $chart.Invalidate() })

    $lvWatch.add_SelectedIndexChanged({
        if ($lvWatch.SelectedItems.Count -gt 0) {
            $tbName.Text = $lvWatch.SelectedItems[0].Text
            Read-Current -Quiet
        }
    })
    $btnRemove.add_Click({
        $sel = @($lvWatch.SelectedItems)
        foreach ($it in $sel) { $lvWatch.Items.Remove($it) }
    })
    $btnClear.add_Click({ $lvWatch.Items.Clear() })

    # 100 ms tick: the Live variable (and chart) every tick, the watch list every 5th.
    $timer.add_Tick({
        if ($S.Busy -or -not $S.Client.Connected) { return }
        $S.Busy = $true
        try {
            $S.Tick++
            if ($cbLive.Checked) { Read-Current -Quiet }
            if ($cbWatchLive.Checked -and ($S.Tick % 5 -eq 0)) { Update-Watch }
        }
        finally { $S.Busy = $false }
    })

    $form.add_Shown({
        $split.SplitterDistance = [Math]::Min(620, [int]($form.ClientSize.Width * 0.5))
        $split2.SplitterDistance = [int]($split2.Height * 0.62)
        Fit-Column $lvSym 0
        Fit-Column $lvWatch 1
        $n = Update-RouteList
        if ($NetId) { $cbNet.Text = $NetId } else { $cbNet.SelectedIndex = 0 }
        Set-Status ('Ready. {0} route(s) found. TcAdsDll: {1}' -f $n, $dll)
        [void]$btnConnect.Focus()
    })
    $form.add_FormClosing({
        $timer.Stop()
        $S.Client.Dispose()
    })

    Set-Theme $startDark   # before showing, so there's no light flash
    $timer.Start()
    [void]$form.ShowDialog()
    $timer.Dispose()
    $searchTimer.Dispose()
    $form.Dispose()
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
