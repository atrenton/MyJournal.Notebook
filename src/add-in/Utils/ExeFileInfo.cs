using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using DWORD = System.UInt32;
using WORD = System.UInt16;

namespace MyJournal.Notebook.Utils
{
    class ExeFileInfo
    {
#if WIN32
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool Wow64RevertWow64FsRedirection(IntPtr ptr);
#endif
        #region Properties

        /// <summary>
        /// Common Object File Format (COFF) header for this executable.
        /// </summary>
        internal IMAGE_FILE_HEADER COFFHeader { get; private set; }

        /// <summary>
        /// File properties for this executable.
        /// </summary>
        internal FileInfo FileInfo { get; private set; }

        /// <summary>
        /// Image Type (32 vs 64-bit) for this executable.
        /// </summary>
        internal string ImageType
        {
            get
            {
                return Is32Bit() ? "32-bit" : "64-bit";
            }
        }

        /// <summary>
        /// Machine Type (x86 vs x64) for this executable.
        /// </summary>
        internal string MachineType
        {
            get
            {
                return Is32Bit() ? "x86" : "x64";
            }
        }

        /// <summary>
        /// Date and time this executable was created by the linker.
        /// </summary>
        internal DateTimeOffset TimeDateStamp
        {
            get
            {
                var seconds = (long)COFFHeader.TimeDateStamp;
                return DateTimeOffset.FromUnixTimeSeconds(seconds);
            }
        }

        #endregion

        internal ExeFileInfo(string filePath)
        {
            FileInfo = new FileInfo(filePath);
            if (FileInfo.Extension.ToLower(CultureInfo.CurrentCulture) != ".exe")
            {
                const string Msg = "Not an Executable file";
                throw new ArgumentException(Msg, nameof(filePath));
            }
#if WIN32
            var ptr = new IntPtr();
            var isWow64FsRedirectionDisabled = Wow64DisableWow64FsRedirection(ref ptr);
#endif
            var fs = new FileStream(FileInfo.FullName, FileMode.Open, FileAccess.Read);
            InitCOFFHeader(fs);

#if WIN32
            if (isWow64FsRedirectionDisabled)
            {
                Wow64RevertWow64FsRedirection(ptr);
            }
#endif
        }

        void InitCOFFHeader(FileStream fs)
        {

            IMAGE_FILE_HEADER header;
            WORD magic;
            using (var reader = new BinaryReader(fs))
            {
                if (reader.ReadUInt16() != DOS_HEADER_SIGNATURE)
                {
                    const string Msg = "Invalid DOS Header";
                    throw new BadImageFormatException(Msg, FileInfo.Name);
                }
                fs.Position = PE_POINTER_OFFSET;
                fs.Position = reader.ReadUInt32();
                if (reader.ReadUInt32() != COFF_HEADER_SIGNATURE)
                {
                    const string Msg = "Invalid COFF Header";
                    throw new BadImageFormatException(Msg, FileInfo.Name);
                }
                header.Machine = reader.ReadUInt16();
                header.NumberOfSections = reader.ReadUInt16();
                header.TimeDateStamp = reader.ReadUInt32();
                header.PointerToSymbolTable = reader.ReadUInt32();
                header.NumberOfSymbols = reader.ReadUInt32();
                header.SizeOfOptionalHeader = reader.ReadUInt16();
                header.Characteristics = reader.ReadUInt16();
                magic = reader.ReadUInt16();
            }
            COFFHeader = header;
        }

        internal bool Is32Bit()
        {
            return (COFFHeader.Machine == IMAGE_FILE_MACHINE_I386);
        }

        internal bool Is64Bit()
        {
            var cpu = COFFHeader.Machine;
            return ((cpu == IMAGE_FILE_MACHINE_AMD64) ||
                    (cpu == IMAGE_FILE_MACHINE_IA64));
        }

        /// <summary>
        /// Returns multi-line file header values for this executable.
        /// Format is the same as the Windows SDK DUMPBIN utility.
        /// </summary>
        public override string ToString()
        {
            const string HEX = "X";
            const int WIDTH = 11;
            object[] args = {
                FileInfo.ToString(),
                COFFHeader.Machine.ToString(HEX).PadLeft(WIDTH),
                MachineType,
                COFFHeader.NumberOfSections.ToString(HEX).PadLeft(WIDTH),
                COFFHeader.TimeDateStamp.ToString(HEX).PadLeft(WIDTH),
                TimeDateStamp.LocalDateTime.ToString("ddd MMM dd HH:mm:ss yyyy"),
                COFFHeader.PointerToSymbolTable.ToString(HEX).PadLeft(WIDTH),
                COFFHeader.NumberOfSymbols.ToString(HEX).PadLeft(WIDTH),
                COFFHeader.SizeOfOptionalHeader.ToString(HEX).PadLeft(WIDTH),
                COFFHeader.Characteristics.ToString(HEX).PadLeft(WIDTH)
            };
            string[] format = {
                "EXE FILE INFO: {0}",
                "FILE HEADER VALUES",
                "{1} machine ({2})",
                "{3} number of sections",
                "{4} time date stamp {5}",
                "{6} file pointer to symbol table",
                "{7} number of symbols",
                "{8} size of optional header",
                "{9} characteristics"
            };
            return string.Format(string.Join(Environment.NewLine, format), args);
        }

        #region Constants

        /// <summary>
        /// Common Object File Format (COFF) header.
        /// SEE: https://msdn.microsoft.com/en-us/library/windows/desktop/ms680313(v=vs.85).aspx 
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct IMAGE_FILE_HEADER
        {
#pragma warning disable IDE1006 // Naming Styles
            internal WORD Machine;
            internal WORD NumberOfSections;
            internal DWORD TimeDateStamp;
            internal DWORD PointerToSymbolTable;
            internal DWORD NumberOfSymbols;
            internal WORD SizeOfOptionalHeader;
            internal WORD Characteristics;
#pragma warning restore IDE1006 // Naming Styles
        }

        internal const ushort
          IMAGE_FILE_MACHINE_UNKNOWN = 0,
          IMAGE_FILE_MACHINE_I386 = 0x014c,  // x86
          IMAGE_FILE_MACHINE_IA64 = 0x0200,  // Intel Itanium
          IMAGE_FILE_MACHINE_AMD64 = 0x8664; // x64

        const DWORD COFF_HEADER_SIGNATURE = 0x4550;
        const WORD DOS_HEADER_SIGNATURE = 0x5A4D;
        const long PE_POINTER_OFFSET = 0x3C;

        #endregion
    }
}
