using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace BibleCommon.Helpers
{
    public static class SystemUtils
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        public enum MachineType : ushort
        {
            IMAGE_FILE_MACHINE_UNKNOWN = 0x0,
            IMAGE_FILE_MACHINE_AM33 = 0x1d3,
            IMAGE_FILE_MACHINE_AMD64 = 0x8664,
            IMAGE_FILE_MACHINE_ARM = 0x1c0,
            IMAGE_FILE_MACHINE_EBC = 0xebc,
            IMAGE_FILE_MACHINE_I386 = 0x14c,
            IMAGE_FILE_MACHINE_IA64 = 0x200,
            IMAGE_FILE_MACHINE_M32R = 0x9041,
            IMAGE_FILE_MACHINE_MIPS16 = 0x266,
            IMAGE_FILE_MACHINE_MIPSFPU = 0x366,
            IMAGE_FILE_MACHINE_MIPSFPU16 = 0x466,
            IMAGE_FILE_MACHINE_POWERPC = 0x1f0,
            IMAGE_FILE_MACHINE_POWERPCFP = 0x1f1,
            IMAGE_FILE_MACHINE_R4000 = 0x166,
            IMAGE_FILE_MACHINE_SH3 = 0x1a2,
            IMAGE_FILE_MACHINE_SH3DSP = 0x1a3,
            IMAGE_FILE_MACHINE_SH4 = 0x1a6,
            IMAGE_FILE_MACHINE_SH5 = 0x1a8,
            IMAGE_FILE_MACHINE_THUMB = 0x1c2,
            IMAGE_FILE_MACHINE_WCEMIPSV2 = 0x169,
        }

        public static bool TouchInputAvailable()
        {
            var NID_READY = 0x80;
            var NID_MULTI_INPUT = 0x40;
            var SM_DIGITIZER = 94;
            var value = GetSystemMetrics(SM_DIGITIZER);

            if ((value & NID_READY) == NID_READY)                 // stack ready 
            {
                if ((value & NID_MULTI_INPUT) == NID_MULTI_INPUT)           // digitizer is multitouch 
                {
                    return true;
                }
            }

            return false;
        }

        public static Encoding GetFileEncoding(string filePath)
        {
            System.Text.Encoding result = null;
            using (FileStream fs = new System.IO.FileStream(filePath,
                FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (fs.CanSeek)
                {
                    byte[] bom = new byte[4]; // Get the byte-order mark, if there is one 
                    fs.Read(bom, 0, 4);
                    if ((bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
                        || (bom[0] == 47 && bom[1] == 47 && bom[2] == 32 && bom[3] == 208)
                        || (bom[0] == 60 && bom[1] == 109 && bom[2] == 101 && bom[3] == 116)
                        || (bom[0] == 60 && bom[1] == 116 && bom[2] == 105 && bom[3] == 116))  // utf-8 
                    {
                        result = System.Text.Encoding.UTF8;
                    }
                    else if ((bom[0] == 0xff && bom[1] == 0xfe)   // ucs-2le, ucs-4le, and ucs-16le 
                        || (bom[0] == 0xfe && bom[1] == 0xff) // utf-16 and ucs-2 
                        || (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff)) // ucs-4 
                    {
                        result = System.Text.Encoding.Unicode;
                    }
                    else
                    {
                        result = System.Text.Encoding.Default;
                    }

                    // Now reposition the file cursor back to the start of the file 
                    fs.Seek(0, System.IO.SeekOrigin.Begin);
                }
                else
                {
                    // The file cannot be randomly accessed, so you need to decide what to set the default to 
                    // based on the data provided. If you're expecting data from a lot of older applications, 
                    // default your encoding to Encoding.ASCII. If you're expecting data from a lot of newer 
                    // applications, default your encoding to Encoding.Unicode. Also, since binary files are 
                    // single byte-based, so you will want to use Encoding.ASCII, even though you'll probably 
                    // never need to use the encoding then since the Encoding classes are really meant to get 
                    // strings from the byte array that is the file. 

                    result = System.Text.Encoding.Default;
                }
            }

            return result;
        }

        public static string GetOneNoteProgramFilePath()
        {
            var registryPath = "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\onenote.exe";
            
            var key = Registry.CurrentUser.OpenSubKey(registryPath, false);
            if (key == null)
                key = Registry.LocalMachine.OpenSubKey(registryPath, false);
            if (key == null)
                throw new Exception("it does not exist!");

            return key.GetValue(string.Empty).ToString();
        }

        // returns true if the dll is 64-bit, false if 32-bit, and null if unknown
        public static bool? UnmanagedDllIs64Bit(string dllPath)
        {
            try
            {
                switch (GetDllMachineType(dllPath))
                {
                    case MachineType.IMAGE_FILE_MACHINE_AMD64:
                    case MachineType.IMAGE_FILE_MACHINE_IA64:
                        return true;
                    case MachineType.IMAGE_FILE_MACHINE_I386:
                        return false;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            return null;
        }


        public static MachineType GetDllMachineType(string dllPath)
        {
            //see http://www.microsoft.com/whdc/system/platform/firmware/PECOFF.mspx
            //offset to PE header is always at 0x3C
            //PE header starts with "PE\0\0" =  0x50 0x45 0x00 0x00
            //followed by 2-byte machine type field (see document above for enum)
            using (var fs = new FileStream(dllPath, FileMode.Open, FileAccess.Read))
            {
                using (var br = new BinaryReader(fs))
                {
                    fs.Seek(0x3c, SeekOrigin.Begin);
                    Int32 peOffset = br.ReadInt32();
                    fs.Seek(peOffset, SeekOrigin.Begin);
                    UInt32 peHead = br.ReadUInt32();
                    if (peHead != 0x00004550) // "PE\0\0", little-endian
                        throw new Exception("Can't find PE header");
                    return (MachineType)br.ReadUInt16();
                }
            }
        }
    }
}
