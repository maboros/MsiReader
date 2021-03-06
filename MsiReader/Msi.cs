using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MsiReader
{
    public static class Msi
    {
        [DllImport("msi.dll", SetLastError = true)]
        static extern uint MsiOpenDatabase(string szDatabasePath, IntPtr phPersist, out IntPtr phDatabase);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiDatabaseOpenView(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr phView);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiViewExecute(IntPtr hView, IntPtr hRecord);

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern uint MsiViewFetch(IntPtr hView, out IntPtr hRecord);


        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        static extern int MsiRecordGetString(IntPtr hRecord, int iField, StringBuilder szValueBuf, ref int pcchValueBuf);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordDataSize(IntPtr hRecord, int iField);

        public const int MSICOLINFO_NAMES = 0;  
        public const int MSICOLINFO_TYPES = 1;  
        [DllImport("msi.dll", ExactSpelling = true)]
        static extern uint MsiViewGetColumnInfo(IntPtr hView, int eColumnInfo, out IntPtr hRecord);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordGetFieldCount(IntPtr hRecord);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern int MsiRecordGetInteger(IntPtr hRecord, int iField);

        [DllImport("msi.dll", ExactSpelling = true)]
        static extern uint MsiRecordReadStream(IntPtr hRecord, uint iField,[Out] byte[] szDataBuf, ref int pcbDataBuf);

        public static uint OpenDatabase(string szDatabasePath, IntPtr phPersist, out IntPtr phDatabase)
        {
            return MsiOpenDatabase(szDatabasePath, phPersist, out phDatabase);
        }
        public static int DatabaseOpenView(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr hView)
        {
            return MsiDatabaseOpenView(hDatabase,szQuery, out hView);
        }
        public static int ViewExecute(IntPtr hView, IntPtr hRecord)
        {
            return MsiViewExecute(hView, hRecord);
        }
        public static uint ViewFetch(IntPtr hView, out IntPtr hRecord)
        {
            return MsiViewFetch(hView,out hRecord);
        }
        public static int RecordGetString(IntPtr hRecord, int iField, StringBuilder szValueBuf, ref int pcchValueBuf)
        {
            return MsiRecordGetString(hRecord,iField,szValueBuf,ref pcchValueBuf);
        }
        public static uint ViewGetColumnInfo(IntPtr hView, int eColumnInfo, out IntPtr hRecord)
        {
            return MsiViewGetColumnInfo(hView,eColumnInfo,out hRecord);
        }
        public static int RecordGetFieldCount(IntPtr hRecord)
        {
            return MsiRecordGetFieldCount(hRecord);
        }
        public static int RecordDataSize(IntPtr hRecord, int iField)
        {
            return MsiRecordDataSize(hRecord,iField);
        }
        public static int RecordGetInteger(IntPtr hRecord, int iField)
        {
            return MsiRecordGetInteger(hRecord,iField);
        }
        public static uint RecordReadStream(IntPtr hRecord, uint iField, [Out] byte[] szDataBuf, ref int pcbDataBuf)
        {
            return MsiRecordReadStream(hRecord, iField, szDataBuf, ref pcbDataBuf);
        }
    }
}
