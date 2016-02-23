using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Contains any globally accessible constants or methods
    /// </summary>
    static class App
    {
        /// <summary>
        /// The log filename
        /// </summary>
        public static readonly string LOGFILENAME = AppDomain.CurrentDomain.BaseDirectory + "MCiTunesSynchronizerSyncLog.xml";

        /// <summary>
        /// The settings filename
        /// </summary>
        public static string SETTINGSFILENAME = AppDomain.CurrentDomain.BaseDirectory + "MCiTunesSynchronizer.settings";

        /// <summary>
        /// Where to get updates info
        /// </summary>
        public static string UPDATESURL = "http://www.misterpete.co.uk/files/prodversions.xml";

        /// <summary>
        /// The name of the On-The-Go playlist folder in MC
        /// </summary>
        public const string ONTHEGOPATH = "On-The-Go";

        /// <summary>
        /// The date used as an empty or null date
        /// </summary>
        public static readonly DateTime NULLDATE = new DateTime(1899, 12, 30, 0, 0, 0);

        /// <summary>
        /// Field flags
        /// </summary>

        public const ulong NAME                     = 0x1;
        public const ulong ARTIST                   = 0x2;
        public const ulong ALBUM                    = 0x4;
        public const ulong ALBUMARTIST              = 0x8;
        public const ulong COMPOSER                 = 0x10;
        public const ulong GENRE                    = 0x20;
        public const ulong COMMENT                  = 0x40;
        public const ulong GROUPING                 = 0x80;
        public const ulong DESCRIPTION              = 0x100;
        public const ulong DISCNUMBER               = 0x200;
        public const ulong TRACKNUMBER              = 0x400;
        public const ulong YEAR                     = 0x800;
        public const ulong LYRICS                   = 0x1000;
        public const ulong RATING                   = 0x2000;
        public const ulong LASTPLAYED               = 0x4000;
        public const ulong PLAYCOUNT                = 0x8000;
        public const ulong SKIPCOUNT                = 0x10000;
        public const ulong LASTSKIPPED              = 0x20000;
        public const ulong COMPILATION              = 0x40000;
        public const ulong BPM                      = 0x80000;
        public const ulong DISCCOUNT                = 0x100000;
        public const ulong TRACKCOUNT               = 0x200000;
        public const ulong ALBUMRATING              = 0x400000;
        public const ulong PLAYLISTS                = 0x800000;
        public const ulong SHOWHIDDENPLAYLISTS      = 0x1000000;
        public const ulong SORTNAME                 = 0x2000000;
        public const ulong SORTARTIST               = 0x4000000;
        public const ulong SORTALBUMARTIST          = 0x8000000;
        public const ulong SORTALBUM                = 0x10000000;
        public const ulong SORTCOMPOSER             = 0x20000000;
        public const ulong ONTHEGOPLAYLISTS         = 0x40000000;
        public const ulong USEPLAYLISTROOT          = 0x80000000;
        public const ulong MANAGETICKS              = 0x100000000;

        /// <summary>
        /// Settings flags
        /// </summary>

        public const ulong ADDARTWORKTOPODCASTS     = 0x1000000000;
        public const ulong REFRESHPODCASTS          = 0x2000000000;
        public const ulong SYNCHRONIZEIPOD          = 0x4000000000;
        public const ulong QUITITUNES               = 0x8000000000;
        public const ulong USEALTERNATIVEFILEKEY    = 0x10000000000;
        public const ulong TESTMODE                 = 0x20000000000;
        public const ulong MINIMISEITUNES           = 0x40000000000;
        public const ulong IMPORTFILESTOITUNES      = 0x80000000000;
        public const ulong SYNCCERTAINFILETYPES     = 0x100000000000;
        public const ulong REMOVEBROKENLINKS        = 0x200000000000;
        public const ulong SYNCHRONIZEIPODBEFORE    = 0x400000000000;
        public const ulong REPLACEFORWARDSLASHES    = 0x800000000000;

        // Log flags
        public const ulong NOLOG                    = 0x10000000000000;
        public const ulong LOGTOCLIPBOARD           = 0x20000000000000;
        public const ulong LOGFILENOTFOUND          = 0x40000000000000;
        public const ulong LOGTRACKCHANGES          = 0x80000000000000;         
        public const ulong LOGIMPORTFILESUCCESS     = 0x100000000000000;         
        public const ulong LOGIMPORTFILENOSUCCESS   = 0x200000000000000;         
        public const ulong LOGREMOVEBROKENLINK      = 0x400000000000000;         
        public const ulong LOGFILEKEYWARNING        = 0x800000000000000;         
        public const ulong LOGPLAYLISTCREATED       = 0x1000000000000000;         
        public const ulong LOGPLAYLISTDELETED       = 0x2000000000000000;         
        public const ulong LOGPLAYLISTADDFILE       = 0x4000000000000000;         
        public const ulong LOGPLAYLISTDELETEFILE    = 0x8000000000000000;

        /// <summary>
        /// All the field flags; 36 available bits
        /// </summary>
        public const ulong FIELDFLAGS               = 0x0000000FFFFFFFFF;
        /// <summary>
        /// All the about box flags; 28 available bits
        /// </summary>
        public const ulong ABOUTBOXFLAGS            = 0xFFFFFFF000000000;

        /// <summary>
        /// Internal flags
        /// </summary>
        
        /// <remarks>Whether to sync a playlist</remarks>
        public const uint PLAYLIST_SELECTED         = 0x1;
        /// <remarks>Whether to rebuild a playlist from scratch</remarks>
        public const uint PLAYLIST_REBUILD          = 0x2;

        /// <summary>
        /// Attaches a tool tip to a control
        /// </summary>
        /// <param name="control">The control to which to attach the tool tip</param>
        /// <param name="balloon">Whether the tooltip should be the balloon style</param>
        internal static void SetToolTip(Control control, bool balloon)
        {
            ToolTip tip = new ToolTip();
            tip.IsBalloon = balloon;
            tip.ToolTipTitle = control.Text;
            tip.AutoPopDelay = 30000;
            tip.SetToolTip(control, control.Tag.ToString());
            tip.ShowAlways = true;
        }
    }

    /// <summary>
    /// Media Center specific commands/constants/enums
    /// </summary>
    static class MC
    {
        /// <summary>
        /// The date MC uses as a NULL (empty) date when writing to a field
        /// </summary>
        public static readonly DateTime BASEDATE = new DateTime(1970, 1, 1, 0, 0, 0);

        /// <summary>
        /// Command instructor
        /// </summary>
        public const uint WM_MC_COMMAND = 33768;

        /// <summary>
        /// MC Commands
        /// </summary>

        public const uint MCC_REFRESH = 22007;
        public const uint MCC_LOAD_LIBRARY = 20025;

        /// <summary>
        /// Playlist types
        /// </summary>
        public enum PLAYLIST_TYPES
        {
            PLAYLIST,
            PLAYLIST_GROUP,
            SMARTLIST,
            SERVICELIST,
        };

        /// <summary>
        /// MC playlist IDs
        /// </summary>

        private const int PLAYLIST_ID_SYSTEM_BASE                = -1000;
        public const int PLAYLIST_ID_SYSTEM_TOP_HITS            = (PLAYLIST_ID_SYSTEM_BASE - 1);
        public const int PLAYLIST_ID_SYSTEM_RECENTLY_PLAYED     = (PLAYLIST_ID_SYSTEM_BASE - 2);
        public const int PLAYLIST_ID_SYSTEM_RECENTLY_IMPORTED   = (PLAYLIST_ID_SYSTEM_BASE - 3);
        public const int PLAYLIST_ID_SYSTEM_RECENTLY_RIPPED     = (PLAYLIST_ID_SYSTEM_BASE - 4);
        public const int PLAYLIST_ID_SYSTEM_PURCHASED           = (PLAYLIST_ID_SYSTEM_BASE - 5);
    }
}
