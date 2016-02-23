using System;
using System.Collections.Generic;
using System.Text;
using MediaCenter;
using System.Runtime.InteropServices;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Encapsulates the MC COM object and
    /// handles read/write operations
    /// </summary>
    class MCTrack
    {
        /// <summary>
        /// The template to be sent as a request to MC
        /// </summary>
        protected static string _template = string.Empty;

        /// <summary>
        /// Read-only fields
        /// </summary>

        public readonly string Key = string.Empty;
        public readonly string AlbumKey = string.Empty;
        public readonly string Name = string.Empty;
        public readonly string Artist = string.Empty;
        public readonly string Album = string.Empty;
        public readonly string AlbumArtist = string.Empty;
        public readonly string Composer = string.Empty;
        public readonly string Genre = string.Empty;
        public readonly string Comment = string.Empty;
        public readonly string Grouping = string.Empty;
        public readonly string Description = string.Empty;
        public readonly int DiscNumber = 0;
        public readonly int TrackNumber = 0;
        public readonly int Year = 0;
        public readonly string Lyrics = string.Empty;
        public readonly bool Compilation = false;
        public readonly int BPM = 0;
        public readonly string iTunesFileKey = string.Empty;
        public readonly string Pathname = string.Empty;
        public readonly string Filename = string.Empty;
        public readonly string Fullpathname = string.Empty;
        public readonly string FileType = string.Empty;
        public readonly string SortName = string.Empty;
        public readonly string SortArtist = string.Empty;
        public readonly string SortAlbumArtist = string.Empty;
        public readonly string SortAlbum = string.Empty;
        public readonly string SortComposer = string.Empty;

        /// <summary>
        /// Read-write fields
        /// </summary>
        
        protected int _rating = 0;
        protected DateTime _lastPlayed = App.NULLDATE;
        protected int _playCount = 0;
        protected int _skipCount = 0;
        protected DateTime _lastSkipped = App.NULLDATE;
        protected int _albumRating = 0;
        protected bool _albumRatingCalculated = false;
        protected int _iTunesRatingSync = -1;
        protected int _iTunesPlayedCountSync = 0;
        protected int _iTunesSkippedCountSync = 0;

        /// <summary>
        /// The underlying COM object
        /// </summary>
        public IMJFileAutomation _track;

        /// <summary>
        /// The corresponding index in iTunes
        /// </summary>
        public int iTunesIndex = -1;

        /// <summary>
        /// The persistent ID of the track in iTunes
        /// </summary>
        public long PersistentID = 0;

        /// <summary>
        /// The constituents of the persistent ID 
        /// </summary>
        
        public int PersistentIDLow = 0;
        public int PersistentIDHigh = 0;

        /// <summary>
        /// Whether the track is ticked in iTunes
        /// </summary>
        public bool Ticked;
        public bool DoNotTick = false;
        public bool VerifyTick = true;

        /// <summary>
        /// Whether the track has been part of a synchronize operation
        /// </summary>
        public bool Synchronised = false;

        /// <summary>
        /// Modifiers for read-write fields
        /// </summary>
        
        #region Modifiers
        public int Rating
        {
            get { return _rating * 20; }
            set
            {
                int newRating = (value == 0) ? 0 : value / 20;
                _track.Rating = newRating;
                _rating = newRating;
            }
        }

        public DateTime LastPlayed
        {
            get { return _lastPlayed; }
            set
            {
                DateTime lastPlayed = value;

                if (lastPlayed != App.NULLDATE)
                {
                    lastPlayed = value - TimeZone.CurrentTimeZone.GetUtcOffset(value);
                    if (lastPlayed < MC.BASEDATE)
                        lastPlayed = App.NULLDATE;
                }

                if (lastPlayed != _lastPlayed)
                {
                    _track.LastPlayed = (lastPlayed == App.NULLDATE) ? 0:(int)(new TimeSpan(lastPlayed.Ticks).TotalSeconds - new TimeSpan(MC.BASEDATE.Ticks).TotalSeconds);
                    _lastPlayed = lastPlayed;
                }
            }
        }

        public int PlayCount
        {
            get { return _playCount; }
            set
            {
                _track.PlayCounter = value;
                _playCount = value;
            }
        }

        public int SkipCount
        {
            get { return _skipCount; }
            set
            {
                _track.Set("Skip Count", value.ToString());
                _skipCount = value;
            }
        }

        public DateTime LastSkipped
        {
            get { return _lastSkipped; }
            set
            {
                DateTime lastSkipped = value;

                if (lastSkipped != App.NULLDATE)
                {
                    lastSkipped = lastSkipped - TimeZone.CurrentTimeZone.GetUtcOffset(lastSkipped);
                    if (lastSkipped < MC.BASEDATE)
                        lastSkipped = App.NULLDATE;
                }

                if (lastSkipped != _lastSkipped)
                {
                    _track.Set("Last Skipped", ((lastSkipped == App.NULLDATE) ? 0 : (int)(new TimeSpan(lastSkipped.Ticks).TotalSeconds - new TimeSpan(MC.BASEDATE.Ticks).TotalSeconds)).ToString());
                    _lastSkipped = lastSkipped;
                }
            }
        }

        public int AlbumRating
        {
            get { return _albumRating; }
            set
            {
                _track.Set("iTunesAlbumRating", value.ToString());
                int newRating = (value == 0) ? 0 : value / 20;
                _track.Set("Album Rating", newRating.ToString());
                _albumRating = value;
            }
        }

        public bool AlbumRatingCalculated
        {
            get { return _albumRatingCalculated; }
            set
            {
                _track.Set("Album Rating Calculated", (value) ? "1":"0");
                _albumRatingCalculated = value;
            }
        }

        public int iTunesRatingSync
        {
            get { return _iTunesRatingSync; }
            set
            {
                _track.Set("iTunesRatingSync", value.ToString());
                _iTunesRatingSync = value;
            }
        }

        public int iTunesPlayedCountSync
        {
            get { return _iTunesPlayedCountSync; }
            set
            {
                _track.Set("iTunesPlayedCountSync", value.ToString());
                _iTunesPlayedCountSync = value;
            }
        }

        public int iTunesSkippedCountSync
        {
            get { return _iTunesSkippedCountSync; }
            set
            {
                _track.Set("iTunesSkippedCountSync", value.ToString());
                _iTunesSkippedCountSync = value;
            }
        }

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="track">The underlying COM object that this object will encapsulate</param>
        /// <param name="flags">Which fields to assign</param>
        public MCTrack(IMJFileAutomation track, ulong flags)
        {
            _track = track;

            // Get the values
            string[] values = _track.GetFilledTemplate(_template).Split('\v');
            if (values.Length != 34)
                throw new Exception(string.Format("Values length is {0}", values.Length));

            // Assign the values
            Name = values[0];
            Artist = values[1];
            Album = values[2];
            AlbumArtist = values[3];
            Comment = values[4];
            Composer = values[5];
            Description = values[6];

            if (values[7] != string.Empty)
            {
                try { DiscNumber = int.Parse(values[7]); }
                catch {/* Do nothing */}
            }

            Genre = values[8];
            Grouping = values[9];

            if (values[10] != string.Empty)
            {
                try
                {
                    double rawdate = double.Parse(values[10]);
                    _lastPlayed = App.NULLDATE.AddDays(rawdate);
                }
                catch {/* Do nothing */}
            }

            Lyrics = values[11];

            if (values[12] != string.Empty)
            {
                try { _playCount = int.Parse(values[12]); }
                catch {/* Do nothing */}
            }

            if (values[13] != string.Empty)
            {
                try { _rating = int.Parse(values[13]); }
                catch {/* Do nothing */}
            }

            if (values[14] != string.Empty)
            {
                try { TrackNumber = int.Parse(values[14]); }
                catch {/* Do nothing */}
            }

            if (values[15] != string.Empty)
            {
                try { Year = int.Parse(values[15]); }
                catch {/* Do nothing */}
            }

            if (values[16] != string.Empty)
            {
                try { _skipCount = int.Parse(values[16]); }
                catch {/* Do nothing */}
            }

            if (values[17] != string.Empty)
            {
                try
                {
                    double rawdate = double.Parse(values[17]);
                    _lastSkipped = App.NULLDATE.AddDays(rawdate);
                }
                catch {/* Do nothing */}
            }

            Compilation = (values[18] == "1");

            if (values[19] != string.Empty)
            {
                try { BPM = int.Parse(values[19]); }
                catch {/* Do nothing */}
            }

            if (values[20] != string.Empty && values[20] != "[iTunesAlbumRating]")
            {
                try { _albumRating = int.Parse(values[20]); }
                catch {/* Do nothing */}
            }

            if (values[21] != string.Empty && values[21] != "[iTunesRatingSync]")
            {
                try { _iTunesRatingSync = int.Parse(values[21]); }
                catch {/* Do nothing */}
            }

            if (values[22] != string.Empty && values[22] != "[iTunesPlayedCountSync]")
            {
                try { _iTunesPlayedCountSync = int.Parse(values[22]); }
                catch {/* Do nothing */}
            }

            if (values[23] != string.Empty && values[23] != "[iTunesSkippedCountSync]")
            {
                try { _iTunesSkippedCountSync = int.Parse(values[23]); }
                catch {/* Do nothing */}
            }

            if (values[24] != "[iTunesFileKey]")
            {
                iTunesFileKey = values[24].Trim();
            }

            if(values[25] != "[AlbumRatingCalculated]")
                _albumRatingCalculated = (values[25] == "1");

            Pathname = values[26];
            Filename = values[27];
            Fullpathname = Pathname + Filename;
            FileType = values[28].ToUpper();

            Key = (((flags & App.USEALTERNATIVEFILEKEY) > 0 && iTunesFileKey != string.Empty) ? iTunesFileKey:Fullpathname).ToUpper();

            AlbumKey = (Pathname + Album).ToUpper();

            SortName = values[29];
            SortArtist = values[30];
            SortAlbumArtist = values[31];
            SortAlbum = values[32];
            SortComposer = values[33];
        }

        /// <summary>
        /// Refreshes the field mappings according to settings
        /// </summary>
        /// <param name="flags">Which fields to assign</param>
        /// <param name="expressions">The expressions to assign</param>
        static public void RefreshTemplate(ulong flags, Dictionary<string, string> expressions)
        {
            string nameExp = expressions["Name"];
            string artistExp = expressions["Artist"];
            string albumExp = expressions["Album"];
            string albumArtistExp = expressions["Album Artist"];
            string commentExp = expressions["Comment"];
            string composerExp = expressions["Composer"];
            string descriptionExp = expressions["Description"];
            string discNoExp = expressions["Disc #"];
            string trackNoExp = expressions["Track #"];
            string genreExp = expressions["Genre"];
            string groupingExp = expressions["Grouping"];
            string lyricsExp = expressions["Lyrics"];
            string yearExp = expressions["Year"];
            string mixAlbumExp = expressions["Mix Album"];
            string bpmExp = expressions["BPM"];
            string sortNameExp = expressions["Sort Name"];
            string sortArtistExp = expressions["Sort Artist"];
            string sortAlbumArtistExp = expressions["Sort Album Artist"];
            string sortAlbumExp = expressions["Sort Album"];
            string sortComposerExp = expressions["Sort Composer"];

            // Construct the template
            StringBuilder t = new StringBuilder();

            string[] fields = new string[34];

            if ((flags & App.NAME) > 0)
                fields[0] = nameExp;

            if ((flags & App.ARTIST) > 0)
                fields[1] = artistExp;

            if ((flags & (App.ALBUM | App.ALBUMRATING | App.DISCCOUNT | App.TRACKCOUNT)) > 0)
                fields[2] = albumExp;

            if ((flags & App.ALBUMARTIST) > 0)
                fields[3] = albumArtistExp;

            if ((flags & App.COMMENT) > 0)
                fields[4] = commentExp;

            if ((flags & App.COMPOSER) > 0)
                fields[5] = composerExp;

            if ((flags & App.DESCRIPTION) > 0)
                fields[6] = descriptionExp;

            if ((flags & (App.DISCNUMBER | App.DISCCOUNT | App.TRACKNUMBER | App.TRACKCOUNT)) > 0)
            {
                fields[7] = discNoExp;
                fields[14] = trackNoExp;
            }

            if ((flags & App.GENRE) > 0)
                fields[8] = genreExp;

            if ((flags & App.GROUPING) > 0)
                fields[9] = groupingExp;

            if ((flags & App.LASTPLAYED) > 0)
                fields[10] = "[Last Played,0]";

            if ((flags & App.LYRICS) > 0)
                fields[11] = lyricsExp;

            if ((flags & App.PLAYCOUNT) > 0)
            {
                fields[12] = "[Number Plays]";
                fields[22] = "[iTunesPlayedCountSync]";
            }

            if ((flags & App.RATING) > 0)
            {
                fields[13] = "[Rating,0]";
                fields[21] = "[iTunesRatingSync]";
            }

            if ((flags & App.YEAR) > 0)
                fields[15] = yearExp;

            if ((flags & App.SKIPCOUNT) > 0)
            {
                fields[16] = "[Skip Count]";
                fields[23] = "[iTunesSkippedCountSync]";
            }

            if ((flags & App.LASTSKIPPED) > 0)
                fields[17] = "[Last Skipped,0]";

            if ((flags & App.COMPILATION) > 0)
                fields[18] = mixAlbumExp;

            if ((flags & App.BPM) > 0)
                fields[19] = bpmExp;

            if ((flags & App.ALBUMRATING) > 0)
            {
                fields[20] = "[iTunesAlbumRating]";
                fields[25] = "[Album Rating Calculated]";
            }

            fields[24] = "[iTunesFileKey]";
            fields[26] = "[Filename (path)]";
            fields[27] = "[Filename (name)]";
            fields[28] = "[File Type]";

            if ((flags & App.SORTNAME) > 0)
                fields[29] = sortNameExp;

            if ((flags & App.SORTARTIST) > 0)
                fields[30] = sortArtistExp;

            if ((flags & App.SORTALBUMARTIST) > 0)
                fields[31] = sortAlbumArtistExp;

            if ((flags & App.SORTALBUM) > 0)
                fields[32] = sortAlbumExp;

            if ((flags & App.SORTCOMPOSER) > 0)
                fields[33] = sortComposerExp;

            for (int i = 0, max = fields.Length; i < max; i++)
            {
                if (i > 0)
                    t.Append('\v');
                t.Append(fields[i]);
            }

            _template = t.ToString();
        }
    }
}
