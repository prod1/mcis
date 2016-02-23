using System;
using System.Collections.Generic;
using System.Text;
using iTunesLib;
using System.Runtime.InteropServices;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Encapsulates the iTunes COM object and
    /// handles read/write operations
    /// </summary>
    class iTunesTrack
    {
        /// <summary>
        /// Read-only fields
        /// </summary>

        public readonly string Key = string.Empty;
        public readonly string Pathname = string.Empty;
        public readonly string Fullpathname = string.Empty;
        public readonly bool AudioPodcast = false;

        /// <summary>
        /// Read-write fields
        /// </summary>

        protected string _name = string.Empty;
        protected string _artist = string.Empty;
        protected string _album = string.Empty;
        protected string _albumArtist = string.Empty;
        protected string _composer = string.Empty;
        protected string _genre = string.Empty;
        protected string _comment = string.Empty;
        protected string _grouping = string.Empty;
        protected string _description = string.Empty;
        protected int _discNumber = 0;
        protected int _trackNumber = 0;
        protected int _discCount = 0;
        protected int _trackCount = 0;
        protected int _year = 0;
        protected string _lyrics = string.Empty;
        protected bool _compilation = false;
        protected int _bpm = 0;
        protected int _rating = 0;
        protected DateTime _lastPlayed = App.NULLDATE;
        protected int _playCount = 0;
        protected int _skipCount = 0;
        protected DateTime _lastSkipped = App.NULLDATE;
        protected string _sortName = string.Empty;
        protected string _sortArtist = string.Empty;
        protected string _sortAlbumArtist = string.Empty;
        protected string _sortAlbum = string.Empty;
        protected string _sortComposer = string.Empty;

        /// <summary>
        /// The underlying COM object
        /// </summary>
        private IITFileOrCDTrack _track;

        /// <summary>
        /// Whether the rating has changed
        /// </summary>
        private bool _ratingChanged = false;

        /// <summary>
        /// Modifiers for read-write fields
        /// </summary>

        #region Modifiers
        public string Name
        {
            get { return _name; }
            set
            {
                _track.Name = value;
                _name = value;
            }
        }

        public string Artist
        {
            get { return _artist; }
            set
            {
                _track.Artist = value;
                _artist = value;
            }
        }

        public string Album
        {
            get { return _album; }
            set
            {
                _track.Album = value;
                _album = value;
            }
        }

        public string AlbumArtist
        {
            get { return _albumArtist; }
            set
            {
                _track.AlbumArtist = value;
                _albumArtist = value;
            }
        }

        public string Composer
        {
            get { return _composer; }
            set
            {
                _track.Composer = value;
                _composer = value;
            }
        }

        public string Genre
        {
            get { return _genre; }
            set
            {
                _track.Genre = value;
                _genre = value;
            }
        }

        public string Comment
        {
            get { return _comment; }
            set
            {
                _track.Comment = value;
                _comment = value;
            }
        }

        public string Grouping
        {
            get { return _grouping; }
            set
            {
                _track.Grouping = value;
                _grouping = value;
            }
        }

        public string Description
        {
            get { return _description; }
            set
            {
                _track.Description = value;
                _description = value;
            }
        }

        public int DiscNumber
        {
            get { return _discNumber; }
            set
            {
                _track.DiscNumber = value;
                _discNumber = value;
            }
        }

        public int TrackNumber
        {
            get { return _trackNumber; }
            set
            {
                _track.TrackNumber = value;
                _trackNumber = value;
            }
        }

        public int DiscCount
        {
            get { return _discCount; }
            set
            {
                _track.DiscCount = value;
                _discCount = value;
            }
        }

        public int TrackCount
        {
            get { return _trackCount; }
            set
            {
                _track.TrackCount = value;
                _trackCount = value;
            }
        }

        public int Year
        {
            get { return _year; }
            set
            {
                _track.Year = value;
                _year = value;
            }
        }

        public string Lyrics
        {
            get { return _lyrics; }
            set
            {
                _track.Lyrics = value;
                _lyrics = value;
            }
        }

        public bool Compilation
        {
            get { return _compilation; }
            set
            {
                _track.Compilation = value;
                _compilation = value;
            }
        }

        public int BPM
        {
            get { return _bpm; }
            set
            {
                _track.BPM = value;
                _bpm = value;
            }
        }

        public int Rating
        {
            get { return (_rating == 1) ? 0 : _rating; }
            set
            {
                int val = (value == 0) ? 1 : value;
                _track.Rating = val;
                _ratingChanged = true;
                _rating = val;
            }
        }

        public DateTime LastPlayed
        {
            get { return _lastPlayed; }
            set
            {
                if (Synchronizer.iTunesVersion.Major >= 10)
                    _track.PlayedDate = value.ToUniversalTime();
                else
                    _track.PlayedDate = value;

                _lastPlayed = value;
            }
        }

        public int PlayCount
        {
            get { return _playCount; }
            set
            {
                _track.PlayedCount = value;
                _playCount = value;
            }
        }

        public int SkipCount
        {
            get { return _skipCount; }
            set
            {
                _track.SkippedCount = value;
                _skipCount = value;
            }
        }

        public DateTime LastSkipped
        {
            get { return _lastSkipped; }
            set
            {
                if (Synchronizer.iTunesVersion.Major >= 10)
                    _track.SkippedDate = value.ToUniversalTime();
                else
                    _track.SkippedDate = value;

                _lastSkipped = value;
            }
        }

        public bool HasArtwork
        {
            get
            {
                IITArtworkCollection artworkCollection = _track.Artwork;
                if (artworkCollection != null)
                {
                    foreach (IITArtwork artwork in artworkCollection)
                    {
                        if (artwork == null)
                            continue;
                        else
                            return true;
                    }
                }

                return false;
            }
        }

        public string Artwork
        {
            set { _track.AddArtworkFromFile(value); }
        }

        public int AlbumRating
        {
            get
            {
                return _track.AlbumRating;
            }
        }

        public bool AlbumRatingCalculated
        {
            get
            {
                return (_track.AlbumRatingKind == ITRatingKind.ITRatingKindComputed);
            }
        }

        public bool RatingChanged { get { return _ratingChanged; } }

        public string SortName
        {
            get { return _sortName; }
            set
            {
                _track.SortName = value;
                _sortName = value;
            }
        }

        public string SortArtist
        {
            get { return _sortArtist; }
            set
            {
                _track.SortArtist = value;
                _sortArtist = value;
            }
        }

        public string SortAlbumArtist
        {
            get { return _sortAlbumArtist; }
            set
            {
                _track.SortAlbumArtist = value;
                _sortAlbumArtist = value;
            }
        }

        public string SortAlbum
        {
            get { return _sortAlbum; }
            set
            {
                _track.SortAlbum = value;
                _sortAlbum = value;
            }
        }

        public string SortComposer
        {
            get { return _sortComposer; }
            set
            {
                _track.SortComposer = value;
                _sortComposer = value;
            }
        }

        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="track">The underlying COM object that this object will encapsulate</param>
        /// <param name="flags">Which fields to assign</param>
        public iTunesTrack(IITFileOrCDTrack track, ulong flags)
        {
            _track = track;

            if ((flags & App.NAME) > 0)
            {
                _name = _track.Name;
                if (_name == null)
                    _name = string.Empty;
            }

            if ((flags & App.ARTIST) > 0)
            {
                _artist = _track.Artist;
                if (_artist == null)
                    _artist = string.Empty;
            }

            if ((flags & App.ALBUM) > 0)
            {
                _album = _track.Album;
                if (_album == null)
                    _album = string.Empty;
            }

            if ((flags & App.ALBUMARTIST) > 0)
            {
                _albumArtist = _track.AlbumArtist;
                if (_albumArtist == null)
                    _albumArtist = string.Empty;
            }

            if ((flags & App.COMPOSER) > 0)
            {
                _composer = _track.Composer;
                if (_composer == null)
                    _composer = string.Empty;
            }

            if ((flags & App.GENRE) > 0)
            {
                _genre = _track.Genre;
                if (_genre == null)
                    _genre = string.Empty;
            }

            if ((flags & App.COMMENT) > 0)
            {
                _comment = _track.Comment;
                if (_comment == null)
                    _comment = string.Empty;
            }

            if ((flags & App.GROUPING) > 0)
            {
                _grouping = _track.Grouping;
                if (_grouping == null)
                    _grouping = string.Empty;
            }

            if ((flags & App.DESCRIPTION) > 0)
            {
                _description = _track.Description;
                if (_description == null)
                    _description = string.Empty;
            }

            if ((flags & (App.DISCNUMBER | App.DISCCOUNT | App.TRACKNUMBER | App.TRACKCOUNT)) > 0)
            {
                _discNumber = _track.DiscNumber;
                _discCount = _track.DiscCount;
                _trackNumber = _track.TrackNumber;
                _trackCount = _track.TrackCount;
            }

            if ((flags & App.YEAR) > 0)
                _year = _track.Year;

            if ((flags & App.LYRICS) > 0)
            {
                try { _lyrics = _track.Lyrics; }
                catch {/* Do nothing */}
                if (_lyrics == null)
                    _lyrics = string.Empty;
            }

            if ((flags & App.COMPILATION) > 0)
                _compilation = _track.Compilation;

            if ((flags & App.BPM) > 0)
                _bpm = _track.BPM;

            if ((flags & App.RATING) > 0)
            {
                int currentRating = _track.Rating;
                if (_track.ratingKind == ITRatingKind.ITRatingKindUser && currentRating > 0)
                    _rating = currentRating;
                else
                {
                    _ratingChanged = true;
                    _track.Rating = 1;
                    _rating = 1;
                }
            }

            if ((flags & App.LASTPLAYED) > 0)
                _lastPlayed = _track.PlayedDate;

            if ((flags & App.PLAYCOUNT) > 0)
                _playCount = _track.PlayedCount;

            if ((flags & App.SKIPCOUNT) > 0)
                _skipCount = _track.SkippedCount;

            if ((flags & App.LASTSKIPPED) > 0)
                _lastSkipped = _track.SkippedDate;

            Fullpathname = _track.Location;
            if (Fullpathname == null)
                Fullpathname = string.Empty;

            if(Fullpathname != string.Empty)
                Pathname = Fullpathname.Substring(0, Fullpathname.LastIndexOf('\\') + 1);

            Key = Fullpathname.ToUpper();

            if ((flags & App.ADDARTWORKTOPODCASTS) > 0)
                AudioPodcast = (_track.Podcast && _track.VideoKind == ITVideoKind.ITVideoKindNone);

            if ((flags & App.SORTNAME) > 0)
            {
                _sortName = _track.SortName;
                if (_sortName == null)
                    _sortName = string.Empty;
            }

            if ((flags & App.SORTARTIST) > 0)
            {
                _sortArtist = _track.SortArtist;
                if (_sortArtist == null)
                    _sortArtist = string.Empty;
            }

            if ((flags & App.SORTALBUMARTIST) > 0)
            {
                _sortAlbumArtist = _track.SortAlbumArtist;
                if (_sortAlbumArtist == null)
                    _sortAlbumArtist = string.Empty;
            }

            if ((flags & App.SORTALBUM) > 0)
            {
                _sortAlbum = _track.SortAlbum;
                if (_sortAlbum == null)
                    _sortAlbum = string.Empty;
            }

            if ((flags & App.SORTCOMPOSER) > 0)
            {
                _sortComposer = _track.SortComposer;
                if (_sortComposer == null)
                    _sortComposer = string.Empty;
            }
        }
    }
}
