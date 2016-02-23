using System;
using System.Xml;
using System.Collections.Generic;
using iTunesLib;
using MediaCenter;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// The valid change types
    /// </summary>
    public enum ChangeType { None, iTunes, MC };

    /// <summary>
    /// Monitors changes to tracks so we can log them
    /// </summary>
    class TrackMonitor
    {
        /// <summary>
        /// The iTunes track being monitored
        /// </summary>
        public iTunesTrack iTunesTrack;
        /// <summary>
        /// The MC track being monitored
        /// </summary>
        public MCTrack MCTrack;

        /// <summary>
        /// Field names
        /// </summary>
        public List<string> FieldNames = new List<string>();
        /// <summary>
        /// Change types
        /// </summary>
        public List<ChangeType> ChangeTypes = new List<ChangeType>();
        /// <summary>
        /// Old values
        /// </summary>
        public List<object> OldValues = new List<object>();
        /// <summary>
        /// New values
        /// </summary>
        public List<object> NewValues = new List<object>();
        /// <summary>
        /// Whether any of the track values have changed
        /// </summary>
        private bool _trackChanged = false;

        /// <summary>
        /// Whether any of the track values have been changed
        /// </summary>
        public bool Changed
        {
            get
            {
                return _trackChanged;
            }
        }

        /// <summary>
        /// Constructs a TrackMonitor object
        /// </summary>
        /// <param name="itunesTrack">The iTunes track we are monitoring</param>
        /// <param name="mcTrack">The MC track we are monitoring</param>
        public TrackMonitor(iTunesTrack itunesTrack, MCTrack mcTrack)
        {
            iTunesTrack = itunesTrack;
            MCTrack = mcTrack;
        }

        #region The Set methods

        private void SetField(string fieldName, ChangeType changeType, object oldValue, object newValue, bool replace)
        {
            if (replace)
            {
                // If the fieldname already exists, replace it with the new values
                for (int i = 0; i < FieldNames.Count; i++)
                {
                    if (FieldNames[i] == fieldName)
                    {
                        ChangeTypes[i] = changeType;
                        OldValues[i] = (oldValue == null) ? string.Empty : oldValue;
                        NewValues[i] = (newValue == null) ? string.Empty : newValue;
                        return;
                    }
                }
            }

            // If we get this far replace must be false or we couldn't find an existing fieldname
            _trackChanged = true;
            FieldNames.Add(fieldName);
            ChangeTypes.Add(changeType);
            OldValues.Add((oldValue == null) ? string.Empty : oldValue);
            NewValues.Add((newValue == null) ? string.Empty : newValue);
        }

        public void SetName(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Name", changeType, oldValue, newValue, false);
        }

        public void SetArtist(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Artist", changeType, oldValue, newValue, false);
        }

        public void SetAlbum(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Album", changeType, oldValue, newValue, false);
        }

        public void SetAlbumArtist(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("AlbumArtist", changeType, oldValue, newValue, false);
        }

        public void SetComposer(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Composer", changeType, oldValue, newValue, false);
        }

        public void SetGenre(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Genre", changeType, oldValue, newValue, false);
        }

        public void SetComment(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Comment", changeType, oldValue, newValue, false);
        }

        public void SetGrouping(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Grouping", changeType, oldValue, newValue, false);
        }

        public void SetDescription(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Description", changeType, oldValue, newValue, false);
        }

        public void SetDiscNumber(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("DiscNumber", changeType, oldValue, newValue, false);
        }

        public void SetTrackNumber(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("TrackNumber", changeType, oldValue, newValue, false);
        }

        public void SetYear(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("Year", changeType, oldValue, newValue, false);
        }

        public void SetLyrics(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Lyrics", changeType, oldValue, newValue, false);
        }

        public void SetRating(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("Rating", changeType, oldValue, newValue, false);
        }

        public void SetAlbumRating(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("AlbumRating", changeType, oldValue, newValue, true);
        }

        public void SetAlbumRatingCalculated(ChangeType changeType, bool oldValue, bool newValue)
        {
            SetField("AlbumRatingCalculated", changeType, oldValue, newValue, true);
        }

        public void SetLastPlayed(ChangeType changeType, DateTime oldValue, DateTime newValue)
        {
            SetField("LastPlayed", changeType, oldValue, newValue, false);
        }

        public void SetPlayCount(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("PlayCount", changeType, oldValue, newValue, false);
        }

        public void SetSkipCount(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("SkipCount", changeType, oldValue, newValue, false);
        }

        public void SetLastSkipped(ChangeType changeType, DateTime oldValue, DateTime newValue)
        {
            SetField("LastSkipped", changeType, oldValue, newValue, false);
        }

        public void SetCompilation(ChangeType changeType, bool oldValue, bool newValue)
        {
            SetField("Compilation", changeType, oldValue, newValue, false);
        }

        public void SetBPM(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("BPM", changeType, oldValue, newValue, false);
        }

        public void SetDiscCount(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("DiscCount", changeType, oldValue, newValue, false);
        }

        public void SetTrackCount(ChangeType changeType, int oldValue, int newValue)
        {
            SetField("TrackCount", changeType, oldValue, newValue, false);
        }

        public void SetError(string message)
        {
            SetField("Error", ChangeType.None, null, message, false);
        }

        public void SetError(string field, string message)
        {
            SetField("Error", ChangeType.None, field, message, false);
        }

        public void SetArtwork(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("Artwork", changeType, oldValue, newValue, false);
        }

        public void SetSortName(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("SortName", changeType, oldValue, newValue, false);
        }

        public void SetSortArtist(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("SortArtist", changeType, oldValue, newValue, false);
        }

        public void SetSortAlbumArtist(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("SortAlbumArtist", changeType, oldValue, newValue, false);
        }

        public void SetSortAlbum(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("SortAlbum", changeType, oldValue, newValue, false);
        }

        public void SetSortComposer(ChangeType changeType, string oldValue, string newValue)
        {
            SetField("SortComposer", changeType, oldValue, newValue, false);
        }

        #endregion
    }
}
