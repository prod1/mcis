using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using MediaCenter;
using iTunesLib;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Controls logging
    /// </summary>
    static class Logger
    {
        /// <summary>
        /// Configuration settings
        /// </summary>
        static private ulong _flags = 0;
        /// <summary>
        /// The logger XML writer
        /// </summary>
        static private XmlWriter _logWriter = null;
        /// <summary>
        /// The underlying file stream
        /// </summary>
        static private Stream _stream = null;
        /// <summary>
        /// Whether the log is enabled
        /// </summary>
        static private bool _enabled = true;

        /// <summary>
        /// Creates the logger from a MemoryStream
        /// </summary>
        static public void Create()
        {
            // Create the stream
            _stream = new MemoryStream();
            // Set up the log
            Initialize();
        }

        /// <summary>
        /// Creates the logger from a FileStream
        /// </summary>
        /// <param name="stream">The filename of the log</param>
        static public void Create(string filename)
        {
            if (filename == null)
            {
                _enabled = false;
                return;
            }

            // Create the file
            _stream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
            // Set up the log
            Initialize();
        }

        /// <summary>
        /// Sets up the XML log
        /// </summary>
        static private void Initialize()
        {
            if (!_enabled)
                return;

            // Setup the XML log
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.CloseOutput = false;
            _logWriter = XmlWriter.Create(_stream, settings);
        }

        /// <summary>
        /// Writes the log header
        /// </summary>
        /// <param name="flags">Configuration settings</param>
        /// <param name="syncFileTypesList">Sync file types</param>
        /// <param name="importFileTypesList">Import file types</param>
        /// <param name="playlists">Playlists to sync</param>
        /// <param name="rootPlaylistFolder">Root playlist sync folder</param>
        /// <param name="expressions">MC expressions</param>
        /// <param name="exportOnlyMCInfo">Whether to export only MC to iTunes - no aggregation</param>
        /// <param name="exportOnlyiTunesInfo">Whether to export only iTunes to MC - no aggregation</param>
        static public void WriteHeader(ulong flags, List<string> syncFileTypesList, List<string> importFileTypesList, Dictionary<string, PlaylistItem> playlists, string rootPlaylistFolder, Dictionary<string, string> expressions, bool exportOnlyMCInfo, bool exportOnlyiTunesInfo)
        {
            if (!_enabled)
                return;

            _flags = flags;
            _logWriter.WriteStartDocument();
            // Write the root node
            _logWriter.WriteStartElement("MCiS");
            _logWriter.WriteStartAttribute("Version");
            _logWriter.WriteString(AboutBox.AssemblyVersion);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("StartDateTime");
            _logWriter.WriteString(DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Options");
            _logWriter.WriteString(flags.ToString("X"));
            _logWriter.WriteEndAttribute();

            if (exportOnlyMCInfo)
            {
                _logWriter.WriteStartAttribute("ExportAggregatesOnlyFrom");
                _logWriter.WriteString("MC");
                _logWriter.WriteEndAttribute();
            }

            if (exportOnlyiTunesInfo)
            {
                _logWriter.WriteStartAttribute("ExportAggregatesOnlyFrom");
                _logWriter.WriteString("iTunes");
                _logWriter.WriteEndAttribute();
            }

            if ((flags & App.SYNCCERTAINFILETYPES) > 0)
            {
                _logWriter.WriteStartElement("SyncFileTypes");
                foreach (string fileType in syncFileTypesList)
                {
                    _logWriter.WriteStartElement("FileType");
                    _logWriter.WriteString(fileType);
                    _logWriter.WriteEndElement();
                }
                _logWriter.WriteEndElement();
            }

            if ((flags & App.IMPORTFILESTOITUNES) > 0)
            {
                _logWriter.WriteStartElement("ImportFileTypes");
                foreach (string fileType in importFileTypesList)
                {
                    _logWriter.WriteStartElement("FileType");
                    _logWriter.WriteString(fileType);
                    _logWriter.WriteEndElement();
                }
                _logWriter.WriteEndElement();
            }

            if ((flags & App.PLAYLISTS) > 0)
            {
                _logWriter.WriteStartElement("ExportPlaylists");

                _logWriter.WriteStartAttribute("iTunesPlaylistFolder");
                _logWriter.WriteString(((flags & App.USEPLAYLISTROOT) > 0) ? "<root>":rootPlaylistFolder);
                _logWriter.WriteEndAttribute();

                foreach (string playlist in playlists.Keys)
                {
                    PlaylistItem pl = playlists[playlist];

                    if (pl.Selected)
                    {
                        _logWriter.WriteStartElement("Playlist");
                        _logWriter.WriteStartAttribute("Rebuild");
                        _logWriter.WriteString(pl.Rebuild.ToString());
                        _logWriter.WriteEndAttribute();
                        _logWriter.WriteStartAttribute("Shuffle");
                        _logWriter.WriteString(pl.Shuffle.ToString());
                        _logWriter.WriteEndAttribute();
                        _logWriter.WriteStartAttribute("RemoveTracks");
                        _logWriter.WriteString(pl.RemoveTracks.ToString());
                        _logWriter.WriteEndAttribute();
                        _logWriter.WriteString(playlist);
                        _logWriter.WriteEndElement();
                    }
                }
                _logWriter.WriteEndElement();
            }

            _logWriter.WriteStartElement("Expressions");
            foreach (string field in expressions.Keys)
            {
                _logWriter.WriteStartElement("Expression");
                _logWriter.WriteStartAttribute("Field");
                _logWriter.WriteString(field);
                _logWriter.WriteEndAttribute();
                _logWriter.WriteString(expressions[field]);
                _logWriter.WriteEndElement();
            }
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes information about MC
        /// </summary>
        /// <param name="mc">The MC object</param>
        /// <param name="viewSchemeName">The MC view scheme we are using</param>
        /// <param name="playListName">The MC playlist we are using</param>
        static public void WriteMCHeader(MCAutomation mc, string viewSchemeName, string playListName)
        {
            if (!_enabled)
                return;

            string libraryName = string.Empty, libraryPath = string.Empty;

            IMJVersionAutomation ver = mc.GetVersion();
            mc.GetLibrary(ref libraryName, ref libraryPath);

            _logWriter.WriteStartElement("MC");
            _logWriter.WriteStartAttribute("Version");
            _logWriter.WriteString(ver.Version);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("LibraryName");
            _logWriter.WriteString(libraryName);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("LibraryPath");
            _logWriter.WriteString(libraryPath);
            _logWriter.WriteEndAttribute();

            if (viewSchemeName != null && viewSchemeName != string.Empty)
            {
                _logWriter.WriteStartAttribute("ViewScheme");
                _logWriter.WriteString(viewSchemeName);
                _logWriter.WriteEndAttribute();
            }
            else if (playListName != null && playListName != string.Empty)
            {
                _logWriter.WriteStartAttribute("Playlist");
                _logWriter.WriteString(playListName);
                _logWriter.WriteEndAttribute();
            }

            _logWriter.WriteEndElement();
        }

        static public void WriteiTunesHeader(iTunesApp iTunes)
        {
            if (!_enabled)
                return;

            _logWriter.WriteStartElement("iTunes");
            _logWriter.WriteStartAttribute("Version");
            _logWriter.WriteString(iTunes.Version);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Finalises the log file
        /// </summary>
        /// <param name="totals">Results totals</param>
        /// <param name="message">Result final message</param>
        static public void WriteFooter(int[] totals, string message)
        {
            if (!_enabled)
                return;

            _logWriter.WriteStartElement("Result");
            _logWriter.WriteStartAttribute("EndDateTime");
            _logWriter.WriteString(DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
            _logWriter.WriteEndAttribute();

            if (totals != null)
            {
                _logWriter.WriteStartAttribute("Analyzed");
                _logWriter.WriteString(totals[0].ToString());
                _logWriter.WriteEndAttribute();
                _logWriter.WriteStartAttribute("Synchronized");
                _logWriter.WriteString(totals[1].ToString());
                _logWriter.WriteEndAttribute();
                _logWriter.WriteStartAttribute("Imported");
                _logWriter.WriteString(totals[2].ToString());
                _logWriter.WriteEndAttribute();
                _logWriter.WriteStartAttribute("Removed");
                _logWriter.WriteString(totals[3].ToString());
                _logWriter.WriteEndAttribute();
            }

            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Closes the underlying XML writer and stream
        /// </summary>
        /// <param name="returnXml">Whether to return the resulting XML</param>
        static public string Destroy(bool returnXml)
        {
            if (!_enabled)
                return null;

            _logWriter.Close();

            try
            {
                if (returnXml)
                {
                    _stream.Position = 0;
                    StreamReader reader = new StreamReader(_stream);
                    return reader.ReadToEnd();
                }
            }
            finally
            {
                _stream.Dispose();
            }

            return null;
        }

        /// <summary>
        /// Logs an MC caching error
        /// </summary>
        /// <param name="artist">Artist</param>
        /// <param name="album">Album</param>
        /// <param name="name">Track name</param>
        /// <param name="message">Exception message</param>
        static public void WriteMCCacheError(string artist, string album, string name, string message)
        {
            if (!_enabled)
                return;

            _logWriter.WriteStartElement("MCCacheError");
            _logWriter.WriteStartAttribute("Artist");
            _logWriter.WriteString(artist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Album");
            _logWriter.WriteString(album);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Name");
            _logWriter.WriteString(name);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Logs an iTunes caching error
        /// </summary>
        /// <param name="artist">Artist</param>
        /// <param name="album">Album</param>
        /// <param name="name">Track name</param>
        /// <param name="message">Exception message</param>
        static public void WriteiTunesCacheError(string artist, string album, string name, string message)
        {
            if (!_enabled)
                return;

            _logWriter.WriteStartElement("iTunesCacheError");
            _logWriter.WriteStartAttribute("Artist");
            _logWriter.WriteString(artist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Album");
            _logWriter.WriteString(album);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Name");
            _logWriter.WriteString(name);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes the the track XML to the specified logger
        /// </summary>
        static public void WriteTrackMonitor(TrackMonitor tm)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGTRACKCHANGES) == 0)
                return;

            // Setup the element and attributes
            _logWriter.WriteStartElement("Track");

            // Only write one filename if they are the same
            if (tm.MCTrack == null || tm.MCTrack.Fullpathname.Trim() == string.Empty || tm.iTunesTrack.Fullpathname.Equals(tm.MCTrack.Fullpathname, StringComparison.OrdinalIgnoreCase))
            {
                _logWriter.WriteStartAttribute("Filename");
                _logWriter.WriteString(tm.iTunesTrack.Fullpathname);
                _logWriter.WriteEndAttribute();
            }
            else
            {
                // They are different filenames so write them both
                _logWriter.WriteStartAttribute("iTunesFilename");
                _logWriter.WriteString(tm.iTunesTrack.Fullpathname);
                _logWriter.WriteEndAttribute();
                _logWriter.WriteStartAttribute("MCFilename");
                _logWriter.WriteString(tm.MCTrack.Fullpathname);
                _logWriter.WriteEndAttribute();
            }

            // Write each tag change
            for (int i = 0; i < tm.FieldNames.Count; i++)
            {
                WriteTrackMonitorElement(tm.FieldNames[i], tm.ChangeTypes[i], tm.OldValues[i], tm.NewValues[i]);
            }

            // End of the track element
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes an import file operation to the log
        /// </summary>
        /// <param name="filename">Filename of the imported file</param>
        /// <param name="successful">Whether the operation was successful</param>
        /// <param name="message">The message to write</param>
        static public void WriteImportFile(string filename, bool successful, string message)
        {
            if (!_enabled)
                return;

            if (((_flags & App.LOGIMPORTFILESUCCESS) == 0 && successful) || ((_flags & App.LOGIMPORTFILENOSUCCESS) == 0 && !successful))
                return;

            _logWriter.WriteStartElement("ImportToiTunes");
            _logWriter.WriteStartAttribute("Filename");
            _logWriter.WriteString(filename);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a remove broken link operation to the log
        /// </summary>
        /// <param name="artist">Artist</param>
        /// <param name="album">Album</param>
        /// <param name="name">Track name</param>
        /// <param name="message">The message to write</param>
        static public void WriteRemoveBrokenLink(string artist, string album, string name, string message)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGREMOVEBROKENLINK) == 0)
                return;

            _logWriter.WriteStartElement("RemoveBrokenLink");
            _logWriter.WriteStartAttribute("Artist");
            _logWriter.WriteString(artist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Album");
            _logWriter.WriteString(album);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Name");
            _logWriter.WriteString(name);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a file key warning
        /// </summary>
        /// <param name="fileKey">The file key</param>
        /// <param name="message">The message to write</param>
        static public void WriteFileKeyWarning(string fileKey, string message)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGFILEKEYWARNING) == 0)
                return;

            _logWriter.WriteStartElement("FileKeyWarning");
            _logWriter.WriteStartAttribute("FileKey");
            _logWriter.WriteString(fileKey);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteString(message);
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a create playlist operation to the log
        /// </summary>
        /// <param name="playlist">The playlist path and name</param>
        static public void WritePlaylistCreated(string playlist)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGPLAYLISTCREATED) == 0)
                return;

            _logWriter.WriteStartElement("CreatedPlaylist");
            _logWriter.WriteStartAttribute("Playlist");
            _logWriter.WriteString(playlist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a delete playlist operation to the log
        /// </summary>
        /// <param name="playlist">The playlist path and name</param>
        static public void WritePlaylistDeleted(string playlist)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGPLAYLISTDELETED) == 0)
                return;

            _logWriter.WriteStartElement("DeletedPlaylist");
            _logWriter.WriteStartAttribute("Playlist");
            _logWriter.WriteString(playlist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a playlist add file operation to the log
        /// </summary>
        /// <param name="playlist">The playlist path and name</param>
        /// <param name="filename">The filename we are adding</param>
        static public void WritePlaylistAddFile(string playlist, string filename)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGPLAYLISTADDFILE) == 0)
                return;

            _logWriter.WriteStartElement("AddFileToPlaylist");
            _logWriter.WriteStartAttribute("Playlist");
            _logWriter.WriteString(playlist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Filename");
            _logWriter.WriteString(filename);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes a playlist delete file operation to the log
        /// </summary>
        /// <param name="playlist">The playlist path and name</param>
        /// <param name="filename">The filename we are deleting</param>
        static public void WritePlaylistDeleteFile(string playlist, string filename)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGPLAYLISTDELETEFILE) == 0)
                return;

            _logWriter.WriteStartElement("DeleteFileFromPlaylist");
            _logWriter.WriteStartAttribute("Playlist");
            _logWriter.WriteString(playlist);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteStartAttribute("Filename");
            _logWriter.WriteString(filename);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Writes file not found to the log
        /// </summary>
        /// <param name="filename">The filename that was not found</param>
        static public void WriteFileNotFound(string filename)
        {
            if (!_enabled)
                return;

            if ((_flags & App.LOGFILENOTFOUND) == 0)
                return;

            _logWriter.WriteStartElement("FileNotFound");
            _logWriter.WriteStartAttribute("Pathname");
            _logWriter.WriteString(filename);
            _logWriter.WriteEndAttribute();
            _logWriter.WriteEndElement();
        }

        /// <summary>
        /// Utility method for writing the tag change elements
        /// </summary>
        /// <param name="elementName">The name of the element (or tag)</param>
        /// <param name="changeType">The application in which the change was made</param>
        /// <param name="oldValue">The old value</param>
        /// <param name="newValue">The new value</param>
        static private void WriteTrackMonitorElement(string elementName, ChangeType changeType, object oldValue, object newValue)
        {
            if (!_enabled)
                return;

            // Start the element                
            _logWriter.WriteStartElement(elementName);

            if (changeType == ChangeType.None)
            {
                if (elementName == "Error")
                {
                    if (oldValue != null && oldValue.ToString() != string.Empty)
                    {
                        _logWriter.WriteStartAttribute("Field");
                        _logWriter.WriteString(oldValue.ToString());
                        _logWriter.WriteEndAttribute();
                    }

                    _logWriter.WriteStartAttribute("Message");
                    _logWriter.WriteString(newValue.ToString());
                    _logWriter.WriteEndAttribute();
                }
            }
            else
            {
                // Write where the change was made
                _logWriter.WriteStartAttribute("ChangedIn");

                switch (changeType)
                {
                    case ChangeType.iTunes:
                        _logWriter.WriteString("iTunes");
                        break;
                    case ChangeType.MC:
                        _logWriter.WriteString("MC");
                        break;
                }

                // End ChangedIn attribute
                _logWriter.WriteEndAttribute();

                // Write old value
                _logWriter.WriteStartAttribute("OldValue");

                // If the value is a date we want to display it in a specific way
                // otherwise just call the class' own ToString()
                if (oldValue is DateTime)
                {
                    DateTime d = (DateTime)oldValue;
                    _logWriter.WriteString((d == App.NULLDATE) ? "(null)" : d.ToString("yyyy-MM-ddTHH:mm:ss"));
                }
                else
                    _logWriter.WriteString(oldValue.ToString());

                // End OldValue attribute
                _logWriter.WriteEndAttribute();

                // Write new value
                _logWriter.WriteStartAttribute("NewValue");

                // If the value is a date we want to display it in a specific way
                // otherwise just call the class' own ToString()
                if (newValue is DateTime)
                {
                    DateTime d = (DateTime)newValue;
                    _logWriter.WriteString((d == App.NULLDATE) ? "(null)" : d.ToString("yyyy-MM-ddTHH:mm:ss"));
                }
                else
                    _logWriter.WriteString(newValue.ToString());

                // End NewValue attribute
                _logWriter.WriteEndAttribute();
            }

            // End tag element
            _logWriter.WriteEndElement();
        }
    }
}
