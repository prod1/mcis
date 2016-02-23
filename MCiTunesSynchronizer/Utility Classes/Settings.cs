using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Net;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Configuration file maintenance
    /// </summary>
    static class Settings
    {
        /// <summary>
        /// The filename of the configuration file
        /// </summary>
        static private string Filename = App.SETTINGSFILENAME;

        /// <summary>
        /// Whether this is the first time the app has run
        /// </summary>
        static public bool FirstRun = true;
        /// <summary>
        /// Height of the window
        /// </summary>
        static public int MainWindowHeight = 0;
        /// <summary>
        /// Height of the window
        /// </summary>
        static public int MainWindowWidth = 0;
        /// <summary>
        /// Default checkbox options
        /// </summary>
        static public ulong Options = 0xff80000000000000;
        /// <summary>
        /// File types to import
        /// </summary>
        static public List<string> ImportFileTypes = new List<string>(new string[] { "mp3", "wav", "m4a" });
        /// <summary>
        /// File types to sync
        /// </summary>
        static public List<string> SyncFileTypes = new List<string>(new string[] { "mp3", "ogg", "wav", "m4a", "wma", "flac", "ape", "apl" });
        /// <summary>
        /// Playlists to sync
        /// </summary>
        static public Dictionary<string, PlaylistItem> Playlists = new Dictionary<string, PlaylistItem>();
        /// <summary>
        /// Root playlist in iTunes to sync to
        /// </summary>
        static public string RootPlaylistFolder = "MC Playlists";
        /// <summary>
        /// Expressions
        /// </summary>
        static public Dictionary<string, string> Expressions = new Dictionary<string, string>();
        /// <summary>
        /// Default expressions
        /// </summary>
        static public Dictionary<string, string> DefaultExpressions = new Dictionary<string, string>();
        /// <summary>
        /// Number of milliseconds to wait before sending the sync ipod command (before library sync starts)
        /// </summary>
        static public int WaitTillSyncIPodBeforeSync = 5000;
        /// <summary>
        /// Number of milliseconds to wait before sending the refresh podcast feeds command
        /// </summary>
        static public int WaitTillRefreshPodcastFeeds = 5000;
        /// <summary>
        /// Number of milliseconds to wait after sending the refresh podcast feeds command
        /// </summary>
        static public int WaitAfterRefreshPodcastFeeds = 0;
        /// <summary>
        /// Number of times to send the refresh podcast feeds command
        /// </summary>
        static public int PodcastRefreshIterations = 1;
        /// <summary>
        /// Number of milliseconds to wait between refresh podcast feeds commands
        /// </summary>
        static public int WaitBetweenPodcastRefreshIterations = 30000;
        /// <summary>
        /// Number of milliseconds to wait when a sync track error occurs
        /// </summary>
        static public int WaitOnSyncTrackError = 1000;
        /// <summary>
        /// Number of retries on a sync track error before giving up
        /// </summary>
        static public int SyncTrackErrorRetries = 10;
        /// <summary>
        /// Number of seconds to wait when a library load error occurs
        /// </summary>
        static public int WaitOnLibraryLoadError = 1000;
        /// <summary>
        /// Number of retries loading a library before giving up
        /// </summary>
        static public int LibraryLoadRetries = 30;
        /// <summary>
        /// Current version
        /// </summary>
        static public int CurrentVersion;
        /// <summary>
        /// List of version changes
        /// </summary>
        static public Dictionary<int, List<VersionChange>> VersionChanges = new Dictionary<int, List<VersionChange>>();

        /// <summary>
        /// Manages the ImportFileTypes list as a string
        /// </summary>
        static public string ImportFileTypesRaw
        {
            get
            {
                string ret = string.Empty;
                for (int i = 0; i < ImportFileTypes.Count; i++)
                {
                    ret += ImportFileTypes[i];
                    if (i < (ImportFileTypes.Count - 1))
                        ret += ';';
                }

                return ret;
            }
            set
            {
                ImportFileTypes.Clear();
                if (value != null)
                {
                    foreach (string fileType in value.Split(';'))
                    {
                        if (fileType.Trim() != string.Empty)
                            ImportFileTypes.Add(fileType);
                    }
                }
            }
        }

        /// <summary>
        /// Manages the SyncFileTypes list as a string
        /// </summary>
        static public string SyncFileTypesRaw
        {
            get
            {
                string ret = string.Empty;
                for (int i = 0; i < SyncFileTypes.Count; i++)
                {
                    ret += SyncFileTypes[i];
                    if (i < (SyncFileTypes.Count - 1))
                        ret += ';';
                }

                return ret;
            }
            set
            {
                SyncFileTypes.Clear();
                if (value != null)
                {
                    foreach (string fileType in value.Split(';'))
                    {
                        if (fileType.Trim() != string.Empty)
                            SyncFileTypes.Add(fileType);
                    }
                }
            }
        }

        /// <summary>
        /// Static constructor loads the settings
        /// </summary>
        static Settings()
        {
            try
            {
                // Setup the default expressions
                DefaultExpressions.Add("Name", "[Name]");
                DefaultExpressions.Add("Artist", "[Artist]");
                DefaultExpressions.Add("Album", "[Album]");
                DefaultExpressions.Add("Album Artist", "[Album Artist]");
                DefaultExpressions.Add("Comment", "[Comment]");
                DefaultExpressions.Add("Composer", "[Composer]");
                DefaultExpressions.Add("Description", "[Description]");
                DefaultExpressions.Add("Disc #", "[Disc #]");
                DefaultExpressions.Add("Track #", "[Track #]");
                DefaultExpressions.Add("Genre", "[Genre]");
                DefaultExpressions.Add("Grouping", "[Grouping]");
                DefaultExpressions.Add("Lyrics", "[Lyrics]");
                DefaultExpressions.Add("Year", "[Year]");
                DefaultExpressions.Add("Mix Album", "[Mix Album]");
                DefaultExpressions.Add("BPM", "[BPM]");
                DefaultExpressions.Add("Sort Name", "Clean(Replace(Replace([Name],/(),/)),2)");
                DefaultExpressions.Add("Sort Artist", "Clean([Artist],2)");
                DefaultExpressions.Add("Sort Album Artist", "Clean([Album Artist (auto)],2)");
                DefaultExpressions.Add("Sort Album", "Clean([Album],2) - Clean([Album Artist (auto)],2)");
                DefaultExpressions.Add("Sort Composer", "Clean([Composer],2)");
                // Set expressions to defaults
                Expressions = new Dictionary<string, string>(DefaultExpressions);

                // If the file does not exist don't continue
                if (!File.Exists(Filename))
                    return;

                using (FileStream fs = new FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // If there's a config file, can't be a first run
                    FirstRun = false;

                    XmlDocument doc = new XmlDocument();
                    doc.Load(fs);

                    XmlElement root = doc["MCiS"];
                    // Get the version number of the settings file
                    int version = int.Parse(root.Attributes["Version"].InnerText.Replace(".", string.Empty).PadRight(4, '0'));

                    XmlAttribute attr;
                    attr = root.Attributes["MainWindowHeight"];
                    if (attr != null)
                        int.TryParse(attr.InnerText, out MainWindowHeight);
                    attr = root.Attributes["MainWindowWidth"];
                    if (attr != null)
                        int.TryParse(attr.InnerText, out MainWindowWidth);

                    XmlElement settings = root["Settings"];

                    Options = ulong.Parse(settings["Options"].InnerText, System.Globalization.NumberStyles.AllowHexSpecifier);
                    if (version < 5000)
                    {
                        if((Options & App.SHOWHIDDENPLAYLISTS) > 0)
                            Options ^= App.SHOWHIDDENPLAYLISTS;
                    }

                    XmlElement s;
                    s = settings["WaitTillSyncIPodBeforeSync"];
                    if (s != null)
                        WaitTillSyncIPodBeforeSync = int.Parse(s.InnerText);
                    s = settings["WaitTillRefreshPodcastFeeds"];
                    if (s != null)
                        WaitTillRefreshPodcastFeeds = int.Parse(s.InnerText);
                    s = settings["WaitAfterRefreshPodcastFeeds"];
                    if (s != null)
                        WaitAfterRefreshPodcastFeeds = int.Parse(s.InnerText);
                    s = settings["PodcastRefreshIterations"];
                    if (s != null)
                        PodcastRefreshIterations = int.Parse(s.InnerText);
                    s = settings["WaitBetweenPodcastRefreshIterations"];
                    if (s != null)
                        WaitBetweenPodcastRefreshIterations = int.Parse(s.InnerText);
                    s = settings["WaitOnSyncTrackError"];
                    if (s != null)
                        WaitOnSyncTrackError = int.Parse(s.InnerText);
                    s = settings["SyncTrackErrorRetries"];
                    if (s != null)
                        SyncTrackErrorRetries = int.Parse(s.InnerText);
                    s = settings["WaitOnLibraryLoadError"];
                    if (s != null)
                        WaitOnLibraryLoadError = int.Parse(s.InnerText);
                    s = settings["LibraryLoadRetries"];
                    if (s != null)
                        LibraryLoadRetries = int.Parse(s.InnerText);

                    SyncFileTypes.Clear();
                    foreach (XmlElement fileType in settings["SyncFileTypes"])
                    {
                        if (fileType.Name == "FileType")
                            SyncFileTypes.Add(fileType.InnerText);
                    }

                    ImportFileTypes.Clear();
                    foreach (XmlElement fileType in settings["ImportFileTypes"])
                    {
                        if (fileType.Name == "FileType")
                            ImportFileTypes.Add(fileType.InnerText);
                    }

                    Playlists.Clear();
                    foreach (XmlElement playlist in settings["Playlists"])
                    {
                        if (playlist.Name == "Playlist")
                        {
                            PlaylistItem pl = new PlaylistItem();
                            pl.Path = playlist.InnerText;

                            if (version >= 5000)
                            {
                                attr = playlist.Attributes["Hide"];
                                if (attr != null)
                                    pl.Hide = bool.Parse(attr.InnerText);
                                attr = playlist.Attributes["Rebuild"];
                                if (attr != null)
                                    pl.Rebuild = bool.Parse(attr.InnerText);
                                attr = playlist.Attributes["Shuffle"];
                                if (attr != null)
                                    pl.Shuffle = bool.Parse(attr.InnerText);
                                attr = playlist.Attributes["Selected"];
                                if (attr != null)
                                    pl.Selected = bool.Parse(attr.InnerText);
                                attr = playlist.Attributes["Ticked"];
                                if (attr != null)
                                    pl.Ticked = bool.Parse(attr.InnerText);
                                attr = playlist.Attributes["RemoveTracks"];
                                if (attr != null)
                                    pl.RemoveTracks = bool.Parse(attr.InnerText);
                            }
                            else//Older version
                            {
                                if (pl.Path == "\\\\")
                                    pl.Path = "\\*";
                                else
                                    pl.Path = pl.Path.Replace("\\\\", "");
                                pl.Selected = true;
                            }

                            Playlists.Add(pl.Path, pl);
                        }
                    }

                    RootPlaylistFolder = settings["RootPlaylistFolder"].InnerText;

                    XmlElement expressions = settings["Expressions"];
                    if (expressions != null)
                    {
                        foreach (XmlElement expression in expressions)
                        {
                            if (expression.Name == "Expression")
                            {
                                XmlAttribute field = expression.Attributes["Field"];
                                // Only add if this is an allowed field
                                if (DefaultExpressions.ContainsKey(field.InnerText))
                                {
                                    if (version < 5100)
                                    {
                                        switch (field.InnerText)
                                        {
                                            case "Sort Name":
                                                {
                                                    if (expression.InnerText != "[Name]")
                                                        Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                            case "Sort Artist":
                                                {
                                                    if (expression.InnerText != "[Artist]")
                                                        Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                            case "Sort Album Artist":
                                                {
                                                    if (expression.InnerText != "[Album Artist]")
                                                        Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                            case "Sort Album":
                                                {
                                                    if (expression.InnerText != "[Album]")
                                                        Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                            case "Sort Composer":
                                                {
                                                    if (expression.InnerText != "[Composer]")
                                                        Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                            default:
                                                {
                                                    Expressions[field.InnerText] = expression.InnerText;
                                                    break;
                                                }
                                        }
                                    }
                                    else
                                        Expressions[field.InnerText] = expression.InnerText;
                                }
                            }
                        }
                    }
                }
            }
            catch {/* Do nothing */}
        }

        /// <summary>
        /// Saves the configuration file
        /// </summary>
        /// <param name="h">Final height of window</param>
        /// <param name="w">Final width of window</param>
        static public void Save(int h, int w)
        {
            try
            {
                using (FileStream fs = new FileStream(Filename, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    XmlWriter writer = null;

                    try
                    {
                        // Setup the XML log file
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        writer = XmlWriter.Create(fs, settings);
                        writer.WriteStartDocument();

                        writer.WriteStartElement("MCiS");
                        writer.WriteStartAttribute("Version");
                        writer.WriteString(AboutBox.AssemblyVersion);
                        writer.WriteEndAttribute();

                        writer.WriteStartAttribute("MainWindowHeight");
                        writer.WriteString(h.ToString());
                        writer.WriteEndAttribute();

                        writer.WriteStartAttribute("MainWindowWidth");
                        writer.WriteString(w.ToString());
                        writer.WriteEndAttribute();

                        writer.WriteStartElement("Settings");

                        writer.WriteStartElement("Options");
                        writer.WriteString(Options.ToString("X"));
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitTillSyncIPodBeforeSync");
                        writer.WriteString(WaitTillSyncIPodBeforeSync.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitTillRefreshPodcastFeeds");
                        writer.WriteString(WaitTillRefreshPodcastFeeds.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitAfterRefreshPodcastFeeds");
                        writer.WriteString(WaitAfterRefreshPodcastFeeds.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("PodcastRefreshIterations");
                        writer.WriteString(PodcastRefreshIterations.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitBetweenPodcastRefreshIterations");
                        writer.WriteString(WaitBetweenPodcastRefreshIterations.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitOnSyncTrackError");
                        writer.WriteString(WaitOnSyncTrackError.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("SyncTrackErrorRetries");
                        writer.WriteString(SyncTrackErrorRetries.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("WaitOnLibraryLoadError");
                        writer.WriteString(WaitOnLibraryLoadError.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("LibraryLoadRetries");
                        writer.WriteString(LibraryLoadRetries.ToString());
                        writer.WriteEndElement();

                        writer.WriteStartElement("ImportFileTypes");
                        foreach (string fileType in ImportFileTypes)
                        {
                            writer.WriteStartElement("FileType");
                            writer.WriteString(fileType);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();

                        writer.WriteStartElement("SyncFileTypes");
                        foreach (string fileType in SyncFileTypes)
                        {
                            writer.WriteStartElement("FileType");
                            writer.WriteString(fileType);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();

                        writer.WriteStartElement("Playlists");
                        foreach (PlaylistItem playlist in Playlists.Values)
                        {
                            writer.WriteStartElement("Playlist");

                            writer.WriteStartAttribute("Hide");
                            writer.WriteString(playlist.Hide.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteStartAttribute("Rebuild");
                            writer.WriteString(playlist.Rebuild.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteStartAttribute("Shuffle");
                            writer.WriteString(playlist.Shuffle.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteStartAttribute("Selected");
                            writer.WriteString(playlist.Selected.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteStartAttribute("Ticked");
                            writer.WriteString(playlist.Ticked.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteStartAttribute("RemoveTracks");
                            writer.WriteString(playlist.RemoveTracks.ToString());
                            writer.WriteEndAttribute();

                            writer.WriteString(playlist.Path);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();

                        writer.WriteStartElement("RootPlaylistFolder");
                        writer.WriteString(RootPlaylistFolder);
                        writer.WriteEndElement();

                        writer.WriteStartElement("Expressions");
                        foreach (string fieldName in Expressions.Keys)
                        {
                            writer.WriteStartElement("Expression");

                            writer.WriteStartAttribute("Field");
                            writer.WriteString(fieldName);
                            writer.WriteEndAttribute();

                            writer.WriteString(Expressions[fieldName]);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();

                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    finally
                    {
                        if (writer != null)
                            writer.Close();
                    }
                }
            }
            catch {/* Do nothing */}
        }
    }
}
