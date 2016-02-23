using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Threading;
using iTunesLib;
using MediaCenter;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// The different status' of the synchronizer object
    /// </summary>
    public enum StatusEnum
    {
        Idle,
        Running,
        StopRequest,
        CloseRequest,
        ShutdownRequest,
        FinishingSuccessful,
        FinishingNotSuccessful
    }

    /// <summary>
    /// Handles all synchronizing operations
    /// </summary>
    public class Synchronizer
    {
        /// <summary>
        /// Holds the version of iTunes we are using
        /// </summary>
        public static Version iTunesVersion = new Version(0, 0, 0, 0);

        /// <summary>
        /// Holds the version of MC we are using
        /// </summary>
        public static Version MCVersion = new Version(0, 0, 0, 0);

        /// <summary>
        /// For posting messages to MC
        /// </summary>
        /// <param name="hwnd">Windows handle to MC</param>
        /// <param name="wMsg">Message to send</param>
        /// <param name="wParam">Message parameter</param>
        /// <param name="lParam">Message parameter</param>
        /// <returns>Whether successful</returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        /// <summary>
        /// Indicates the thread should continue to execute.
        /// Always access this through the modifier, not directly.
        /// The modifier is thread safe, direct access is not!
        /// </summary>
        private StatusEnum _running = StatusEnum.Idle;

        /// <summary>
        /// Whether to export only MC to iTunes - no aggregation
        /// </summary>
        protected bool _exportOnlyMCInfo;

        /// <summary>
        /// Whether to export only iTunes to MC - no aggregation
        /// </summary>
        protected bool _exportOnlyiTunesInfo;

        /// <summary>
        /// The thread safe modifier for _running.
        /// Indicates the status of the thread.
        /// </summary>
        public StatusEnum Running
        {
            get
            {
                lock (this)
                {
                    return _running;
                }
            }
            set
            {
                lock (this)
                {
                    _running = value;
                }
            }
        }

        /// <summary>
        /// The delegate for updating progress of synchronization
        /// </summary>
        /// <param name="e">Event data</param>
        public delegate void ProgressEventHandler(ProgressEventArgs e);

        /// <summary>
        /// The event handler for progress updates
        /// </summary>
        public event ProgressEventHandler UpdateProgress;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="exportOnlyMCInfo">Whether to export only MC to iTunes - no aggregation</param>
        /// <param name="exportOnlyiTunesInfo">Whether to export only iTunes to MC - no aggregation</param>
        public Synchronizer(bool exportOnlyMCInfo, bool exportOnlyiTunesInfo)
        {
            _exportOnlyMCInfo = exportOnlyMCInfo;
            _exportOnlyiTunesInfo = exportOnlyiTunesInfo;
        }

        /// <summary>
        /// Starts the synchronization process
        /// </summary>
        /// <param name="logFilename">The filename to log results to</param>
        /// <param name="flags">Synchronization settings</param>
        /// <param name="importFileTypes">The contents of the import file types text box</param>
        /// <param name="syncFileTypes">The contents of the sync file types text box</param>
        /// <param name="playlists">Playlists to sync</param>
        /// <param name="rootPlaylistFolder">The root folder for playlists</param>
        /// <param name="libraryIndex">The MC library index to use</param>
        /// <param name="libraryName">The MC library name to use</param>
        /// <param name="libraryPath">The path of the MC library to use</param>
        /// <param name="libraryInfo">Library information</param>
        /// <param name="expressions">MC Expressions</param>
        public void Start(string logFilename, ulong flags, string syncFileTypes, string importFileTypes, Dictionary<string, PlaylistItem> playlists, string rootPlaylistFolder, object[] libraryInfo, Dictionary<string, string> expressions)
        {
            // If it's already doing something, don't continue
            if (Running != StatusEnum.Idle)
                return;

            // Start our synchronization process
            Thread worker = new Thread(this.Go);
            worker.Start(new object[] { logFilename, flags, syncFileTypes, importFileTypes, playlists, rootPlaylistFolder, libraryInfo, expressions });

            // Set our status
            Running = StatusEnum.Running;
        }

        /// <summary>
        /// Do the work
        /// </summary>
        /// <param name="o">An array of arguments</param>
        private void Go(object o)
        {
            // Our settings
            object[] pars = (object[])o;
            string logFilename = (string)pars[0];
            ulong flags = (ulong)pars[1];
            string syncFileTypes = (string)pars[2];
            string importFileTypes = (string)pars[3];
            Dictionary<string, PlaylistItem> playlists = (Dictionary<string, PlaylistItem>)pars[4];
            string rootPlaylistFolder = (string)pars[5];
            object[] libraryInfo = (object[])pars[6];
            Dictionary<string, string> expressions = (Dictionary<string, string>)pars[7];

            // Get the view scheme or playlist name, if one has been specified
            string viewSchemeName = string.Empty, playListName = string.Empty;
            if (libraryInfo != null)
            {
                if (libraryInfo[1] != null)
                    viewSchemeName = (string)libraryInfo[1];
                else if (libraryInfo[2] != null)
                    playListName = (string)libraryInfo[2];
            }

            // Whether the sync is successful
            bool success = false;
            // Our resulting XML if required
            string results = null;

            try
            {
                try
                {
                    // Delete old log
                    if (File.Exists(logFilename))
                        File.Delete(logFilename);
                }
                catch {/* Do nothing */}

                // Setup the logger object
                if ((flags & App.NOLOG) > 0)
                    Logger.Create(null);
                else if ((flags & App.LOGTOCLIPBOARD) > 0)
                    Logger.Create();
                else
                    Logger.Create(logFilename);

                try
                {
                    // Will contain the sync file types if required
                    List<string> syncFileTypesList = null;
                    // Get the import file types into a list
                    if ((flags & App.SYNCCERTAINFILETYPES) > 0 && syncFileTypes.Trim() != string.Empty)
                    {
                        string[] arr = syncFileTypes.ToUpper().Split(';');
                        if (arr != null && arr.Length > 0)
                            syncFileTypesList = new List<string>(arr);
                    }

                    // Will contain the import file types if required
                    List<string> importFileTypesList = null;
                    // Get the import file types into a list
                    if ((flags & App.IMPORTFILESTOITUNES) > 0 && importFileTypes.Trim() != string.Empty)
                    {
                        string[] arr = importFileTypes.ToUpper().Split(';');
                        if (arr != null && arr.Length > 0)
                            importFileTypesList = new List<string>(arr);
                    }

                    // Write the log header
                    Logger.WriteHeader(flags, syncFileTypesList, importFileTypesList, playlists, rootPlaylistFolder, expressions, _exportOnlyMCInfo, _exportOnlyiTunesInfo);

                    // The results string to UI
                    string message = string.Empty;
                    // The results string to log
                    string logMessage = string.Empty;

                    // Our MC object
                    MCAutomation mc = null;
                    // Our iTunes object
                    iTunesApp iTunes = null;

                    // Our return values from the sync
                    int[] totals = null;

                    try
                    {
                        // Communicate with MC
                        UpdateProgress(new ProgressEventArgs("Looking for MC..."));
                        mc = GetJRMediaCenter(libraryInfo);

                        // Log MC info
                        Logger.WriteMCHeader(mc, viewSchemeName, playListName);

                        // Communicate with iTunes
                        UpdateProgress(new ProgressEventArgs("Looking for iTunes..."));
                        iTunes = GetiTunes();

                        // Log iTunes info
                        Logger.WriteiTunesHeader(iTunes);

                        // Hide iTunes if desired
                        if ((flags & App.MINIMISEITUNES) > 0)
                        {
                            try
                            {
                                foreach (IITWindow itunesWindow in iTunes.Windows)
                                {
                                    itunesWindow.Minimized = true;
                                }
                            }
                            catch {/* Do nothing */}
                        }

                        // Send sync iPod request to iTunes
                        if ((flags & App.SYNCHRONIZEIPODBEFORE) > 0)
                        {
                            DateTime waitTill = DateTime.Now.AddMilliseconds(Settings.WaitTillSyncIPodBeforeSync);
                            while (DateTime.Now < waitTill)
                            {
                                UpdateProgress(new ProgressEventArgs("Sending Synchronize iPod command in {0} seconds...", Math.Ceiling((waitTill - DateTime.Now).TotalSeconds)));
                                // Check the current status - whether we need to abort or not
                                CheckCurrentStatus();
                                Thread.Sleep(100);
                            }
                            iTunes.UpdateIPod();
                        }

                        // Send a refresh Podcast feeds message to iTunes before we start
                        // Hopefully by the time the sync finishes the refresh will have been done
                        if ((flags & App.REFRESHPODCASTS) > 0)
                        {
                            DateTime waitTill;

                            for (int i = 0; i < Settings.PodcastRefreshIterations; i++)
                            {
                                waitTill = DateTime.Now.AddMilliseconds((i == 0) ? Settings.WaitTillRefreshPodcastFeeds:Settings.WaitBetweenPodcastRefreshIterations);
                                while (DateTime.Now < waitTill)
                                {
                                    UpdateProgress(new ProgressEventArgs("Sending Refresh Podcast Feeds command in {0} seconds...", Math.Ceiling((waitTill - DateTime.Now).TotalSeconds)));
                                    // Check the current status - whether we need to abort or not
                                    CheckCurrentStatus();
                                    Thread.Sleep(100);
                                }

                                // Do the refresh
                                iTunes.UpdatePodcastFeeds();
                            }

                            waitTill = DateTime.Now.AddMilliseconds(Settings.WaitAfterRefreshPodcastFeeds);
                            while (DateTime.Now < waitTill)
                            {
                                UpdateProgress(new ProgressEventArgs("Synchronizing in {0} seconds...", Math.Ceiling((waitTill - DateTime.Now).TotalSeconds)));
                                // Check the current status - whether we need to abort or not
                                CheckCurrentStatus();
                                Thread.Sleep(100);
                            }
                        }

                        // If required, refresh smartlists in MC
                        if (playListName != null && playListName != string.Empty)
                            PostMessage(new IntPtr(mc.GetWindowHandle()), MC.WM_MC_COMMAND, MC.MCC_REFRESH, 0);

                        // Synchronize!
                        totals = Synchronize(mc, iTunes, flags, syncFileTypesList, importFileTypesList, playlists, rootPlaylistFolder, viewSchemeName, playListName, expressions);

                        // The result
                        message = "Successful";
                        success = true;
                    }
                    catch (UserAbortException ex)
                    {
                        // The result will be the exception
                        message = ex.Message;
                    }
                    catch (Exception ex)
                    {
                        // The result will be the exception
                        message = "Please review the log for further detail.";
                        logMessage = ex.Message;
                    }
                    finally
                    {
                        try
                        {
                            UpdateProgress(new ProgressEventArgs("Cleaning up..."));

                            // Send sync iPod request to iTunes
                            if (success && (flags & App.SYNCHRONIZEIPOD) > 0)
                                iTunes.UpdateIPod();

                            // Tell iTunes to Quit
                            if (success && (flags & App.QUITITUNES) > 0)
                                iTunes.Quit();

                            if (totals != null)
                                UpdateProgress(new ProgressEventArgs("Successfully analyzed {0} tracks, of which {1} were synchronized.", totals[0], totals[1]));
                            else
                                UpdateProgress(new ProgressEventArgs("Synchronize unsuccessful. {0}", message));

                            // Finalize the log file
                            Logger.WriteFooter(totals, (logMessage == string.Empty) ? message:logMessage);
                        }
                        finally
                        {
                            // We're finished with MC and iTunes, clear the COM objects
                            if (mc != null)
                            {
                                Marshal.ReleaseComObject(mc);
                                mc = null;
                            }
                            if (iTunes != null)
                            {
                                Marshal.ReleaseComObject(iTunes);
                                iTunes = null;
                            }

                            // Garbage collection - shouldn't really be necessary
                            // but the COM objects act up if we don't clean them up
                            // quickly.
                            GC.Collect();
                        }
                    }
                }
                finally
                {
                    // Make sure we close the log file properly no matter what
                    if((flags & App.LOGTOCLIPBOARD) > 0)
                        results = Logger.Destroy(true);
                    else
                        Logger.Destroy(false);
                }
            }
            catch (Exception ex)
            {
                // Something went wrong processing the log file
                UpdateProgress(new ProgressEventArgs(ex.Message));
            }

            StatusEnum finalStatus = Running;
            if (finalStatus == StatusEnum.Running)
            {
                finalStatus = (success) ? StatusEnum.FinishingSuccessful : StatusEnum.FinishingNotSuccessful;
                Running = finalStatus;
            }

            // We are no longer running
            Running = StatusEnum.Idle;

            // Finished! (or about to at least)
            UpdateProgress(new ProgressEventArgs(true, finalStatus, results));
        }

        /// <summary>
        /// Synchronizes iTunes and MC libraries
        /// </summary>
        /// <param name="mc">Reference to MC object</param>
        /// <param name="iTunes">Reference to iTunes object</param>
        /// <param name="flags">Synchronization settings</param>
        /// <param name="syncFileTypesList">The contents of the sync file types text box</param>
        /// <param name="importFileTypesList">The contents of the import file types text box</param>
        /// <param name="playlists">Any playlists to sync</param>
        /// <param name="rootPlaylistFolder">The root folder for playlists</param>
        /// <param name="viewSchemeName">The MC view scheme to use</param>
        /// <param name="playListName">The MC playlist to use</param>
        /// <param name="expressions">MC Expressions</param>
        /// <returns>Array of integers, first element should contain the total analysed, the second element should contain the total sync'd</returns>
        private int[] Synchronize(MCAutomation mc, iTunesApp iTunes, ulong flags, List<string> syncFileTypesList, List<string> importFileTypesList, Dictionary<string, PlaylistItem> playlists, string rootPlaylistFolder, string viewSchemeName, string playListName, Dictionary<string, string> expressions)
        {
            // Create our fields
            CreateCustomFields(mc, flags);

            // Contains MC Library
            Dictionary<string, MCTrack> mcCache = null;
            // Contains MC Albums
            Dictionary<string, List<MCTrack>> mcAlbumCache = null;

            // Cache the MC library - we have to do this so we can hash on the pathname
            GetMCCache(mc, ref mcCache, ref mcAlbumCache, syncFileTypesList, flags, viewSchemeName, playListName, expressions);
            if (mcCache.Count == 0)
                throw new Exception("Could not find any tracks in MC.");

            // Get all tracks in the iTunes libary
            IITTrackCollection itunesTracks = iTunes.LibraryPlaylist.Tracks;
            if (itunesTracks == null)
                throw new Exception("iTunes didn't return a valid library, try importing a file into iTunes.");
            // Progress variables
            DateTime before = DateTime.MinValue;
            int i = 1;
            int total = itunesTracks.Count;
            // Sync counter
            int c = 0;

            // Holds MC tracks that may have had their album rating re-calculated in the first pass
            List<MCTrack> secondPassTracks = new List<MCTrack>();
            // Holds broken links
            List<IITTrack> brokenLinks = new List<IITTrack>();

            // Will hold imports count
            int importCount = 0;
            // Will hold broken links count
            int brokenCount = 0;

            // Now go through all the tracks in the iTunes library
            do
            {
                UpdateProgress(new ProgressEventArgs("Analyzing..."));
                if(total > 0)
                    UpdateProgress(new ProgressEventArgs(1, total, i));

                for (; i <= total; i++)
                {
                    IITTrack t = itunesTracks[i];

                    // Update the progress text and bar
                    if (DateTime.Now > before.AddSeconds(0.5) || i == total)
                    {
                        UpdateProgress(new ProgressEventArgs(i, "Analyzing {0} of {1}, synchronized {2} so far", i, total, c));
                        before = DateTime.Now;
                    }

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    // If the object is not a physical track, skip it
                    if (!(t is IITFileOrCDTrack))
                        continue;

                    iTunesTrack itunesTrack = null;

                    try { itunesTrack = new iTunesTrack((IITFileOrCDTrack)t, flags); }
                    catch (Exception ex)
                    {
                        Logger.WriteiTunesCacheError(t.Artist, t.Album, t.Name, ex.Message);
                        continue;
                    }

                    // If the object has no location, must be a broken link, skip it
                    if (itunesTrack.Fullpathname == string.Empty)
                    {
                        // Add to the list to be removed later
                        if ((flags & App.REMOVEBROKENLINKS) > 0)
                            brokenLinks.Add(t);
                        continue;
                    }

                    // Get the track from the MC database
                    MCTrack mcTrack = null;

                    // Get from the MC cache
                    try { mcTrack = mcCache[itunesTrack.Key]; }
                    catch { /* Do nothing */ }

                    // Sets up logging of the track changes
                    TrackMonitor trackSync = new TrackMonitor(itunesTrack, mcTrack);

                    try
                    {
                        // If the add artwork to podcasts option is on then add artwork to the podcast (if there is none added already)
                        if ((flags & App.ADDARTWORKTOPODCASTS) > 0 && itunesTrack.AudioPodcast && !itunesTrack.HasArtwork)
                        {
                            string folderjpg = itunesTrack.Pathname + "Folder.jpg";
                            if (File.Exists(folderjpg))
                            {
                                trackSync.SetArtwork(ChangeType.iTunes, null, folderjpg);
                                if ((flags & App.TESTMODE) == 0)
                                    itunesTrack.Artwork = folderjpg;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        trackSync.SetError("Artwork", ex.Message);
                    }

                    // Was a valid track retrieved? If not skip it.
                    if (mcTrack == null || mcTrack.Fullpathname == string.Empty)
                    {
                        Logger.WriteFileNotFound(itunesTrack.Fullpathname);
                        if (trackSync.Changed)
                            Logger.WriteTrackMonitor(trackSync);
                        continue;
                    }

                    // We found the track in iTunes, set the track index
                    mcTrack.iTunesIndex = i;

                    // Get the persistent ID of the track
                    // We only need this for playlists later
                    if((flags & App.PLAYLISTS) > 0)
                    {
                        int h, l;
                        object o = (object)t;
                        iTunes.GetITObjectPersistentIDs(ref o, out h, out l);

                        mcTrack.PersistentIDHigh = h;
                        mcTrack.PersistentIDLow = l;
                        mcTrack.PersistentID = GetPersistentID(h, l);
                        mcTrack.Ticked = t.Enabled;
                    }

                    // Syncs the track
                    int errors = 0;
                    while (this.SynchronizeTrack(trackSync, flags, mcAlbumCache, ++errors == Settings.SyncTrackErrorRetries))
                    {
                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();
                        if (errors == Settings.SyncTrackErrorRetries)
                            break;
                        Thread.Sleep(Settings.WaitOnSyncTrackError);
                    }

                    // Has anything changed?
                    if (trackSync.Changed)
                    {
                        // There have been changes related to this track
                        mcTrack.Synchronised = true;

                        // Write to the log
                        Logger.WriteTrackMonitor(trackSync);

                        // Increment the changed counter
                        c++;

                        // If the rating has been changed and Album Rating is checked, we have to go over this one again once we're done
                        if ((flags & App.ALBUMRATING) > 0 && itunesTrack.RatingChanged && itunesTrack.AlbumRatingCalculated)
                        {
                            List<MCTrack> album = null;

                            try { album = mcAlbumCache[mcTrack.AlbumKey]; }
                            catch {/* Do nothing */}

                            if (album != null)
                            {
                                foreach (MCTrack track in album)
                                {
                                    secondPassTracks.Add(track);
                                }
                            }
                        }
                    }
                }

                // Are we importing and does the file extension match?
                if ((flags & App.IMPORTFILESTOITUNES) > 0 && importFileTypesList != null)
                {
                    // Will hold the tracks to import
                    Dictionary<string, MCTrack> importTracks = new Dictionary<string, MCTrack>(StringComparer.OrdinalIgnoreCase);

                    // We only want to do this once, so remove the import flag
                    flags ^= App.IMPORTFILESTOITUNES;

                    int j = 0;
                    int k = mcCache.Values.Count;

                    UpdateProgress(new ProgressEventArgs("Searching for files requiring import..."));
                    UpdateProgress(new ProgressEventArgs(1, k, 1));

                    foreach (MCTrack track in mcCache.Values)
                    {
                        j++;

                        // Update the progress text and bar
                        if (DateTime.Now > before.AddSeconds(0.5) || j == k)
                        {
                            UpdateProgress(new ProgressEventArgs(j, "Searching for files requiring import... {0} of {1}", j, k));
                            before = DateTime.Now;
                        }

                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();

                        // This will contain the filename we will be importing
                        string filename;
                        // Is there a corresponding iTunes track for this MC track? If so, skip it
                        if (track.iTunesIndex > 0)
                            continue;
                        // Are we using alternative keys and is there an alternative key for this MC track?
                        else if ((flags & App.USEALTERNATIVEFILEKEY) > 0 && track.iTunesFileKey != string.Empty)
                        {
                            // If the file type is not one we want to import, skip
                            string fileType = track.iTunesFileKey.Substring(track.iTunesFileKey.LastIndexOf('.') + 1).ToUpper();
                            if (importFileTypesList.Contains(fileType))
                                filename = track.iTunesFileKey;
                            else
                                continue;
                        }
                        // If the file type is not one we want to import, skip
                        else if (importFileTypesList.Contains(track.FileType))
                            filename = track.Fullpathname;
                        else
                            continue;

                        // We got this far, the track needs importing
                        importTracks.Add(filename, track);
                    }

                    // Import the tracks into iTunes
                    k = importTracks.Count;

                    if (k > 0)
                    {
                        UpdateProgress(new ProgressEventArgs("Importing files into iTunes..."));
                        UpdateProgress(new ProgressEventArgs(1, k, 1));

                        j = 0;

                        List<string> filenames = new List<string>(importTracks.Keys);
                        object o = (object)filenames.ToArray();
                        IITOperationStatus op = null;
                        string errormsg = "Unsuccessful";

                        if ((flags & App.TESTMODE) == 0)
                        {
                            try { op = iTunes.LibraryPlaylist.AddFiles(ref o); }
                            catch (Exception ex) { errormsg = ex.Message; }
                        }

                        if (op != null)
                        {
                            importCount = op.Tracks.Count;

                            foreach (IITTrack t in op.Tracks)
                            {
                                j++;

                                // Update the progress text and bar
                                if (DateTime.Now > before.AddSeconds(0.5) || j == k)
                                {
                                    UpdateProgress(new ProgressEventArgs(j, "Importing files into iTunes... {0} of {1}", j, k));
                                    before = DateTime.Now;
                                }

                                // Check the current status - whether we need to abort or not
                                CheckCurrentStatus();

                                if (!(t is IITFileOrCDTrack))
                                    continue;

                                IITFileOrCDTrack file = (IITFileOrCDTrack)t;

                                string filename = file.Location;

                                MCTrack track = null;

                                try { track = importTracks[filename]; }
                                catch { /* Do nothing */ }

                                if (track != null)
                                {
                                    track.iTunesRatingSync = -1;
                                    track.iTunesPlayedCountSync = 0;
                                    track.iTunesSkippedCountSync = 0;
                                    importTracks.Remove(filename);
                                    Logger.WriteImportFile(filename, true, "Successful");
                                }
                            }
                        }

                        foreach (string filename in importTracks.Keys)
                        {
                            Logger.WriteImportFile(filename, false, errormsg);
                        }

                        UpdateProgress(new ProgressEventArgs(k));
                    }
                }
            } while (i <= (total = itunesTracks.Count));

            // Go over the second pass tracks for Album Ratings that may have changed
            if ((flags & App.ALBUMRATING) > 0 && secondPassTracks.Count > 0)
            {
                int arcount = secondPassTracks.Count;
                int arindex = 0;

                UpdateProgress(new ProgressEventArgs("Re-checking Album Ratings..."));
                UpdateProgress(new ProgressEventArgs(1, arcount, 1));

                foreach (MCTrack secondPassTrack in secondPassTracks)
                {
                    bool changed = secondPassTrack.Synchronised;

                    arindex++;

                    // Update the progress text and bar
                    if (DateTime.Now > before.AddSeconds(0.5) || arindex == arcount)
                    {
                        UpdateProgress(new ProgressEventArgs(arindex, "Re-checking Album Ratings... {0} of {1}", arindex, arcount));
                        before = DateTime.Now;
                    }

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    if (secondPassTrack.iTunesIndex <= 0)
                        continue;

                    iTunesTrack t = new iTunesTrack((IITFileOrCDTrack)itunesTracks[secondPassTrack.iTunesIndex], App.ALBUMRATING);

                    TrackMonitor secondPassTrackSync = new TrackMonitor(t, secondPassTrack);

                    // Syncs the track
                    int errors = 0;
                    while (this.SynchronizeTrack(secondPassTrackSync, App.ALBUMRATING, null, ++errors == Settings.SyncTrackErrorRetries))
                    {
                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();
                        if (errors == Settings.SyncTrackErrorRetries)
                            break;
                        Thread.Sleep(Settings.WaitOnSyncTrackError);
                    }

                    if (secondPassTrackSync.Changed)
                    {
                        Logger.WriteTrackMonitor(secondPassTrackSync);
                        if(!changed)
                            c++;
                    }
                }
            }

            if ((flags & App.REMOVEBROKENLINKS) > 0 && brokenLinks.Count > 0)
            {
                int blcount = brokenLinks.Count;
                int blindex = 0;

                UpdateProgress(new ProgressEventArgs("Removing broken links in iTunes..."));
                UpdateProgress(new ProgressEventArgs(1, blcount, 1));

                foreach (IITTrack t in brokenLinks)
                {
                    blindex++;

                    // Update the progress text and bar
                    if (DateTime.Now > before.AddSeconds(0.5) || blindex == blcount)
                    {
                        UpdateProgress(new ProgressEventArgs(blindex, "Removing broken links in iTunes... {0} of {1}", blindex, blcount));
                        before = DateTime.Now;
                    }

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    // Store these values before we delete so we can write them to the log
                    string artist = t.Artist, album = t.Album, name = t.Name;
                    try
                    {
                        if ((flags & App.TESTMODE) == 0)
                            t.Delete();
                        brokenCount++;
                        Logger.WriteRemoveBrokenLink(artist, album, name, "Successful");
                    }
                    catch (Exception ex)
                    {
                        Logger.WriteRemoveBrokenLink(artist, album, name, ex.Message);
                    }
                }
            }

            // Sync the On-The-Go playlists across to MC
            if ((flags & App.ONTHEGOPLAYLISTS) > 0)
            {
                // Update progress text
                UpdateProgress(new ProgressEventArgs("Searching for On-The-Go playlists in iTunes...."));

                // Get the MC playlist object
                IMJPlaylistsAutomation mcps = mc.GetPlaylists();
                // Delete the On-The-Go playlists before we start - that way we're starting from scratch
                mcps.DeletePlaylist("", App.ONTHEGOPATH);

                foreach (IITPlaylist itunesPlaylist in iTunes.LibrarySource.Playlists)
                {
                    if (!(itunesPlaylist is IITUserPlaylist))
                        continue;

                    IITUserPlaylist onTheGo = (IITUserPlaylist)itunesPlaylist;
                    // Skip if this is not an On-The-Go playlist, otherwise create it in MC
                    if (!onTheGo.Name.StartsWith("On-The-Go") || onTheGo.get_Parent() != null)
                        continue;

                    // Must be an On-The-Go playlist...
                    // Get the name of the playlist
                    string name = onTheGo.Name;

                    // Update progress text and bar
                    UpdateProgress(new ProgressEventArgs("Creating On-The-Go playlist: {0}\\{1}", App.ONTHEGOPATH, name));

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    // Create the playlist
                    IMJPlaylistAutomation p = mcps.CreatePlaylist(App.ONTHEGOPATH, name);
                    if (p == null || p.Name != name)
                        continue;

                    // Log
                    Logger.WritePlaylistCreated(string.Format("{0}\\{1}", App.ONTHEGOPATH, name));

                    int k = onTheGo.Tracks.Count;
                    if (k <= 0)
                        continue;

                    // Initialise progress bar and text
                    UpdateProgress(new ProgressEventArgs(1, k, 1));

                    for (int j = 1; j <= k; j++)
                    {
                        if (DateTime.Now > before.AddSeconds(0.5) || j == k)
                        {
                            UpdateProgress(new ProgressEventArgs(j, "{0}\\{1}: adding track {2} of {3}", App.ONTHEGOPATH, name, j, k));
                            before = DateTime.Now;
                        }

                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();

                        IITTrack t = onTheGo.Tracks[j];
                        if (t is IITFileOrCDTrack)
                        {
                            IITFileOrCDTrack file = (IITFileOrCDTrack)t;
                            if (file.Location != null)
                            {
                                MCTrack mcTrack = null;
                                try { mcTrack = mcCache[file.Location.ToUpper()]; }
                                catch {/* Do nothing */}
                                if (mcTrack != null)
                                {
                                    p.AddFile(mcTrack.Fullpathname, -1);
                                    Logger.WritePlaylistAddFile(string.Format("{0}\\{1}", App.ONTHEGOPATH, name), mcTrack.Fullpathname);
                                }
                            }
                        }
                    }
                }
            }

            if ((flags & (App.PLAYLISTS | App.TESTMODE)) == App.PLAYLISTS)
            {
                UpdateProgress(new ProgressEventArgs("Compiling iTunes playlists..."));

                // Refresh MC smartlists
                PostMessage(new IntPtr(mc.GetWindowHandle()), MC.WM_MC_COMMAND, MC.MCC_REFRESH, 0);

                // Will hold all itunes playlists that relate to MC
                Dictionary<string, IITUserPlaylist> itunesPlaylists = new Dictionary<string, IITUserPlaylist>();

                IITPlaylistCollection alliTunesPlaylists = iTunes.LibrarySource.Playlists;

                int plCount = alliTunesPlaylists.Count;
                int plIndex = 0;

                UpdateProgress(new ProgressEventArgs(1, plCount, 1));

                Dictionary<string, IITUserPlaylist> iTunesPlaylistsHash = new Dictionary<string, IITUserPlaylist>(alliTunesPlaylists.Count);
                foreach (IITPlaylist iTunesPlaylist in alliTunesPlaylists)
                {
                    if (!(iTunesPlaylist is IITUserPlaylist))
                        continue;

                    UpdateProgress(new ProgressEventArgs(++plIndex));

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    IITUserPlaylist iTunesUserPlaylist = (IITUserPlaylist)iTunesPlaylist;

                    if (iTunesUserPlaylist.Kind != ITPlaylistKind.ITPlaylistKindUser 
                        || (iTunesUserPlaylist.SpecialKind != ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone && iTunesUserPlaylist.SpecialKind != ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindFolder)
                        || (iTunesUserPlaylist.SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone && iTunesUserPlaylist.Smart))
                        continue;

                    string path = string.Empty;
                    string name = iTunesPlaylist.Name;

                    IITUserPlaylist parent = iTunesUserPlaylist.get_Parent();
                    while(parent != null)
                    {
                        path = parent.Name + '\v' + path;
                        parent = parent.get_Parent();
                    }

                    string fullpath = (path == string.Empty) ? name : path + '\v' + name;

                    iTunesPlaylistsHash[path + iTunesPlaylist.Name] = iTunesUserPlaylist;
                }

                plCount = playlists.Count;
                plIndex = 0;

                UpdateProgress(new ProgressEventArgs("Searching for selected playlists..."));
                UpdateProgress(new ProgressEventArgs(1, plCount, 1));

                // Now the playlists
                foreach (string playlistPath in playlists.Keys)
                {
                    plIndex++;

                    PlaylistItem pl = playlists[playlistPath];
                    // Is the playlist selected
                    if (!pl.Selected)
                        continue;

                    UpdateProgress(new ProgressEventArgs(plIndex, "Searching for playlist {0}", playlistPath));

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();

                    string iTunesPlaylistPath = (((flags & App.USEPLAYLISTROOT) > 0) ? string.Empty : rootPlaylistFolder + '\v') + playlistPath.Replace('\\', '\v');
                    if((flags & App.REPLACEFORWARDSLASHES) > 0)
                        iTunesPlaylistPath = iTunesPlaylistPath.Replace('/', '\\');
                    IITUserPlaylist p = GetPlaylist(iTunes, iTunesPlaylistsHash, iTunesPlaylistPath, pl.Folder);
                    if (p != null)
                        itunesPlaylists.Add(playlistPath, p);
                }

                plCount = iTunes.LibrarySource.Playlists.Count;
                plIndex = 0;

                UpdateProgress(new ProgressEventArgs(1, plCount, 1));

                // Delete any playlists that shouldn't be there
                List<IITUserPlaylist> playlistsToDelete = new List<IITUserPlaylist>();
                List<string> playlistsToDeleteNames = new List<string>();
                foreach (string ituneshash in iTunesPlaylistsHash.Keys)
                {
                    IITUserPlaylist itunesPlaylist = iTunesPlaylistsHash[ituneshash];

                    // Check the current status - whether we need to abort or not
                    CheckCurrentStatus();
                    // Update progress
                    UpdateProgress(new ProgressEventArgs(++plIndex, "Checking playlist {0}", itunesPlaylist.Name));

                    // If the playlist is one that is selected then we don't want to delete it
                    string mchash;
                    if ((flags & App.USEPLAYLISTROOT) > 0)
                    {
                        if ((flags & App.REPLACEFORWARDSLASHES) > 0)
                            mchash = ituneshash.Replace('\\', '/');
                        else
                            mchash = ituneshash;

                        mchash = mchash.Replace('\v', '\\');
                    }
                    else
                    {
                        if (ituneshash.StartsWith(rootPlaylistFolder + '\v'))
                        {
                            mchash = ituneshash.Substring(rootPlaylistFolder.Length + 1);
                            if ((flags & App.REPLACEFORWARDSLASHES) > 0)
                                mchash = mchash.Replace('\\', '/');

                            mchash = mchash.Replace('\v', '\\');
                        }
                        else
                            continue;
                    }

                    PlaylistItem pitem = null;
                    if (!((flags & App.REPLACEFORWARDSLASHES) > 0 && ituneshash.Contains("/")))
                    {
                        try
                        {
                            pitem = playlists[mchash];
                        }
                        catch {/*Do nothing*/}
                    }

                    bool parentSelected = false;
                    string path = string.Empty;
                    int n = mchash.LastIndexOf('\\');
                    if (n >= 0)
                        path = mchash.Substring(0, mchash.Length - (mchash.Length - n));
                    else
                        path = string.Empty;

                    PlaylistItem parent = null;
                    // If the parent is selected, delete it
                    if (playlists.TryGetValue(path, out parent))
                    {
                        if (parent.Selected)
                            parentSelected = true;
                    }

                    if (pitem == null)
                    {
                        // If playlist is under the specified root folder, delete it
                        if ((flags & App.USEPLAYLISTROOT) == 0 || parentSelected)
                        {
                            playlistsToDelete.Add(itunesPlaylist);
                            playlistsToDeleteNames.Add(mchash);
                            continue;
                        }
                    }
                    else
                    {
                        // If it's selected, just go to the next playlist
                        if (pitem.Selected)
                            continue;

                        // Rules are different depending on whether syncing to root
                        if ((flags & App.USEPLAYLISTROOT) == 0 || parentSelected)
                        {
                            // If a folder, check if any children are selected. If not, delete folder.
                            if (pitem.Folder)
                            {
                                bool hasSelectedChildren = false;
                                foreach (PlaylistItem p in playlists.Values)
                                {
                                    if (p.Path != pitem.Path && p.Path.StartsWith(pitem.Path))
                                    {
                                        if (p.Selected)
                                        {
                                            hasSelectedChildren = true;
                                            break;
                                        }
                                    }
                                }
                                if (!hasSelectedChildren)
                                {
                                    playlistsToDelete.Add(itunesPlaylist);
                                    playlistsToDeleteNames.Add(mchash);
                                }
                            }
                            else
                            {
                                playlistsToDelete.Add(itunesPlaylist);
                                playlistsToDeleteNames.Add(mchash);
                            }
                        }
                    }
                }
                // Delete the playlists that shouldn't be there
                for (int j = 0; j < playlistsToDelete.Count; j++)
                {
                    IITUserPlaylist p = playlistsToDelete[j];
                    try
                    {
                        p.Delete();
                        Logger.WritePlaylistDeleted(playlistsToDeleteNames[j]);
                    }
                    catch {/* Do nothing */}
                }

                // Get MC playlists
                Dictionary<string, IMJPlaylistAutomation> mcPlaylists = GetMCPlaylists(mc, false);

                // Compile the list of playlists to sync, do the unticked ones first
                List<PlaylistItem> playlistList = new List<PlaylistItem>(mcPlaylists.Count);
                foreach (string path in mcPlaylists.Keys)
                {
                    PlaylistItem pl;
                    try { pl = playlists[path]; }
                    catch { continue; }

                    // If the playlist isn't one that is selected, skip it
                    if (!pl.Selected)
                        continue;

                    if (!pl.Ticked)
                        playlistList.Insert(0, pl);
                    else
                        playlistList.Add(pl);
                }

                // Go through each playlist
                foreach (PlaylistItem pl in playlistList)
                {
                    IMJPlaylistAutomation mcPlaylist = mcPlaylists[pl.Path];
                    IITUserPlaylist itunesPlaylist = itunesPlaylists[pl.Path];

                    // Shuffle the playlist in MC if required
                    if (!pl.Smart && pl.Shuffle && pl.Rebuild)
                    {
                        IMJFilesAutomation files = mcPlaylist.GetFiles();
                        files.Shuffle();
                        mcPlaylist.SetFiles(files);
                    }

                    // Delete all tracks in the playlist if this is to be rebuilt exact
                    if (pl.Rebuild)
                    {
                        int j = 0, k = itunesPlaylist.Tracks.Count;

                        // If the playlist is not empty
                        if (k > 0)
                        {
                            // Initialise progress bar
                            UpdateProgress(new ProgressEventArgs(1, k, 1));

                            while (itunesPlaylist.Tracks.Count > 0)
                            {
                                j++;
                                if (DateTime.Now > before.AddSeconds(0.5) || j == k)
                                {
                                    UpdateProgress(new ProgressEventArgs(j, "{0}: cleaning up track {1} of {2}", pl.Path, j, k));
                                    before = DateTime.Now;
                                }

                                // Check the current status - whether we need to abort or not
                                CheckCurrentStatus();

                                // Delete the track from the playlist
                                IITTrack t = itunesPlaylist.Tracks[1];
                                if (t is IITFileOrCDTrack)
                                    Logger.WritePlaylistDeleteFile(pl.Path, ((IITFileOrCDTrack)t).Location);
                                t.Delete();
                            }
                        }
                    }

                    // Store a check list of tracks so we can remove any unwanted tracks later
                    List<long> checkList = new List<long>();
                    IITTrackCollection playlistTracks = itunesPlaylist.Tracks;
                    int itunesFileCount = playlistTracks.Count;

                    if (itunesFileCount > 0)
                    {
                        UpdateProgress(new ProgressEventArgs(1, itunesFileCount, 1));

                        for (int itunesFileIndex = 1; itunesFileIndex <= itunesFileCount; itunesFileIndex++)
                        {
                            IITTrack t = playlistTracks[itunesFileIndex];

                            // Update the progress text and bar
                            if (DateTime.Now > before.AddSeconds(0.5) || itunesFileIndex == itunesFileCount)
                            {
                                UpdateProgress(new ProgressEventArgs(itunesFileIndex, "{0}: compiling track {1} of {2}", pl.Path, itunesFileIndex, itunesFileCount));
                                before = DateTime.Now;
                            }

                            // Check the current status - whether we need to abort or not
                            CheckCurrentStatus();

                            // Get the persistent ID of the track
                            int h, l;
                            object o = (object)t;
                            iTunes.GetITObjectPersistentIDs(ref o, out h, out l);
                            // Add to the checklist
                            checkList.Add(GetPersistentID(h, l));
                        }
                    }

                    IMJFilesAutomation tracks = mcPlaylist.GetFiles();
                    int mcFileCount = tracks.GetNumberFiles();

                    if (mcFileCount > 0)
                    {
                        UpdateProgress(new ProgressEventArgs(1, mcFileCount, 1));

                        for (int mcFileIndex = 0; mcFileIndex < mcFileCount; mcFileIndex++)
                        {
                            // Update the progress text and bar
                            int j = mcFileIndex + 1;
                            if (DateTime.Now > before.AddSeconds(0.5) || j == mcFileCount)
                            {
                                UpdateProgress(new ProgressEventArgs(j, "{0}: adding track {1} of {2}", pl.Path, j, mcFileCount));
                                before = DateTime.Now;
                            }

                            // Check the current status - whether we need to abort or not
                            CheckCurrentStatus();

                            // Get the key for the MCTrack object in the cache
                            string key, filename;
                            try
                            {
                                if ((flags & App.USEALTERNATIVEFILEKEY) > 0)
                                {
                                    // If we're using alternative file keys we need to get the iTunesFileKey and filename
                                    IMJFileAutomation file = tracks.GetFile(mcFileIndex);
                                    filename = file.Get("iTunesFileKey", false);
                                    filename = (filename == null) ? string.Empty:filename.Trim();
                                    filename = (filename == string.Empty) ? file.Filename:filename;
                                }
                                else
                                {
                                    IMJFileAutomation file = tracks.GetFile(mcFileIndex);
                                    filename = file.Filename;
                                }

                                key = filename.ToUpper();
                            }
                            catch { continue; }

                            MCTrack track = null;

                            try { track = mcCache[key]; }
                            catch {/* Do nothing */}

                            if (track != null && track.iTunesIndex > 0)
                            {
                                // Check this track off the list
                                bool exists = checkList.Remove(track.PersistentID);

                                // Do we need to get the track? (it's expensive)
                                if (!exists || ((flags & App.MANAGETICKS) > 0 && !track.DoNotTick && track.Ticked != pl.Ticked))
                                {
                                    // Add the track to the playlist
                                    IITTrack t = itunesTracks.get_ItemByPersistentID(track.PersistentIDHigh, track.PersistentIDLow);
                                    if (t != null)
                                    {
                                        // Add track to playlist if it's not already there
                                        if (!exists)
                                        {
                                            object o = (object)t;
                                            Logger.WritePlaylistAddFile(pl.Path, filename);
                                            itunesPlaylist.AddTrack(ref o);
                                        }

                                        // Untick or tick it if required
                                        if ((flags & App.MANAGETICKS) > 0 && !track.DoNotTick && track.Ticked != pl.Ticked)
                                        {
                                            t.Enabled = pl.Ticked;
                                            track.Ticked = pl.Ticked;
                                        }
                                    }
                                }

                                // Part of the process of managing ticks
                                if ((flags & App.MANAGETICKS) > 0)
                                {
                                    // Untick overrides tick
                                    if (!pl.Ticked && !track.Ticked)
                                        track.DoNotTick = true;
                                    // We've set the tick, no need to check it later
                                    track.VerifyTick = false;
                                }
                            }
                        }
                    }

                    // If there are any items left in the check list
                    // they need deleting
                    if (checkList.Count > 0)
                    {
                        int j = 1, checkedTotal = checkList.Count;

                        UpdateProgress(new ProgressEventArgs(1, checkedTotal, 1));

                        foreach (long persistentID in checkList)
                        {
                            int h, l;
                            GetPersistentIDHighLow(persistentID, out h, out l);

                            // Update the progress text and bar
                            j++;
                            if (DateTime.Now > before.AddSeconds(0.5) || j == checkedTotal)
                            {
                                UpdateProgress(new ProgressEventArgs(j, "{0}: cleaning up track {1} of {2}", pl.Path, j, checkedTotal));
                                before = DateTime.Now;
                            }

                            // Check the current status - whether we need to abort or not
                            CheckCurrentStatus();

                            IITTrack t = itunesPlaylist.Tracks.get_ItemByPersistentID(h, l);
                            if (t != null)
                            {
                                if(t is IITFileOrCDTrack)
                                    Logger.WritePlaylistDeleteFile(pl.Path, ((IITFileOrCDTrack)t).Location);
                                t.Delete();
                            }
                        }
                    }

                    // Shuffle the playlist in iTunes if required
                    if (pl.Shuffle && !pl.Rebuild)
                    {
                        itunesPlaylist.Shuffle = false;
                        itunesPlaylist.Shuffle = true;
                    }
                }

                // Check any tracks not covered by the playlists for ticks
                if ((flags & App.MANAGETICKS) > 0)
                {
                    int count = mcCache.Count;
                    int index = 0;

                    UpdateProgress(new ProgressEventArgs(1, count, 1));

                    foreach (MCTrack track in mcCache.Values)
                    {
                        // Update the progress text and bar
                        if (++index == count || DateTime.Now > before.AddSeconds(0.5))
                        {
                            UpdateProgress(new ProgressEventArgs(index, "Checking ticks for track {0} of {1}", index, count));
                            before = DateTime.Now;
                        }

                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();

                        if (track.VerifyTick && !track.Ticked)
                        {
                            IITTrack t = itunesTracks.get_ItemByPersistentID(track.PersistentIDHigh, track.PersistentIDLow);
                            if (t != null)
                                t.Enabled = true;
                        }
                    }
                }

                // Remove tracks from iTunes Library
                // Go through each playlist
                foreach (PlaylistItem pl in playlistList)
                {
                    if (!pl.Selected || !pl.RemoveTracks)
                        continue;

                    IITUserPlaylist itunesPlaylist = itunesPlaylists[pl.Path];
                    IITTrackCollection tracks = itunesPlaylist.Tracks;
                    int rt = tracks.Count;

                    if (rt <= 0)
                        continue;

                    UpdateProgress(new ProgressEventArgs(1, rt, 1));

                    Dictionary<long, IITTrack> removeTracks = new Dictionary<long, IITTrack>(rt);

                    for (i = 1; i <= rt; i++)
                    {
                        // Update the progress text and bar
                        if (i == rt || DateTime.Now > before.AddSeconds(0.5))
                        {
                            UpdateProgress(new ProgressEventArgs(i, "Preparing to remove track {0} of {1}", i, rt));
                            before = DateTime.Now;
                        }

                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();

                        IITTrack t1 = tracks[i];

                        int h, l;
                        object o = (object)t1;
                        iTunes.GetITObjectPersistentIDs(ref o, out h, out l);

                        IITTrack t2 = itunesTracks.get_ItemByPersistentID(h, l);
                        if(t2 != null)
                            removeTracks[GetPersistentID(h, l)] = t2;
                    }

                    i = 0;

                    foreach (IITTrack t in removeTracks.Values)
                    {
                        // Update the progress text and bar
                        if (++i == rt || DateTime.Now > before.AddSeconds(0.5))
                        {
                            UpdateProgress(new ProgressEventArgs(i, "Removing track {0} of {1}", i, rt));
                            before = DateTime.Now;
                        }

                        // Check the current status - whether we need to abort or not
                        CheckCurrentStatus();

                        try
                        {
                            if ((flags & App.TESTMODE) == 0)
                                t.Delete();
                        }
                        catch { /*Do nothing */ }
                    }
                }
            }

            return new int[] { total, c, importCount, brokenCount };
        }

        /// <summary>
        /// Synchronizes the track
        /// </summary>
        /// <param name="trackSync">The track information</param>
        /// <param name="flags">Fields to sync</param>
        /// <param name="albumCache">Holds all albums</param>
        /// <returns>Whether an error occurred</returns>
        private bool SynchronizeTrack(TrackMonitor trackSync, ulong flags, Dictionary<string, List<MCTrack>> albumCache, bool writeErrorMessages)
        {
            iTunesTrack itunesTrack = trackSync.iTunesTrack;
            MCTrack mcTrack = trackSync.MCTrack;

            bool error = false;

            try
            {
                if ((flags & App.NAME) > 0 && itunesTrack.Name != mcTrack.Name)
                {
                    // Sync the Name in iTunes if different in iTunes
                    string oldval = itunesTrack.Name, newval = mcTrack.Name;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Name = mcTrack.Name;
                    trackSync.SetName(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if(writeErrorMessages)
                    trackSync.SetError("Name", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.ARTIST) > 0 && itunesTrack.Artist != mcTrack.Artist)
                {
                    // Sync the Artist in iTunes if different in iTunes
                    string oldval = itunesTrack.Artist, newval = mcTrack.Artist;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Artist = mcTrack.Artist;
                    trackSync.SetArtist(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Artist", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.ALBUM) > 0 && itunesTrack.Album != mcTrack.Album)
                {
                    // Sync the Album in iTunes if different in iTunes
                    string oldval = itunesTrack.Album, newval = mcTrack.Album;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Album = mcTrack.Album;
                    trackSync.SetAlbum(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Album", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.ALBUMARTIST) > 0 && itunesTrack.AlbumArtist != mcTrack.AlbumArtist)
                {
                    // Sync the Album Artist in iTunes if different in iTunes
                    string oldval = itunesTrack.AlbumArtist, newval = mcTrack.AlbumArtist;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.AlbumArtist = mcTrack.AlbumArtist;
                    trackSync.SetAlbumArtist(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("AlbumArtist", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.GENRE) > 0 && itunesTrack.Genre != mcTrack.Genre)
                {
                    // Sync the Genre in iTunes if different in iTunes
                    string oldval = itunesTrack.Genre, newval = mcTrack.Genre;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Genre = mcTrack.Genre;
                    trackSync.SetGenre(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Genre", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.COMMENT) > 0 && itunesTrack.Comment != mcTrack.Comment)
                {
                    // Sync the Comment in iTunes if different in iTunes
                    string oldval = itunesTrack.Comment, newval = mcTrack.Comment;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Comment = mcTrack.Comment;
                    trackSync.SetComment(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Comment", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.COMPOSER) > 0 && itunesTrack.Composer != mcTrack.Composer)
                {
                    // Sync the Composer in iTunes if different in iTunes
                    string oldval = itunesTrack.Composer, newval = mcTrack.Composer;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Composer = mcTrack.Composer;
                    trackSync.SetComposer(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Composer", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.DISCNUMBER) > 0 && itunesTrack.DiscNumber != mcTrack.DiscNumber)
                {
                    // Sync the DiscNumber in iTunes if different in iTunes
                    int oldval = itunesTrack.DiscNumber, newval = mcTrack.DiscNumber;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.DiscNumber = mcTrack.DiscNumber;
                    trackSync.SetDiscNumber(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("DiscNumber", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.GROUPING) > 0 && itunesTrack.Grouping != mcTrack.Grouping)
                {
                    // Sync the Grouping in iTunes if different in iTunes
                    string oldval = itunesTrack.Grouping, newval = mcTrack.Grouping;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Grouping = mcTrack.Grouping;
                    trackSync.SetGrouping(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Grouping", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.TRACKNUMBER) > 0 && itunesTrack.TrackNumber != mcTrack.TrackNumber)
                {
                    // Sync the TrackNumber in iTunes if different in iTunes
                    int oldval = itunesTrack.TrackNumber, newval = mcTrack.TrackNumber;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.TrackNumber = mcTrack.TrackNumber;
                    trackSync.SetTrackNumber(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("TrackNumber", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.YEAR) > 0 && itunesTrack.Year != mcTrack.Year)
                {
                    // Sync the Year in iTunes if different in iTunes
                    int oldval = itunesTrack.Year, newval = mcTrack.Year;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Year = mcTrack.Year;
                    trackSync.SetYear(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Year", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.DESCRIPTION) > 0 && itunesTrack.Description != mcTrack.Description)
                {
                    // Sync the Description in iTunes if different in iTunes
                    string oldval = itunesTrack.Description, newval = mcTrack.Description;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Description = mcTrack.Description;
                    trackSync.SetDescription(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Description", ex.Message);
                error = true;
            }

            try
            {
                // Strip out linefeeds for the comparison - iTunes inserts them
                string mcLyrics = mcTrack.Lyrics.Replace("\r", string.Empty), itunesLyrics = itunesTrack.Lyrics.Replace("\r", string.Empty);

                if ((flags & App.LYRICS) > 0 && itunesLyrics != mcLyrics)
                {
                    // Sync the Lyrics in iTunes if different in iTunes
                    string oldval = itunesTrack.Lyrics, newval = mcTrack.Lyrics;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Lyrics = mcTrack.Lyrics;
                    trackSync.SetLyrics(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Lyrics", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.BPM) > 0 && itunesTrack.BPM != mcTrack.BPM)
                {
                    // Sync the BPM in iTunes if different in iTunes
                    int oldval = itunesTrack.BPM, newval = mcTrack.BPM;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.BPM = mcTrack.BPM;
                    trackSync.SetBPM(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("BPM", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.COMPILATION) > 0 && mcTrack.Compilation != itunesTrack.Compilation)
                {
                    // Sync the Compilation/Mix Album flag
                    bool oldval = itunesTrack.Compilation, newval = mcTrack.Compilation;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.Compilation = mcTrack.Compilation;
                    trackSync.SetCompilation(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Mix Album", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.RATING) > 0)
                {
                    // Does the rating need syncing?
                    if ((itunesTrack.Rating - (itunesTrack.Rating % 20)) != mcTrack.Rating)
                    {
                        if (_exportOnlyMCInfo)
                        {
                            // If we're just exporting from MC, assign the value
                            int oldval = itunesTrack.Rating, newval = mcTrack.Rating;
                            if ((flags & App.TESTMODE) == 0)
                                itunesTrack.Rating = mcTrack.Rating;
                            trackSync.SetRating(ChangeType.iTunes, oldval, newval);
                        }
                        else if (_exportOnlyiTunesInfo)
                        {
                            // If we're just exporting from iTunes, assign the value
                            int oldval = mcTrack.Rating, newval = itunesTrack.Rating;
                            if ((flags & App.TESTMODE) == 0)
                                mcTrack.Rating = itunesTrack.Rating;
                            trackSync.SetRating(ChangeType.MC, oldval, newval);
                        }
                        // If the rating has never been sync'd
                        else if (mcTrack.iTunesRatingSync == -1)
                        {
                            if (mcTrack.Rating == 0)
                            {
                                // The track is rated in iTunes but not in MC so sync the iTunes value to MC
                                int oldval = mcTrack.Rating, newval = itunesTrack.Rating;
                                if ((flags & App.TESTMODE) == 0)
                                    mcTrack.Rating = itunesTrack.Rating;
                                trackSync.SetRating(ChangeType.MC, oldval, newval);
                            }
                            else
                            {
                                // The rating is different to MC in iTunes so sync the MC value to iTunes
                                int oldval = itunesTrack.Rating, newval = mcTrack.Rating;
                                if ((flags & App.TESTMODE) == 0)
                                    itunesTrack.Rating = mcTrack.Rating;
                                trackSync.SetRating(ChangeType.iTunes, oldval, newval);
                            }
                        }
                        else if (itunesTrack.Rating != mcTrack.iTunesRatingSync && mcTrack.Rating == (mcTrack.iTunesRatingSync - (mcTrack.iTunesRatingSync % 20)))
                        {
                            // The rating has changed in iTunes since the last sync so sync the iTunes value to MC
                            int oldval = mcTrack.Rating, newval = itunesTrack.Rating;
                            if ((flags & App.TESTMODE) == 0)
                                mcTrack.Rating = itunesTrack.Rating;
                            trackSync.SetRating(ChangeType.MC, oldval, newval);
                        }
                        else if (itunesTrack.Rating == mcTrack.iTunesRatingSync)
                        {
                            // The rating has changed in MC since the last sync so sync the MC value to iTunes
                            int oldval = itunesTrack.Rating, newval = mcTrack.Rating;
                            if ((flags & App.TESTMODE) == 0)
                                itunesTrack.Rating = mcTrack.Rating;
                            trackSync.SetRating(ChangeType.iTunes, oldval, newval);
                        }
                    }

                    // Does the custom field need updating?
                    if (mcTrack.iTunesRatingSync != itunesTrack.Rating)
                    {
                        // Set our new value in our custom field also, if necessary
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.iTunesRatingSync = itunesTrack.Rating;
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Rating", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.ALBUMRATING) > 0)
                {
                    // MC doesn't support Album Ratings so just iTunes to MC for this one.
                    // The user must have added the "Album Rating" field themselves for this to work.
                    // Does the album rating need syncing?
                    {
                        int newval = itunesTrack.AlbumRating - (itunesTrack.AlbumRating % 20);
                        if (newval != mcTrack.AlbumRating)
                        {
                            // The album rating has changed in iTunes since the last sync so sync the iTunes value to MC
                            // Set our new value
                            int oldval = mcTrack.AlbumRating;
                            if ((flags & App.TESTMODE) == 0)
                                mcTrack.AlbumRating = newval;
                            trackSync.SetAlbumRating(ChangeType.MC, oldval, newval);
                        }
                    }

                    bool albumRatingCalculated = (itunesTrack.AlbumRating > 0 && itunesTrack.AlbumRatingCalculated);

                    // Does the album rating calculated need syncing
                    if (mcTrack.AlbumRatingCalculated != albumRatingCalculated)
                    {
                        // The album rating calculated has changed in iTunes since the last sync so sync the iTunes value to MC
                        // Set our new value
                        bool oldval = mcTrack.AlbumRatingCalculated, newval = albumRatingCalculated;
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.AlbumRatingCalculated = albumRatingCalculated;
                        trackSync.SetAlbumRatingCalculated(ChangeType.MC, oldval, newval);
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Album Rating", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.LASTPLAYED) > 0 && itunesTrack.LastPlayed != mcTrack.LastPlayed)
                {
                    // Is the iTunes last played earlier than the MC last played
                    if (!_exportOnlyiTunesInfo && (itunesTrack.LastPlayed < mcTrack.LastPlayed || _exportOnlyMCInfo))
                    {
                        // The iTunes last played is earlier, sync the MC value to iTunes
                        DateTime oldval = itunesTrack.LastPlayed, newval = mcTrack.LastPlayed;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.LastPlayed = mcTrack.LastPlayed;
                        trackSync.SetLastPlayed(ChangeType.iTunes, oldval, newval);
                    }
                    else if (!_exportOnlyMCInfo && (itunesTrack.LastPlayed > mcTrack.LastPlayed || _exportOnlyiTunesInfo))
                    {
                        // The iTunes last played is later, sync the iTunes value to MC
                        DateTime oldval = mcTrack.LastPlayed, newval = itunesTrack.LastPlayed;
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.LastPlayed = itunesTrack.LastPlayed;
                        trackSync.SetLastPlayed(ChangeType.MC, oldval, newval);
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("LastPlayed", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.PLAYCOUNT) > 0)
                {
                    // Calculate played count
                    int newPlayedCount;

                    // If we're just exorting from MC, assign the MC value
                    if (_exportOnlyMCInfo)
                        newPlayedCount = mcTrack.PlayCount;
                    // If we're just exorting from iTunes, assign the iTunes value
                    else if (_exportOnlyiTunesInfo)
                        newPlayedCount = itunesTrack.PlayCount;
                    // If we've not done this before simply add the two play counts together to get a total number of plays
                    else if (mcTrack.iTunesPlayedCountSync == -1)
                        newPlayedCount = itunesTrack.PlayCount + mcTrack.PlayCount;
                    else
                    {
                        // We've done this before so minus the value we sync'd last time to get a total number of plays
                        int iTunesNumberOfPlays = itunesTrack.PlayCount - mcTrack.iTunesPlayedCountSync;
                        newPlayedCount = mcTrack.PlayCount + ((iTunesNumberOfPlays < 0) ? 0 : iTunesNumberOfPlays);
                    }

                    // Has the play count changed in MC?
                    if (mcTrack.PlayCount != newPlayedCount)
                    {
                        int oldval = mcTrack.PlayCount, newval = newPlayedCount;
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.PlayCount = newPlayedCount;
                        trackSync.SetPlayCount(ChangeType.MC, oldval, newval);
                    }

                    // Has the play count changed in iTunes?
                    if (itunesTrack.PlayCount != newPlayedCount)
                    {
                        int oldval = itunesTrack.PlayCount, newval = newPlayedCount;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.PlayCount = newPlayedCount;
                        trackSync.SetPlayCount(ChangeType.iTunes, oldval, newval);
                    }

                    if (mcTrack.iTunesPlayedCountSync != itunesTrack.PlayCount)
                    {
                        // Set our new value in our custom field also, if necessary
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.iTunesPlayedCountSync = itunesTrack.PlayCount;
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("PlayCount", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.LASTSKIPPED) > 0 && itunesTrack.LastSkipped != mcTrack.LastSkipped)
                {
                    // Is the iTunes last skipped earlier than the MC last skipped or we are not aggregating
                    if (!_exportOnlyiTunesInfo && (itunesTrack.LastSkipped < mcTrack.LastSkipped || _exportOnlyMCInfo))
                    {
                        // The iTunes last skipped is earlier, sync the MC value to iTunes
                        DateTime oldval = itunesTrack.LastSkipped, newval = mcTrack.LastSkipped;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.LastSkipped = mcTrack.LastSkipped;
                        trackSync.SetLastSkipped(ChangeType.iTunes, oldval, newval);
                    }
                    else if (!_exportOnlyMCInfo && (itunesTrack.LastSkipped > mcTrack.LastSkipped || _exportOnlyiTunesInfo))
                    {
                        // The iTunes last skipped is later, sync the iTunes value to MC
                        DateTime oldval = mcTrack.LastSkipped, newval = itunesTrack.LastSkipped;
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.LastSkipped = itunesTrack.LastSkipped;
                        trackSync.SetLastSkipped(ChangeType.MC, oldval, newval);
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("LastSkipped", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SKIPCOUNT) > 0)
                {
                    // Calculate skipped count
                    int newSkippedCount;

                    // If we're just exporting from MC, assign the MC value
                    if (_exportOnlyMCInfo)
                        newSkippedCount = mcTrack.SkipCount;
                    // If we're just exporting from iTunes, assign the iTunes value
                    else if (_exportOnlyiTunesInfo)
                        newSkippedCount = itunesTrack.SkipCount;
                    // If we've not done this before simply add the two skip counts together to get a total number of skips
                    else if (mcTrack.iTunesSkippedCountSync == -1)
                        newSkippedCount = itunesTrack.SkipCount + mcTrack.SkipCount;
                    else
                    {
                        // We've done this before so minus the value we sync'd last time to get a total number of skips
                        int iTunesNumberOfSkips = itunesTrack.SkipCount - mcTrack.iTunesSkippedCountSync;
                        newSkippedCount = mcTrack.SkipCount + ((iTunesNumberOfSkips < 0) ? 0 : iTunesNumberOfSkips);
                    }

                    // Has the skip count changed in MC?
                    if (mcTrack.SkipCount != newSkippedCount)
                    {
                        int oldval = mcTrack.SkipCount, newval = newSkippedCount;
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.SkipCount = newSkippedCount;
                        trackSync.SetSkipCount(ChangeType.MC, oldval, newval);
                    }

                    // Has the skip count changed in iTunes?
                    if (itunesTrack.SkipCount != newSkippedCount)
                    {
                        int oldval = itunesTrack.SkipCount, newval = newSkippedCount;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.SkipCount = newSkippedCount;
                        trackSync.SetSkipCount(ChangeType.iTunes, oldval, newval);
                    }

                    if (mcTrack.iTunesSkippedCountSync != itunesTrack.SkipCount)
                    {
                        // Set our new value in our custom field also, if necessary
                        if ((flags & App.TESTMODE) == 0)
                            mcTrack.iTunesSkippedCountSync = itunesTrack.SkipCount;
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SkipCount", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & (App.DISCCOUNT | App.TRACKCOUNT)) > 0)
                {
                    int trackCount = mcTrack.TrackNumber;
                    int discCount = 0;

                    List<MCTrack> album = null;

                    try
                    {
                        album = albumCache[mcTrack.AlbumKey];
                    }
                    catch {/* Do nothing */}

                    if (album != null)
                    {
                        foreach (MCTrack track in album)
                        {
                            // If this is indeed a genuine fellow track (must have the same disc number)
                            if (track.DiscNumber == mcTrack.DiscNumber)
                            {
                                // Is the track number of the fellow greater than the track number we are holding
                                if (trackCount < track.TrackNumber)
                                    trackCount = track.TrackNumber;
                            }

                            // Is the disc number of the fellow greater than the disc number we are holding
                            if (discCount < track.DiscNumber)
                                discCount = track.DiscNumber;
                        }
                    }

                    // Set the track count, if necessary
                    if ((flags & App.TRACKCOUNT) > 0 && itunesTrack.TrackCount != trackCount)
                    {
                        int oldval = itunesTrack.TrackCount, newval = trackCount;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.TrackCount = trackCount;
                        trackSync.SetTrackCount(ChangeType.iTunes, oldval, newval);
                    }

                    // Set the disc count, if necessary
                    if ((flags & App.DISCCOUNT) > 0 && itunesTrack.DiscCount != discCount)
                    {
                        int oldval = itunesTrack.DiscCount, newval = discCount;
                        if ((flags & App.TESTMODE) == 0)
                            itunesTrack.DiscCount = discCount;
                        trackSync.SetDiscCount(ChangeType.iTunes, oldval, newval);
                    }
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("Disc & Track Counts", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SORTNAME) > 0 && itunesTrack.SortName != mcTrack.SortName)
                {
                    // Sync the Sort Name in iTunes if different in iTunes
                    string oldval = itunesTrack.SortName, newval = mcTrack.SortName;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.SortName = mcTrack.SortName;
                    trackSync.SetSortName(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SortName", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SORTARTIST) > 0 && itunesTrack.SortArtist != mcTrack.SortArtist)
                {
                    // Sync the Sort Artist in iTunes if different in iTunes
                    string oldval = itunesTrack.SortArtist, newval = mcTrack.SortArtist;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.SortArtist = mcTrack.SortArtist;
                    trackSync.SetSortArtist(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SortArtist", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SORTALBUMARTIST) > 0 && itunesTrack.SortAlbumArtist != mcTrack.SortAlbumArtist)
                {
                    // Sync the Sort Album Artist in iTunes if different in iTunes
                    string oldval = itunesTrack.SortAlbumArtist, newval = mcTrack.SortAlbumArtist;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.SortAlbumArtist = mcTrack.SortAlbumArtist;
                    trackSync.SetSortAlbumArtist(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SortAlbumArtist", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SORTALBUM) > 0 && itunesTrack.SortAlbum != mcTrack.SortAlbum)
                {
                    // Sync the Sort Album in iTunes if different in iTunes
                    string oldval = itunesTrack.SortAlbum, newval = mcTrack.SortAlbum;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.SortAlbum = mcTrack.SortAlbum;
                    trackSync.SetSortAlbum(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SortAlbum", ex.Message);
                error = true;
            }

            try
            {
                if ((flags & App.SORTCOMPOSER) > 0 && itunesTrack.SortComposer != mcTrack.SortComposer)
                {
                    // Sync the Sort Composer in iTunes if different in iTunes
                    string oldval = itunesTrack.SortComposer, newval = mcTrack.SortComposer;
                    if ((flags & App.TESTMODE) == 0)
                        itunesTrack.SortComposer = mcTrack.SortComposer;
                    trackSync.SetSortComposer(ChangeType.iTunes, oldval, newval);
                }
            }
            catch (Exception ex)
            {
                if (writeErrorMessages)
                    trackSync.SetError("SortComposer", ex.Message);
                error = true;
            }

            return error;
        }

        /// <summary>
        /// Returns the MC library
        /// </summary>
        /// <param name="mc">MC object</param>
        /// <param name="mcCache">List to add tracks to</param>
        /// <param name="mcAlbumCache">List to add albums to</param>
        /// <param name="syncFileTypesList">Types of files to sync</param>
        /// <param name="flags">User settings</param>
        /// <param name="viewSchemeName">The MC view scheme that should be used</param>
        /// <param name="playListName">The MC playlist that should be used</param>
        /// <param name="expressions">MC Expressions</param>
        private void GetMCCache(MCAutomation mc, ref Dictionary<string, MCTrack> mcCache, ref Dictionary<string, List<MCTrack>> mcAlbumCache, List<string> syncFileTypesList, ulong flags, string viewSchemeName, string playListName, Dictionary<string, string> expressions)
        {
            // Set the expressions
            MCTrack.RefreshTemplate(flags, expressions);

            // Progress variables
            DateTime before = DateTime.MinValue;
            int i = 0;
            int total = 0;

            // The files we are going to cache
            IMJFilesAutomation mediaLibrary = null;

            // Has a playlist been specified?
            if (playListName != null && playListName != string.Empty)
            {
                IMJPlaylistsAutomation playlists = mc.GetPlaylists();
                for (int j = 0; j < playlists.GetNumberPlaylists(); j++)
                {
                    IMJPlaylistAutomation playlist = playlists.GetPlaylist(j);
                    string path = ((playlist.Path == string.Empty) ? string.Empty:playlist.Path + '\\') + playlist.Name;
                    if (path == playListName)
                    {
                        mediaLibrary = playlist.GetFiles();
                        break;
                    }
                }
            }

            if (mediaLibrary == null)
            {
                // Gets the view scheme (if viewSchemeName is empty, gets the entire library)
                IMJSchemeAutomation viewScheme = mc.GetViewScheme(viewSchemeName);
                mediaLibrary = viewScheme.GetFiles();
            }

            total = mediaLibrary.GetNumberFiles();

            // Contains MC Library
            mcCache = new Dictionary<string, MCTrack>(total);
            // Contains MC Albums
            if ((flags & (App.ALBUMRATING | App.DISCCOUNT | App.TRACKCOUNT)) > 0)
                mcAlbumCache = new Dictionary<string, List<MCTrack>>();

            // If there's nothing in the MC library, return
            if (total <= 0)
                return;

            UpdateProgress(new ProgressEventArgs("Loading MC Library..."));
            UpdateProgress(new ProgressEventArgs(1, total, 1));

            for (i = 0; i < total; i++)
            {
                // Update the progress text and bar
                int j = i + 1;
                if (DateTime.Now > before.AddSeconds(0.5) || j == total)
                {
                    UpdateProgress(new ProgressEventArgs(j, "Loading MC Library... {0} of {1}", j, total));
                    before = DateTime.Now;
                }

                // Check the current status - whether we need to abort or not
                CheckCurrentStatus();

                IMJFileAutomation file = mediaLibrary.GetFile(i);
                if (file == null || file.Filename == null || file.Filename == string.Empty || ((flags & App.SYNCCERTAINFILETYPES) > 0 && syncFileTypesList != null && !syncFileTypesList.Contains(file.Filetype.ToUpper())))
                    continue;

                MCTrack track = null;

                try { track = new MCTrack(file, flags); }
                catch (Exception ex)
                {
                    Logger.WriteMCCacheError(file.Artist, file.Album, file.Name, ex.Message);
                    continue;
                }

                if ((flags & App.USEALTERNATIVEFILEKEY) > 0 && track.iTunesFileKey != string.Empty)
                {
                    // Does the file exist?
                    if ((flags & App.LOGFILEKEYWARNING) > 0 && !File.Exists(track.iTunesFileKey))
                        Logger.WriteFileKeyWarning(track.iTunesFileKey, "iTunesFileKey entry does not exist on disk");
                }

                // Part of an album?
                if (mcAlbumCache != null && track.TrackNumber > 0 && track.Album != string.Empty)
                {
                    List<MCTrack> album = null;

                    try
                    {
                        album = mcAlbumCache[track.AlbumKey];
                    }
                    catch {/* Do nothing */}

                    if (album == null)
                    {
                        album = new List<MCTrack>();
                        album.Add(track);
                        mcAlbumCache.Add(track.AlbumKey, album);
                    }
                    else
                    {
                        album.Add(track);
                    }
                }

                try { mcCache.Add(track.Key, track); }
                catch (ArgumentException) 
                {
                    Logger.WriteFileKeyWarning(track.Key, "Duplicate MC cache key encountered");
                }
            }
        }

        /// <summary>
        /// Creates our custom fields
        /// </summary>
        /// <param name="mc">MC object</param>
        /// <param name="flags">User settings</param>
        private void CreateCustomFields(MCAutomation mc, ulong flags)
        {
            // Add our fields into MC libary
            // If they already exist MC does nothing
            IMJFieldsAutomation fields = mc.GetFields();
            if ((flags & App.RATING) > 0)
                fields.CreateFieldSimple("iTunesRatingSync", "iTunesRatingSync", 0, 0);
            if ((flags & App.PLAYCOUNT) > 0)
                fields.CreateFieldSimple("iTunesPlayedCountSync", "iTunesPlayedCountSync", 0, 0);
            if ((flags & App.SKIPCOUNT) > 0)
                fields.CreateFieldSimple("iTunesSkippedCountSync", "iTunesSkippedCountSync", 0, 0);
            if ((flags & App.USEALTERNATIVEFILEKEY) > 0)
                fields.CreateFieldSimple("iTunesFileKey", "iTunesFileKey", 1, 0);
            if ((flags & App.ALBUMRATING) > 0)
            {
                fields.CreateFieldSimple("iTunesAlbumRating", "iTunesAlbumRating", 0, 0);
                fields.CreateFieldSimple("Album Rating", "Album Rating", 0, 0);
                fields.CreateFieldSimple("Album Rating Calculated", "Album Rating Calculated", 0, 0);
            }
            fields = null;
        }

        /// <summary>
        /// Checks the current status and deals with any change as appropriate
        /// </summary>
        private void CheckCurrentStatus()
        {
            // Has the synchronization been aborted?
            switch (Running)
            {
                case StatusEnum.CloseRequest:
                    throw new UserAbortException("User exited the application.");
                case StatusEnum.StopRequest:
                    throw new UserAbortException("User aborted.");
                case StatusEnum.ShutdownRequest:
                    throw new UserAbortException("System shutting down.");
            }
        }

        /// <summary>
        /// Gets a reference to MC and changes to a specified library
        /// </summary>
        /// <param name="libaryInfo">Information used to retrieve the specified MC library</param>
        /// <returns>A reference to MC</returns>
        public static MCAutomation GetJRMediaCenter(object[] libraryInfo)
        {
            MCAutomation mc = null;

            // Is MC already running?
            try
            {
                mc = (MCAutomation)Marshal.GetActiveObject("MediaJukebox Application");
            }
            catch {/* Do nothing */}

            // Do we need to instantiate MC?
            if (mc == null)
            {
                try { mc = new MCAutomationClass(); }
                catch (Exception ex) { throw new Exception(string.Format("Could not access J River Media Center. {0}", ex.Message)); }
            }

            MCVersion = new Version(mc.GetVersion().Version);

            // Is the library specified?
            if (libraryInfo != null && libraryInfo[0] != null)
            {
                string libraryName = ((string)libraryInfo[0]).ToUpper();

                // Get current library
                string name = string.Empty, path = string.Empty;
                mc.GetLibrary(ref name, ref path);
                IMJVersionAutomation ver = mc.GetVersion();

                // If the current library is the one we're after, don't bother
                if (libraryName != name.ToUpper())
                {
                    RegistryKey cuser = Registry.CurrentUser;
                    RegistryKey libkey = cuser.OpenSubKey(string.Format("Software\\J. River\\Media Center {0}\\Properties\\Libraries", ver.Major), false);

                    Dictionary<int, string> regLibraries = new Dictionary<int, string>();

                    if (ver.Major >= 16)
                    {
                        if (libkey == null)
                        {
                            regLibraries[0] = name;
                        }
                        else
                        {
                            string[] libnames = libkey.GetValueNames();

                            foreach (string n in libnames)
                            {
                                int libidx;
                                string regname = string.Empty;

                                if (!int.TryParse(n, out libidx))
                                    continue;

                                string p = (string)libkey.GetValue(n);

                                XmlDocument xml = new XmlDocument();
                                xml.LoadXml(p);

                                foreach (XmlElement item in xml.GetElementsByTagName("Item"))
                                {
                                    if (!item.HasAttributes)
                                        continue;

                                    XmlAttribute itemAttr = item.Attributes["Name"];
                                    if (itemAttr.Value == "DisplayName")
                                    {
                                        regname = item.InnerText;
                                        break;
                                    }
                                }

                                regLibraries[libidx] = regname;
                            }
                        }
                    }
                    else
                    {
                        string[] libnames = libkey.GetValueNames();

                        int libidx = 0;
                        foreach (string n in libnames)
                        {
                            string p = (string)libkey.GetValue(n);
                            if (p.Contains("localservers_autofind"))
                                continue;

                            regLibraries[libidx++] = n;
                        }
                    }

                    int selidx = 0;

                    foreach (int i in regLibraries.Keys)
                    {
                        string n = regLibraries[i];

                        if (n.ToUpper() == libraryName)
                        {
                            selidx = i;
                            break;
                        }
                    }

                    // Did we find the library?
                    if(selidx >= 0)
                    {
                        // Now change the library
                        PostMessage(new IntPtr(mc.GetWindowHandle()), MC.WM_MC_COMMAND, MC.MCC_LOAD_LIBRARY, (uint)selidx);

                        int i = Settings.LibraryLoadRetries;//timeout
                        mc.GetLibrary(ref name, ref path);
                        while (libraryName != name.ToUpper())
                        {
                            mc.GetLibrary(ref name, ref path);
                            Thread.Sleep(Settings.WaitOnLibraryLoadError);
                            i--;
                            if (i == 0)
                                throw new Exception(string.Format("Timeout searching for specified library (index {0}).", selidx));
                        }
                    }
                }
            }

            return mc;
        }

        /// <summary>
        /// Gets a reference to iTunes
        /// </summary>
        /// <returns>A reference to iTunes</returns>
        public static iTunesApp GetiTunes()
        {
            iTunesApp iTunes = null;

            try { iTunes = new iTunesAppClass(); }
            catch (Exception ex) { throw new Exception(string.Format("Could not access iTunes. {0}", ex.Message)); }

            iTunesVersion = new Version(iTunes.Version);

            return iTunes;
        }

        /// <summary>
        /// Gets all playlists in MC
        /// </summary>
        /// <param name="mc">The MC object</param>
        /// <param name="returnGroups">Whether to return playlist groups</param>
        /// <returns>A keyed list of playlists in MC</returns>
        public static Dictionary<string, IMJPlaylistAutomation> GetMCPlaylists(MCAutomation mc, bool returnGroups)
        {
            IMJPlaylistsAutomation mcPlaylists = mc.GetPlaylists();
            int total = mcPlaylists.GetNumberPlaylists();

            Dictionary<string, IMJPlaylistAutomation> ret = new Dictionary<string, IMJPlaylistAutomation>();

            for (int i = 0; i < total; i++)
            {
                IMJPlaylistAutomation mcPlaylist = mcPlaylists.GetPlaylist(i);

                // Get playlist type
                MC.PLAYLIST_TYPES type = (MC.PLAYLIST_TYPES)(-1);
                try { type = (MC.PLAYLIST_TYPES)int.Parse(mcPlaylist.Get("Type")); }
                catch {/* Do nothing */}

                switch (type)
                {
                    case MC.PLAYLIST_TYPES.PLAYLIST_GROUP:
                        {
                            if (returnGroups)
                                goto case MC.PLAYLIST_TYPES.PLAYLIST;
                            break;
                        }
                    case MC.PLAYLIST_TYPES.SMARTLIST:
                    case MC.PLAYLIST_TYPES.PLAYLIST:
                        {
                            string plkey = (mcPlaylist.Path == string.Empty) ? mcPlaylist.Name : (mcPlaylist.Path + '\\' + mcPlaylist.Name);
                            if(!ret.ContainsKey(plkey))
                                ret.Add(plkey, mcPlaylist);
                            break;
                        }
                }
            }

            return ret;
        }

        /// <summary>
        /// Gets/creates playlist in iTunes
        /// </summary>
        /// <param name="iTunes">iTunes object</param>
        /// <param name="iTunesPlaylistsHash">List of iTunes playlists</param>
        /// <param name="playlistPath">Path of the playlist</param>
        /// <param name="isFolder">Whether this is a playlist group</param>
        /// <returns></returns>
        private IITUserPlaylist GetPlaylist(iTunesApp iTunes, Dictionary<string, IITUserPlaylist> iTunesPlaylistsHash, string playlistPath, bool isFolder)
        {
            IITUserPlaylist p = null;

            if (iTunesPlaylistsHash.TryGetValue(playlistPath, out p))
            {
                if (isFolder && p.SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindFolder)
                    return p;
                else if (!isFolder && p.SpecialKind == ITUserPlaylistSpecialKind.ITUserPlaylistSpecialKindNone)
                    return p;
                else
                    p = null;
            }

            string path = string.Empty, name;
            int n = playlistPath.LastIndexOf('\v');
            if (n >= 0)
            {
                path = playlistPath.Substring(0, playlistPath.Length - (playlistPath.Length - n));
                name = playlistPath.Substring(n + 1);
            }
            else
            {
                path = string.Empty;
                name = playlistPath;
            }

            if (path != string.Empty)
            {
                p = GetPlaylist(iTunes, iTunesPlaylistsHash, path, true);

                if (isFolder)
                    p = (IITUserPlaylist)p.CreateFolder(name);
                else
                    p = (IITUserPlaylist)p.CreatePlaylist(name);

                iTunesPlaylistsHash[path + '\v' + name] = p;
            }
            else
            {
                if (isFolder)
                    p = (IITUserPlaylist)iTunes.CreateFolder(name);
                else
                    p = (IITUserPlaylist)iTunes.CreatePlaylist(name);

                iTunesPlaylistsHash[name] = p;
            }
            
            return p;
        }

        /// <summary>
        /// Combines a high and a low integer into a long
        /// </summary>
        /// <param name="h">The high integer</param>
        /// <param name="l">The low integer</param>
        /// <returns>The combination of the high and low integers</returns>
        private static long GetPersistentID(int h, int l)
        {
            long nh = h, nl = l;
            nh <<= 32;
            nl &= 0xffffffff;

            return (nh | nl);
        }

        /// <summary>
        /// Turns a long into a high and a low integer
        /// </summary>
        /// <param name="val">The long value</param>
        /// <param name="h">The high integer</param>
        /// <param name="l">The low integer</param>
        private static void GetPersistentIDHighLow(long val, out int h, out int l)
        {
            long nh = val >> 32;
            long nl = val & 0xffffffff;

            h = (int)nh;
            l = (int)nl;
        }
    }
}
