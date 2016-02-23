using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using MediaCenter;
using System.Runtime.InteropServices;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// The main form that the contains much of the work
    /// </summary>
    public partial class MCiTunesSynchronizerForm : Form
    {
        /// <summary>
        /// Whether to start the sync once the program starts
        /// </summary>
        protected bool _startSyncOnStartup;

        /// <summary>
        /// Whether to exit the program once the sync has finished
        /// </summary>
        protected readonly bool _exitAfterSync;

        /// <summary>
        /// MC Library information passed from command line
        /// </summary>
        protected object[] _libraryInformation;

        /// <summary>
        /// Whether to show hidden playlists in the tree
        /// </summary>
        private bool _showHiddenPlaylists;

        /// <summary>
        /// Keeps track of contents of playlist tree
        /// </summary>
        protected Dictionary<string, PlaylistItem> _playlistCache = new Dictionary<string, PlaylistItem>();

        /// <summary>
        /// The object that performs the synchronization.
        /// </summary>
        protected Synchronizer _synchronizer;

        /// <summary>
        /// Creates and initializes the main form
        /// </summary>
        /// <param name="startSync">Whether to start the sync immediately</param>
        /// <param name="exitAfterSync">Whether to exit the program once sync is finished</param>
        /// <param name="exportOnlyMCInfo">Whether to export only MC to iTunes - no aggregation</param>
        /// <param name="exportOnlyiTunesInfo">Whether to export only iTunes to MC - no aggregation</param>
        /// <param name="libraryInfo">Info about the MC library to use</param>
        internal MCiTunesSynchronizerForm(bool startSync, bool exitAfterSync, bool exportOnlyMCInfo, bool exportOnlyiTunesInfo, object[] libraryInfo)
        {
            InitializeComponent();

            // Assign the arguments we need
            _startSyncOnStartup = startSync;
            _exitAfterSync = exitAfterSync;
           _libraryInformation = libraryInfo;

            // Set the window title
            this.Text = string.Format("{0} {1}", AboutBox.AssemblyTitle, AboutBox.AssemblyConfiguration);

            // Set the window size
            if (Settings.MainWindowHeight > 200 && Settings.MainWindowWidth > 200)
            {
                this.Height = Settings.MainWindowHeight;
                this.Width = Settings.MainWindowWidth;
            }

            // Instantiate the synchronizer and set up the event
            _synchronizer = new Synchronizer(exportOnlyMCInfo, exportOnlyiTunesInfo);
            _synchronizer.UpdateProgress += new Synchronizer.ProgressEventHandler(UpdateProgressSink);

            // Set the config settings
            ulong flags = Settings.Options;

            nameCheck.Checked = (flags & App.NAME) > 0;
            artistCheck.Checked = (flags & App.ARTIST) > 0;
            albumCheck.Checked = (flags & App.ALBUM) > 0;
            albumArtistCheck.Checked = (flags & App.ALBUMARTIST) > 0;
            composerCheck.Checked = (flags & App.COMPOSER) > 0;
            genreCheck.Checked = (flags & App.GENRE) > 0;
            commentCheck.Checked = (flags & App.COMMENT) > 0;
            groupingCheck.Checked = (flags & App.GROUPING) > 0;
            descriptionCheck.Checked = (flags & App.DESCRIPTION) > 0;
            discNumberCheck.Checked = (flags & App.DISCNUMBER) > 0;
            trackNumberCheck.Checked = (flags & App.TRACKNUMBER) > 0;
            yearCheck.Checked = (flags & App.YEAR) > 0;
            lyricsCheck.Checked = (flags & App.LYRICS) > 0;
            ratingCheck.Checked = (flags & App.RATING) > 0;
            lastPlayedCheck.Checked = (flags & App.LASTPLAYED) > 0;
            playCountCheck.Checked = (flags & App.PLAYCOUNT) > 0;
            skipCountCheck.Checked = (flags & App.SKIPCOUNT) > 0;
            lastSkippedCheck.Checked = (flags & App.LASTSKIPPED) > 0;
            compilationCheck.Checked = (flags & App.COMPILATION) > 0;
            bpmCheck.Checked = (flags & App.BPM) > 0;
            discCountCheck.Checked = (flags & App.DISCCOUNT) > 0;
            trackCountCheck.Checked = (flags & App.TRACKCOUNT) > 0;
            albumRatingCheck.Checked = (flags & App.ALBUMRATING) > 0;
            playlistCheck.Checked = (flags & App.PLAYLISTS) > 0;
            playlistTree.Enabled = playlistCheck.Checked;
            refreshTreeButton.Enabled = playlistCheck.Checked;
            checkAllTreeButton.Enabled = playlistCheck.Checked;
            uncheckAllTreeButton.Enabled = playlistCheck.Checked;
            customFolderText.Text = Settings.RootPlaylistFolder;
            customFolderText.Enabled = playlistCheck.Checked;
            _showHiddenPlaylists = (flags & App.SHOWHIDDENPLAYLISTS) > 0;
            toggleHiddenButton.Enabled = playlistCheck.Checked;
            sortNameCheck.Checked = (flags & App.SORTNAME) > 0;
            sortArtistCheck.Checked = (flags & App.SORTARTIST) > 0;
            sortAlbumArtistCheck.Checked = (flags & App.SORTALBUMARTIST) > 0;
            sortAlbumCheck.Checked = (flags & App.SORTALBUM) > 0;
            sortComposerCheck.Checked = (flags & App.SORTCOMPOSER) > 0;
            onTheGoCheck.Checked = (flags & App.ONTHEGOPLAYLISTS) > 0;
            manageTicksCheck.Enabled = playlistCheck.Checked;
            manageTicksCheck.Checked = (flags & App.MANAGETICKS) > 0;

            if((flags & App.USEPLAYLISTROOT) > 0)
            {
                rootFolderRadio.Checked = true;
                customFolderRadio.Checked = false;
            }
            else
            {
                rootFolderRadio.Checked = false;
                customFolderRadio.Checked = true;
            }

            // Set tool tips for checkboxes that require them
            onTheGoCheck.Tag = "Synchronize all On-The-Go playlists from iTunes to MC into a playlist folder named \"On-The-Go\" in MC.\nIf the playlist folder \"On-The-Go\" does not exist in MC, MCiS will create it.";
            playlistCheck.Tag = "Synchronize all checked playlists in the tree from MC to iTunes into the specified iTunes playlist folder.";
            rootFolderRadio.Tag = "Create playlists directly under the Playlists section in iTunes.\nPlaylists will be removed from iTunes where a playlist exists in MC but has not been selected for synchronization.\nAlso, a playlist will be removed from iTunes where the parent folder is selected for synchronize and the playlist does not exist in MC.";
            customFolderRadio.Tag = "Create playlists under the specified playlist folder in iTunes. If this playlist folder does not exist it will be created.\nThe playlist folder structure in iTunes will be an exact match with that selected.\nAny playlists that already exist in the specified iTunes folder but are not selected for synchronize will be removed from iTunes.";
            manageTicksCheck.Tag = "If you right-click a playlist and set \"Untick\", all tracks in that playlist in iTunes will be unticked.\nAll other tracks in your iTunes library will be ticked.";
            App.SetToolTip(onTheGoCheck, true);
            App.SetToolTip(playlistCheck, true);
            App.SetToolTip(rootFolderRadio, true);
            App.SetToolTip(customFolderRadio, true);
            App.SetToolTip(manageTicksCheck, true);

            checkAllAggregateButton.Tag = "Check All";
            checkAllExportButton.Tag = "Check All";
            checkAllTreeButton.Tag = "Check All";
            uncheckAllAggregateButton.Tag = "Uncheck All";
            uncheckAllExportButton.Tag = "Uncheck All";
            uncheckAllTreeButton.Tag = "Uncheck All";
            refreshTreeButton.Tag = "Refresh the tree with any playlist changes you have made in MC since MCiS startup";
            toggleHiddenButton.Tag = "Hide/show playlists in the tree that have been marked as hidden.\nHidden playlists are not synchronized, even if checked.";
            App.SetToolTip(checkAllAggregateButton, false);
            App.SetToolTip(checkAllExportButton, false);
            App.SetToolTip(checkAllTreeButton, false);
            App.SetToolTip(uncheckAllAggregateButton, false);
            App.SetToolTip(uncheckAllExportButton, false);
            App.SetToolTip(uncheckAllTreeButton, false);
            App.SetToolTip(refreshTreeButton, false);
            App.SetToolTip(toggleHiddenButton, false);

            // Set tool tips for groups
            if (exportOnlyMCInfo)
            {
                aggregateSyncGrp.Text = "Fields to export from MC to iTunes";
                aggregateSyncGrp.Tag = "These fields are exported from MC to iTunes if the values are different in iTunes.";
            }
            else if (exportOnlyiTunesInfo)
            {
                aggregateSyncGrp.Text = "Fields to export from iTunes to MC";
                aggregateSyncGrp.Tag = "These fields are exported from iTunes to MC if the values are different in MC.";
            }
            else
                aggregateSyncGrp.Tag = "These fields are calculated, i.e. the most up-to-date values are calculated and then saved to MC, iTunes or both.";

            syncToiTunesGrp.Tag = "These fields are exported from MC to iTunes if the values are different in iTunes.";
            syncToMCGrp.Tag = "These fields are exported from iTunes to MC if the values are different in MC.";
            App.SetToolTip(aggregateSyncGrp, false);
            App.SetToolTip(syncToiTunesGrp, false);
            App.SetToolTip(syncToMCGrp, false);

            // Set tool tips for buttons
            openButton.Tag = "Open the XML log";
            App.SetToolTip(openButton, true);

            // Set the controls to their defaults
            SetControlDefaults();

            // Fill the tree
            RefreshTree(Settings.Playlists);
        }

        /// <summary>
        /// Executes when progress is updated in the synchronizer thread
        /// </summary>
        /// <param name="e">Event data</param>
        void UpdateProgressSink(ProgressEventArgs e)
        {
            // We need to invoke this on the thread under which
            // the controls we are updating are runnning
            Synchronizer.ProgressEventHandler func = new Synchronizer.ProgressEventHandler(UpdateProgress);
            Invoke(func, e);
        }

        /// <summary>
        /// Updates the pertinent progress controls according to the event data
        /// </summary>
        /// <param name="e">Event data</param>
        void UpdateProgress(ProgressEventArgs e)
        {
            // Is the synchronization finished?
            if (e.Finished)
            {
                // It has, set the controls back to default
                SetControlDefaults();
                // If required send the resulting xml to the clipboard
                if (e.ResultingXML != null)
                {
                    try { Clipboard.SetText(e.ResultingXML); }
                    catch {/* Do nothing */}
                }

                // Exit if required
                if ((e.FinalStatus == StatusEnum.FinishingSuccessful && _exitAfterSync) || e.FinalStatus == StatusEnum.CloseRequest || e.FinalStatus == StatusEnum.ShutdownRequest)
                    this.Close();
            }
            else
            {
                // Set the relevant controls to show the pertinent information
                if (e.ProgressText != string.Empty)
                    progressLabel.Text = e.ProgressText.Replace("&", "&&");
                if (e.ProgressMin >= 0)
                    progressBar.Minimum = e.ProgressMin;
                if (e.ProgressMax >= 0)
                    progressBar.Maximum = e.ProgressMax;
                if (e.ProgressValue >= 0)
                {
                    if (e.ProgressValue < progressBar.Minimum)
                        progressBar.Value = progressBar.Minimum;
                    else if (e.ProgressValue > progressBar.Maximum)
                        progressBar.Value = progressBar.Maximum;
                    else
                        progressBar.Value = e.ProgressValue;
                }
            }
        }

        /// <summary>
        /// Sets controls to default values
        /// </summary>
        private void SetControlDefaults()
        {
            progressBar.Value = progressBar.Minimum;
            goButton.Text = "Synchroni&ze";
            goButton.ImageIndex = 0;
            EnableControls(true);
        }

        /// <summary>
        /// Enables or disables form controls
        /// </summary>
        /// <param name="enable">Whether to enable or disable the controls</param>
        private void EnableControls(bool enable)
        {
            goButton.Enabled = true;// Always true

            nameCheck.Enabled = enable;
            artistCheck.Enabled = enable;
            albumCheck.Enabled = enable;
            albumArtistCheck.Enabled = enable;
            composerCheck.Enabled = enable;
            genreCheck.Enabled = enable;
            commentCheck.Enabled = enable;
            groupingCheck.Enabled = enable;
            descriptionCheck.Enabled = enable;
            discNumberCheck.Enabled = enable;
            trackNumberCheck.Enabled = enable;
            yearCheck.Enabled = enable;
            lyricsCheck.Enabled = enable;
            ratingCheck.Enabled = enable;
            lastPlayedCheck.Enabled = enable;
            playCountCheck.Enabled = enable;
            skipCountCheck.Enabled = enable;
            lastSkippedCheck.Enabled = enable;
            compilationCheck.Enabled = enable;
            bpmCheck.Enabled = enable;
            discCountCheck.Enabled = enable;
            trackCountCheck.Enabled = enable;
            checkAllExportButton.Enabled = enable;
            checkAllAggregateButton.Enabled = enable;
            albumRatingCheck.Enabled = enable;
            openButton.Enabled = (enable && File.Exists(App.LOGFILENAME)); 
            playlistCheck.Enabled = enable;
            rootFolderRadio.Enabled = (enable && playlistCheck.Checked);
            customFolderRadio.Enabled = (enable && playlistCheck.Checked);
            playlistTree.Enabled = (enable && playlistCheck.Checked);
            manageTicksCheck.Enabled = (enable && playlistCheck.Checked);
            refreshTreeButton.Enabled = (enable && playlistCheck.Checked);
            checkAllTreeButton.Enabled = (enable && playlistCheck.Checked);
            uncheckAllTreeButton.Enabled = (enable && playlistCheck.Checked);
            customFolderText.Enabled = (enable && playlistCheck.Checked && customFolderRadio.Checked);
            toggleHiddenButton.Enabled = (enable && playlistCheck.Checked);
            sortNameCheck.Enabled = enable;
            sortArtistCheck.Enabled = enable;
            sortAlbumArtistCheck.Enabled = enable;
            sortAlbumCheck.Enabled = enable;
            sortComposerCheck.Enabled = enable;
            onTheGoCheck.Enabled = enable;
            aggregateSyncGrp.Enabled = enable;
            syncToiTunesGrp.Enabled = enable;
            syncToMCGrp.Enabled = enable;
            playlistGroup.Enabled = enable;
        }

        /// <summary>
        /// If the user tries to close the program while the thread is running
        /// we need to stop the thread first.
        /// </summary>
        /// <param name="e">Event data</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Is the worker thread currently running?
            if (_synchronizer.Running != StatusEnum.Idle)
            {
                // It is, so stop the thread and deny the close request
                switch(e.CloseReason)
                {
                    case CloseReason.WindowsShutDown:
                        _synchronizer.Running = StatusEnum.ShutdownRequest;
                        break;
                    default:
                        _synchronizer.Running = StatusEnum.CloseRequest;
                        break;
                }
                e.Cancel = true;
            }
            else
            {
                // The worker thread is not running so we can safely exit.
                // Save the configuration settings

                // Set the check box config
                Settings.Options = GetOptionFlags(true);
                Settings.RootPlaylistFolder = customFolderText.Text;

                // Set the settings to the cache
                foreach (PlaylistItem pl in _playlistCache.Values)
                {
                    Settings.Playlists[pl.Path] = pl;
                }

                // Finally, save the configuration
                Settings.Save(this.Height, this.Width);
            }

            // Bye bye
            base.OnFormClosing(e);
        }

        /// <summary>
        /// This either starts or stops the synchronization process
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void goButton_Click(object sender, EventArgs e)
        {
            goButton.Enabled = false;
            if (customFolderText.Text.Trim() == string.Empty)
                customFolderText.Text = "MC Playlists";
            StartOrStopSync();
        }
        
        /// <summary>
        /// This either starts or stops the synchronization process
        /// </summary>
        private void StartOrStopSync()
        {
            StatusEnum status = _synchronizer.Running;

            // If the worker thread is currently running we're
            // stopping the process, otherwise we're starting
            if (status == StatusEnum.Idle)
            {
                ulong flags = GetOptionFlags(false);
                
                // Make sure some fields are selected, otherwise we have nothing to sync
                if ((flags & App.FIELDFLAGS) > 0)
                {
                    // Disable controls
                    EnableControls(false);

                    Dictionary<string, PlaylistItem> playlists = new Dictionary<string, PlaylistItem>();

                    if ((flags & App.PLAYLISTS) > 0)
                    {
                        foreach (TreeNode tn in GetAllNodes(playlistTree))
                        {
                            PlaylistItem pl = (PlaylistItem)tn.Tag;
                            if(!pl.Root && !pl.Hide)
                                playlists.Add(pl.Path, pl);
                        }
                    }

                    // Start the sync
                    _synchronizer.Start(App.LOGFILENAME, flags, Settings.SyncFileTypesRaw, Settings.ImportFileTypesRaw, playlists, customFolderText.Text.Trim(), _libraryInformation, Settings.Expressions);

                    // Set the button text to the correct context
                    goButton.Text = "Sto&p";
                    goButton.ImageIndex = 1;
                }
                else
                {
                    goButton.Enabled = true;
                    MessageBox.Show("You haven't selected anything to synchronize.", AboutBox.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else if(status == StatusEnum.Running)
            {
                // Tell the worker thread process to exit (nicely)
                _synchronizer.Running = StatusEnum.StopRequest;
            }
        }

        /// <summary>
        /// Gets the option flags
        /// </summary>
        /// <param name="all">If false, takes into account the child options of other options</param>
        /// <returns>The options</returns>
        private ulong GetOptionFlags(bool all)
        {
            // Get just the main form flags
            ulong options = Settings.Options & App.ABOUTBOXFLAGS;
            // Now add the main form flags
            options |=
                ((nameCheck.Checked) ? App.NAME : 0) |
                ((artistCheck.Checked) ? App.ARTIST : 0) |
                ((albumCheck.Checked) ? App.ALBUM : 0) |
                ((albumArtistCheck.Checked) ? App.ALBUMARTIST : 0) |
                ((composerCheck.Checked) ? App.COMPOSER : 0) |
                ((genreCheck.Checked) ? App.GENRE : 0) |
                ((commentCheck.Checked) ? App.COMMENT : 0) |
                ((groupingCheck.Checked) ? App.GROUPING : 0) |
                ((descriptionCheck.Checked) ? App.DESCRIPTION : 0) |
                ((discNumberCheck.Checked) ? App.DISCNUMBER : 0) |
                ((trackNumberCheck.Checked) ? App.TRACKNUMBER : 0) |
                ((yearCheck.Checked) ? App.YEAR : 0) |
                ((lyricsCheck.Checked) ? App.LYRICS : 0) |
                ((ratingCheck.Checked) ? App.RATING : 0) |
                ((lastPlayedCheck.Checked) ? App.LASTPLAYED : 0) |
                ((playCountCheck.Checked) ? App.PLAYCOUNT : 0) |
                ((skipCountCheck.Checked) ? App.SKIPCOUNT : 0) |
                ((lastSkippedCheck.Checked) ? App.LASTSKIPPED : 0) |
                ((compilationCheck.Checked) ? App.COMPILATION : 0) |
                ((bpmCheck.Checked) ? App.BPM : 0) |
                ((discCountCheck.Checked) ? App.DISCCOUNT : 0) |
                ((trackCountCheck.Checked) ? App.TRACKCOUNT : 0) |
                ((albumRatingCheck.Checked) ? App.ALBUMRATING : 0) |
                ((playlistCheck.Checked) ? App.PLAYLISTS : 0) |
                ((_showHiddenPlaylists && (all || playlistCheck.Checked)) ? App.SHOWHIDDENPLAYLISTS : 0) |
                ((sortNameCheck.Checked) ? App.SORTNAME : 0) |
                ((sortArtistCheck.Checked) ? App.SORTARTIST : 0) |
                ((sortAlbumArtistCheck.Checked) ? App.SORTALBUMARTIST : 0) |
                ((sortAlbumCheck.Checked) ? App.SORTALBUM : 0) |
                ((sortComposerCheck.Checked) ? App.SORTCOMPOSER : 0) |
                ((onTheGoCheck.Checked) ? App.ONTHEGOPLAYLISTS : 0) |
                ((rootFolderRadio.Checked && (all || playlistCheck.Checked)) ? App.USEPLAYLISTROOT : 0) |
                ((manageTicksCheck.Checked) ? App.MANAGETICKS : 0);

            return options;
        }

        /// <summary>
        /// Shows the About box
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox(_synchronizer.Running == StatusEnum.Running);
            // Show the About box as modal
            about.ShowDialog(this);
        }

        /// <summary>
        /// Opens up the log xml in the default program
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void openButton_Click(object sender, EventArgs e)
        {
            // Open up the log file
            try
            {
                System.Diagnostics.Process.Start(App.LOGFILENAME);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, AboutBox.AssemblyTitle);
            }
        }

        /// <summary>
        /// Enables or disables the playlist list box
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void playlistCheck_CheckedChanged(object sender, EventArgs e)
        {
            playlistTree.Enabled = playlistCheck.Checked;
            refreshTreeButton.Enabled = playlistCheck.Checked;
            rootFolderRadio.Enabled = playlistCheck.Checked;
            customFolderRadio.Enabled = playlistCheck.Checked;
            customFolderText.Enabled = (customFolderRadio.Checked && playlistCheck.Checked);
            checkAllTreeButton.Enabled = playlistCheck.Checked;
            uncheckAllTreeButton.Enabled = playlistCheck.Checked;
            toggleHiddenButton.Enabled = playlistCheck.Checked;
            manageTicksCheck.Enabled = playlistCheck.Checked;
        }

        /// <summary>
        /// Refreshes the tree
        /// </summary>
        /// <param name="playlists">List of Playlists</param>
        private void RefreshTree(Dictionary<string, PlaylistItem> playlists)
        {
            playlistTree.BeginUpdate();

            MCAutomation mc = null;

            try
            {
                // Get MC
                mc = Synchronizer.GetJRMediaCenter(_libraryInformation);

                // Fill the tree
                FillTree(mc, playlists);

                // Make sure of the checks
                foreach (TreeNode tn in GetAllNodes(playlistTree))
                {
                    if (tn.Parent != null && tn.Parent.Checked && !tn.Checked)
                        tn.Checked = true;
                }
            }
            catch (Exception)
            {
                progressLabel.Text = "An error occurred filling the playlist tree. Try loading MC manually first.";
                playlistTree.Nodes.Clear();
                EnableControls(false);
                goButton.Enabled = false;
                aboutToolStripMenuItem.Enabled = false;
            }
            finally
            {
                // We're finished with MC, clear the COM object
                if (mc != null)
                {
                    Marshal.ReleaseComObject(mc);
                    mc = null;
                }
            }

            playlistTree.EndUpdate();
        }

        /// <summary>
        /// Fills tree
        /// </summary>
        /// <param name="mc">The MC object</param>
        /// <param name="playlists">List of Playlists</param>
        private void FillTree(MCAutomation mc, Dictionary<string, PlaylistItem> playlists)
        {
            // Get the playlists
            Dictionary<string, IMJPlaylistAutomation> mcPlaylists = Synchronizer.GetMCPlaylists(mc, true);
            
            // Put the playlist names into a list and sort them
            List<string> sortedList = new List<string>(mcPlaylists.Keys);
            sortedList.Sort();

            // Sort into four distinct groups - we want to add them in this order
            List<PlaylistItem> playlistGroups = new List<PlaylistItem>();
            List<PlaylistItem> nonRootPlaylists = new List<PlaylistItem>();
            List<PlaylistItem> systemPlaylists = new List<PlaylistItem>();
            List<PlaylistItem> rootPlaylists = new List<PlaylistItem>();

            PlaylistItem pl = null;

            // Inspect each playlist
            foreach (string playlistName in sortedList)
            {
                // Get playlist type
                IMJPlaylistAutomation mcPlaylist = mcPlaylists[playlistName];

                // Don't add On-The-Go playlists, they are iTunes->MC only
                if (playlistName.Contains("On-The-Go"))
                    continue;

                pl = null;
                try { pl = playlists[playlistName]; }
                catch 
                { 
                    pl = new PlaylistItem();
                    pl.Path = playlistName;
                    playlists.Add(pl.Path, pl);
                }

                // Get the playlist type
                MC.PLAYLIST_TYPES type = (MC.PLAYLIST_TYPES)(-1);
                try { type = (MC.PLAYLIST_TYPES)int.Parse(mcPlaylist.Get("Type")); }
                catch {/* Do nothing */}

                // Get the ID
                switch (mcPlaylist.GetID())
                {
                    case MC.PLAYLIST_ID_SYSTEM_RECENTLY_IMPORTED:
                    case MC.PLAYLIST_ID_SYSTEM_RECENTLY_PLAYED:
                    case MC.PLAYLIST_ID_SYSTEM_RECENTLY_RIPPED:
                    case MC.PLAYLIST_ID_SYSTEM_TOP_HITS:
                        systemPlaylists.Add(pl);
                        break;
                    default:
                        {
                            switch (type)
                            {
                                case MC.PLAYLIST_TYPES.PLAYLIST_GROUP:
                                    {
                                        pl.Folder = true;
                                        playlistGroups.Add(pl);
                                        break;
                                    }
                                case MC.PLAYLIST_TYPES.SMARTLIST:
                                    {
                                        pl.Smart = true;
                                        goto case MC.PLAYLIST_TYPES.PLAYLIST;
                                    }
                                case MC.PLAYLIST_TYPES.PLAYLIST:
                                    {
                                        if (playlistName.IndexOf('\\') >= 0)
                                            nonRootPlaylists.Add(pl);
                                        else
                                            rootPlaylists.Add(pl);
                                        break;
                                    }
                            }
                            break;
                        }
                }
            }

            // Combine them all in the order we want
            List<PlaylistItem> allPlaylists = new List<PlaylistItem>();
            // Playlist groups first
            allPlaylists.AddRange(playlistGroups);
            // Now the non root playlists
            allPlaylists.AddRange(nonRootPlaylists);
            // Now the system playlists
            allPlaylists.AddRange(systemPlaylists);
            // Now the user playlists that exist in the root
            allPlaylists.AddRange(rootPlaylists);

            // Now do the updates to the tree
            // Clear the tree
            playlistTree.Nodes.Clear();

            // Create the root item
            TreeNode root = new TreeNode("All", 0, 0);
            root.Name = "\\*";

            pl = null;
            try { pl = playlists[root.Name]; }
            catch
            {
                pl = new PlaylistItem();
                pl.Path = root.Name;
                playlists.Add(pl.Path, pl);
            }

            pl.Root = true;

            root.Checked = pl.Selected;
            root.Tag = pl;

            // Add the root item
            playlistTree.Nodes.Add(root);
            // Clear the cache
            _playlistCache.Clear();
            // Add the root item
            _playlistCache.Add(pl.Path, pl);

            // Now create add them all to the tree
            foreach (PlaylistItem playlistItem in allPlaylists)
            {
                // Get parent attributes, this may be a new playlist
                // so treat it like so
                string parentPath;
                int loc = playlistItem.Path.LastIndexOf('\\');
                if (loc >= 0)
                    parentPath = playlistItem.Path.Remove(loc);
                else
                    parentPath = root.Name;

                PlaylistItem parent = playlists[parentPath];
                if (parent.Selected)
                    playlistItem.Selected = true;
                if (parent.Rebuild)
                    playlistItem.Rebuild = true;
                if (parent.Hide)
                    playlistItem.Hide = true;
                if (parent.Shuffle)
                    playlistItem.Shuffle = true;

                // Add to cache and tree
                _playlistCache[playlistItem.Path] = playlistItem;
                if(!playlistItem.Hide || _showHiddenPlaylists)
                    AddPlaylistToTree(playlistItem, mcPlaylists, root);
            }

            playlistTree.ExpandAll();
            playlistTree.TopNode = root;
        }

        /// <summary>
        /// Adds a playlist to the tree
        /// </summary>
        /// <param name="pl">The playlist to add</param>
        /// <param name="mcPlaylists">The collection of MC playlist objects</param>
        /// <param name="root">The root of the tree</param>
        private void AddPlaylistToTree(PlaylistItem pl, Dictionary<string, IMJPlaylistAutomation> mcPlaylists, TreeNode root)
        {
            // Get playlist type
            IMJPlaylistAutomation mcPlaylist = mcPlaylists[pl.Path];
            MC.PLAYLIST_TYPES type = (MC.PLAYLIST_TYPES)(-1);
            try { type = (MC.PLAYLIST_TYPES)int.Parse(mcPlaylist.Get("Type")); }
            catch {/* Do nothing */}

            // Get the image
            int imageIndex;
            switch (mcPlaylist.GetID())
            {
                case MC.PLAYLIST_ID_SYSTEM_RECENTLY_IMPORTED:
                case MC.PLAYLIST_ID_SYSTEM_RECENTLY_RIPPED:
                    imageIndex = 4;
                    break;
                case MC.PLAYLIST_ID_SYSTEM_RECENTLY_PLAYED:
                    imageIndex = 5;
                    break;
                case MC.PLAYLIST_ID_SYSTEM_TOP_HITS:
                    imageIndex = 2;
                    break;
                default:
                    {
                        switch (type)
                        {
                            case MC.PLAYLIST_TYPES.PLAYLIST_GROUP:
                                imageIndex = 1;
                                break;
                            case MC.PLAYLIST_TYPES.PLAYLIST:
                                imageIndex = 2;
                                break;
                            case MC.PLAYLIST_TYPES.SMARTLIST:
                                imageIndex = 3;
                                break;
                            default:
                                imageIndex = 0;
                                break;
                        }
                        break;
                    }
            }

            // Get the parent
            string[] parentNames = pl.Path.Split('\\');
            TreeNode parent = root;

            if (parentNames != null && (parentNames.Length > 1 || parentNames[0] != string.Empty))
            {
                foreach (string parentName in parentNames)
                {
                    bool found = false;
                    foreach (TreeNode child in parent.Nodes)
                    {
                        if (child.Text == parentName)
                        {
                            found = true;
                            parent = child;
                            break;
                        }
                    }
                    if (!found)
                    {
                        // Create the child item
                        TreeNode child = new TreeNode(parentName, imageIndex, imageIndex);
                        child.Name = pl.Path;
                        child.Checked = pl.Selected;
                        child.Tag = pl;
                        child.ForeColor = (pl.Hide) ? Color.Red:Color.Black;
                        // Add the child item
                        parent.Nodes.Add(child);
                        parent = child;
                    }
                }
            }
        }

        /// <summary>
        /// Gets all nodes from a treeview
        /// </summary>
        /// <param name="tn">The tree view</param>
        /// <returns>Returns a list containing all nodes in the tree</returns>
        private static List<TreeNode> GetAllNodes(TreeView tv)
        {
            List<TreeNode> tnlist = new List<TreeNode>();

            foreach (TreeNode child in tv.Nodes)
            {
                tnlist.Add(child);
                tnlist.AddRange(GetAllChildNodes(child));
            }

            return tnlist;
        }

        /// <summary>
        /// Gets all nodes under a specified parent tree node
        /// </summary>
        /// <param name="tn">The tree node to start at</param>
        /// <returns>Returns a list containing all nodes under the specified parent node in the tree</returns>
        private static List<TreeNode> GetAllChildNodes(TreeNode tn)
        {
            List<TreeNode> tnlist = new List<TreeNode>();

            foreach (TreeNode child in tn.Nodes)
            {
                tnlist.Add(child);
                tnlist.AddRange(GetAllChildNodes(child));
            }

            return tnlist;
        }

        /// <summary>
        /// Refreshes the tree
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void refreshTreeButton_Click(object sender, EventArgs e)
        {
            RefreshTree(Settings.Playlists);
        }

        /// <summary>
        /// Checks/unchecks according to parents
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void playlistTree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            PlaylistItem pl = (PlaylistItem)e.Node.Tag;
            pl.Selected = e.Node.Checked;

            if (e.Node.Checked)
            {
                foreach (TreeNode tn in e.Node.Nodes)
                {
                    tn.Checked = e.Node.Checked;
                    pl = (PlaylistItem)tn.Tag;
                    pl.Selected = true;
                }
            }
            else
            {
                if (e.Node.Parent != null && !e.Node.Checked && e.Node.Parent.Checked)
                {
                    e.Node.Parent.Checked = false;
                    pl = (PlaylistItem)e.Node.Parent.Tag;
                    pl.Selected = false;
                }
            }
        }

        /// <summary>
        /// Checks all items in the tree
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void checkAllTreeButton_Click(object sender, EventArgs e)
        {
            foreach (TreeNode tn in GetAllNodes(playlistTree))
            {
                tn.Checked = true;
            }
        }

        /// <summary>
        /// Unchecks all items in the tree
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void uncheckAllTreeButton_Click(object sender, EventArgs e)
        {
            foreach (TreeNode tn in GetAllNodes(playlistTree))
            {
                tn.Checked = false;
            }
        }

        /// <summary>
        /// Checks aggregate tag checks
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void checkAllAggregateButton_Click(object sender, EventArgs e)
        {
            CheckAllAggregateChecks(true);
        }

        /// <summary>
        /// Checks all export tag checks
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void checkAllExportButton_Click(object sender, EventArgs e)
        {
            CheckAllExportChecks(true);
        }

        /// <summary>
        /// Unchecks all export tag checks
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void uncheckAllAggregateButton_Click(object sender, EventArgs e)
        {
            CheckAllAggregateChecks(false);
        }

        /// <summary>
        /// Unchecks all export tag checks
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void uncheckAllExportButton_Click(object sender, EventArgs e)
        {
            CheckAllExportChecks(false);
        }

        /// <summary>
        /// Checks/unchecks aggregate checks
        /// </summary>
        /// <param name="check">Whether to check or uncheck</param>
        private void CheckAllAggregateChecks(bool check)
        {
            ratingCheck.Checked = check;
            lastPlayedCheck.Checked = check;
            playCountCheck.Checked = check;
            skipCountCheck.Checked = check;
            lastSkippedCheck.Checked = check;
        }

        /// <summary>
        /// Checks/unchecks export checks
        /// </summary>
        /// <param name="check">Whether to check or uncheck</param>
        private void CheckAllExportChecks(bool check)
        {
            nameCheck.Checked = check;
            artistCheck.Checked = check;
            albumCheck.Checked = check;
            albumArtistCheck.Checked = check;
            composerCheck.Checked = check;
            genreCheck.Checked = check;
            commentCheck.Checked = check;
            groupingCheck.Checked = check;
            descriptionCheck.Checked = check;
            discNumberCheck.Checked = check;
            trackNumberCheck.Checked = check;
            yearCheck.Checked = check;
            lyricsCheck.Checked = check;
            compilationCheck.Checked = check;
            bpmCheck.Checked = check;
            discCountCheck.Checked = check;
            trackCountCheck.Checked = check;
            sortNameCheck.Checked = check;
            sortArtistCheck.Checked = check;
            sortAlbumArtistCheck.Checked = check;
            sortAlbumCheck.Checked = check;
            sortComposerCheck.Checked = check;
        }

        /// <summary>
        /// The context menu is opening
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void expressionMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            bool enable = (_synchronizer.Running == StatusEnum.Idle);

            ContextMenuStrip menu = (ContextMenuStrip)sender;
            CheckBox checkBox = (CheckBox)menu.SourceControl;
            menu.Tag = checkBox;

            ToolStripMenuHeaderItem headerMenuItem = null;
            try { headerMenuItem = (ToolStripMenuHeaderItem)menu.Items["headerMenuItem"]; }
            catch { /* Do nothing */ }

            if (headerMenuItem == null)
            {
                headerMenuItem = new ToolStripMenuHeaderItem();
                headerMenuItem.BackColor = Color.LightCyan;
                headerMenuItem.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText;
                headerMenuItem.ForeColor = Color.Black;
                headerMenuItem.Name = "headerMenuItem";
                headerMenuItem.Size = new System.Drawing.Size(434, 22);
                menu.Items.Insert(0, headerMenuItem);
            }

            string expression = null;

            try
            {
                string defaultExpression = Settings.DefaultExpressions[checkBox.Text];
                expression = Settings.Expressions[checkBox.Text];
            }
            catch {/* Do nothing */}

            if (expression != null)
            {
                headerMenuItem.Text = string.Format("Please enter an MC expression for \"{0}\"", checkBox.Text);
                headerMenuItem.Image = Properties.Resources.Modify;

                menu.Items["customExpressionText"].Text = Settings.Expressions[checkBox.Text];
                menu.Items["customExpressionText"].Available = true;
                menu.Items["useDefaultMenuItem"].Available = true;
                menu.Items["saveMenuItem"].Available = true;
                menu.Items["customExpressionText"].Enabled = enable;
                menu.Items["useDefaultMenuItem"].Enabled = enable;
                menu.Items["saveMenuItem"].Enabled = enable;
                menu.Height = 121;
            }
            else
            {
                headerMenuItem.Text = string.Format("Sorry, expressions are not available for \"{0}\"", checkBox.Text);
                headerMenuItem.Image = Properties.Resources.Warning;

                menu.Items["customExpressionText"].Available = false;
                menu.Items["useDefaultMenuItem"].Available = false;
                menu.Items["saveMenuItem"].Available = false;
                menu.Items["customExpressionText"].Enabled = false;
                menu.Items["useDefaultMenuItem"].Enabled = false;
                menu.Items["saveMenuItem"].Enabled = false;
                menu.Height = 55;
            }
        }

        /// <summary>
        /// Reacts to a selection of a menu item
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void expressionMenuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ContextMenuStrip menu = (ContextMenuStrip)sender;
            CheckBox checkBox = (CheckBox)menu.Tag;

            switch (e.ClickedItem.Name)
            {
                case "useDefaultMenuItem":
                    {
                        menu.Items["customExpressionText"].Text = Settings.DefaultExpressions[checkBox.Text];
                        break;
                    }
                case "saveMenuItem"://Save setting and close the menu
                    {
                        ToolStripItem menuItem = menu.Items["customExpressionText"];
                        // Only save if this is a valid field to have expressions
                        if (menuItem.Enabled && Settings.DefaultExpressions.ContainsKey(checkBox.Text))
                        {
                            string newValue = menuItem.Text.Trim();
                            if (newValue != string.Empty)
                                Settings.Expressions[checkBox.Text] = newValue;
                        }
                        menu.Close();
                        break;
                    }
                case "cancelMenuItem":
                    {
                        menu.Close();
                        break;
                    }
            }
        }

        /// <summary>
        /// Stops the menu from closing when an item is clicked
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void expressionMenuStrip_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            if (e.CloseReason == ToolStripDropDownCloseReason.ItemClicked)
                e.Cancel = true;
        }

        /// <summary>
        /// Enables/Disables the custom folder text
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void customFolderRadio_CheckedChanged(object sender, EventArgs e)
        {
            customFolderText.Enabled = customFolderRadio.Checked;
            if (customFolderRadio.Checked)
            {
                customFolderText.Focus();
                customFolderText.SelectAll();
            }
        }

        /// <summary>
        /// Makes sure the clicked node is selected
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void playlistTree_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
                playlistTree.SelectedNode = e.Node;
        }

        /// <summary>
        /// Executes when the tree context menu is opening
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void treeContextMenu_Opening(object sender, CancelEventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
            {
                e.Cancel = true;
                return;
            }

            ToolStripMenuItem rebuildMenu = (ToolStripMenuItem)treeContextMenu.Items["rebuildMenuItem"];
            ToolStripMenuItem shuffleMenu = (ToolStripMenuItem)treeContextMenu.Items["shuffleMenuItem"];
            ToolStripMenuItem hideMenu = (ToolStripMenuItem)treeContextMenu.Items["hideMenuItem"];
            ToolStripMenuItem showMenu = (ToolStripMenuItem)treeContextMenu.Items["showMenuItem"];
            ToolStripMenuItem deselectMenu = (ToolStripMenuItem)treeContextMenu.Items["deselectAllMenuItem"];
            ToolStripMenuItem untickMenu = (ToolStripMenuItem)treeContextMenu.Items["untickMenuItem"];
            ToolStripMenuItem removeTracksMenu = (ToolStripMenuItem)treeContextMenu.Items["removeTracksMenuItem"];
            ToolStripItem sep1 = treeContextMenu.Items["toolStripSeparator1"];
            ToolStripItem sep2 = treeContextMenu.Items["toolStripSeparator2"];

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            rebuildMenu.Checked = pl.Rebuild;
            shuffleMenu.Checked = pl.Shuffle;
            untickMenu.Checked = !pl.Ticked;
            removeTracksMenu.Checked = pl.RemoveTracks;

            if (pl.Root || pl.Folder)
            {
                sep2.Enabled = true;
                sep2.Visible = true;
                deselectAllMenuItem.Enabled = true;
                deselectAllMenuItem.Visible = true;
                untickMenu.Enabled = false;
                untickMenu.Visible = false;
            }
            else
            {
                sep2.Enabled = false;
                sep2.Visible = false;
                deselectAllMenuItem.Enabled = false;
                deselectAllMenuItem.Visible = false;
                untickMenu.Enabled = manageTicksCheck.Checked;
                untickMenu.Visible = manageTicksCheck.Checked;
            }

            if(!pl.Root)
            {
                sep1.Enabled = true;
                sep1.Visible = true;

                if (pl.Hide)
                {
                    hideMenu.Enabled = false;
                    hideMenu.Visible = false;
                    showMenu.Enabled = true;
                    showMenu.Visible = true;
                }
                else
                {
                    hideMenu.Enabled = true;
                    hideMenu.Visible = true;
                    showMenu.Enabled = false;
                    showMenu.Visible = false;
                }
            }
            else
            {
                sep1.Enabled = false;
                sep1.Visible = false;
                hideMenu.Enabled = false;
                hideMenu.Visible = false;
                showMenu.Enabled = false;
                showMenu.Visible = false;
            }
        }

        /// <summary>
        /// Executes on selecting rebuild from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void rebuildMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Rebuild = !pl.Rebuild;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.Rebuild = pl.Rebuild;
            }

            if (!pl.Rebuild)
            {
                TreeNode parent = tn.Parent;
                while (parent != null)
                {
                    PlaylistItem ppl = (PlaylistItem)parent.Tag;
                    ppl.Rebuild = false;
                    parent = parent.Parent;
                }
            }
        }

        /// <summary>
        /// Executes on selecting shuffle from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void shuffleMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Shuffle = !pl.Shuffle;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.Shuffle = pl.Shuffle;
            }

            if (!pl.Shuffle)
            {
                TreeNode parent = tn.Parent;
                while (parent != null)
                {
                    PlaylistItem ppl = (PlaylistItem)parent.Tag;
                    ppl.Shuffle = false;
                    parent = parent.Parent;
                }
            }
        }

        /// <summary>
        /// Executes on selecting untick from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void untickMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Ticked = !pl.Ticked;
        }

        /// <summary>
        /// Executes on selecting remove tracks from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void removeTracksMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.RemoveTracks = !pl.RemoveTracks;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.RemoveTracks = pl.RemoveTracks;
            }

            if (!pl.RemoveTracks)
            {
                TreeNode parent = tn.Parent;
                while (parent != null)
                {
                    PlaylistItem ppl = (PlaylistItem)parent.Tag;
                    ppl.RemoveTracks = false;
                    parent = parent.Parent;
                }
            }
        }

        /// <summary>
        /// Executes on selecting hide from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void hideMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Hide = true;

            tn.ForeColor = (_showHiddenPlaylists) ? Color.Red : Color.Black;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.Hide = true;

                child.ForeColor = (_showHiddenPlaylists) ? Color.Red : Color.Black;
            }

            if (!_showHiddenPlaylists)
                tn.Remove();
        }

        /// <summary>
        /// Executes on selecting show from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void showMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Hide = false;

            tn.ForeColor = Color.Black;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.Hide = false;
                child.ForeColor = Color.Black;
            }

            TreeNode parent = tn.Parent;
            while (parent != null)
            {
                PlaylistItem ppl = (PlaylistItem)parent.Tag;
                ppl.Hide = false;
                parent.ForeColor = Color.Black;
                parent = parent.Parent;
            }
        }

        /// <summary>
        /// Executes on selecting deselect all from the tree context menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void deselectAllMenuItem_Click(object sender, EventArgs e)
        {
            TreeNode tn = playlistTree.SelectedNode;
            if (playlistTree.SelectedNode == null)
                return;

            tn.Checked = false;

            PlaylistItem pl = (PlaylistItem)tn.Tag;
            pl.Selected = false;

            foreach (TreeNode child in GetAllChildNodes(tn))
            {
                child.Checked = false;
                PlaylistItem cpl = (PlaylistItem)child.Tag;
                cpl.Selected = false;
            }
        }

        /// <summary>
        /// Toggle whether to show hidden playlists in the tree
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void toggleHiddenButton_Click(object sender, EventArgs e)
        {
            toggleHiddenButton.Refresh();
            _showHiddenPlaylists = !_showHiddenPlaylists;
            RefreshTree(_playlistCache);
        }

        /// <summary>
        /// Starts the sync when the form first starts
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void MCiTunesSynchronizerForm_Shown(object sender, EventArgs e)
        {
            // Start the sync if required
            if (_startSyncOnStartup)
            {
                _startSyncOnStartup = false;
                StartOrStopSync();
            }
        }
    }
}
