using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// My About Box
    /// </summary>
    partial class AboutBox : Form
    {
        /// <summary>
        /// Whether the sync is currently running
        /// </summary>
        private bool _syncRunning;

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="syncRunning">Whether the synchronizer is currently running</param>
        public AboutBox(bool syncRunning)
        {
            InitializeComponent();

            // Set the member variable
            _syncRunning = syncRunning;

            // Set the window title to the version number and last release date
            this.Text = "MCiS " + AssemblyConfiguration;

            // Set the checkboxes and textboxes
            testModeCheck.Checked = (Settings.Options & App.TESTMODE) > 0;
            testModeCheck.Tag = "When checked, the synchronization will run but not make any changes to either of your libraries.\nThis gives you an opportunity to review the log and see what changes will be made when the synchronization is run for real.";
            alternativeFileKeyCheck.Checked = (Settings.Options & App.USEALTERNATIVEFILEKEY) > 0;
            alternativeFileKeyCheck.Tag = "When checked, will use the iTunesFileKey field in MC to find the corresponding file in iTunes, otherwise the filename is used.\nThis option is for those who wish to synchronize a file with a file held elsewhere on disk, for example synchronizing a FLAC file with an MP3 file.";

            podcastCheck.Checked = (Settings.Options & App.REFRESHPODCASTS) > 0;
            podcastCheck.Tag = "Tell iTunes to refresh your podcast feeds before the synchronization starts.";

            syncIpodCheck.Checked = (Settings.Options & App.SYNCHRONIZEIPODBEFORE) > 0;
            syncIpodCheck.Tag = "Tell iTunes to synchronize a connected iPod, iPhone or iPod Touch before the synchronization starts.";

            if ((Settings.Options & App.SYNCHRONIZEIPOD) > 0)
            {
                syncIpodRadio.Checked = true;
                quitiTunesRadio.Checked = false;
                doNothingRadio.Checked = false;
            }
            else if ((Settings.Options & App.QUITITUNES) > 0)
            {
                syncIpodRadio.Checked = false;
                quitiTunesRadio.Checked = true;
                doNothingRadio.Checked = false;
            }
            else
            {
                syncIpodRadio.Checked = false;
                quitiTunesRadio.Checked = false;
                doNothingRadio.Checked = true;
            }

            syncIpodRadio.Tag = "Tell iTunes to synchronize a connected iPod, iPhone or iPod Touch once the synchronization is finished.";
            quitiTunesRadio.Tag = "Tell iTunes to quit once the synchronization is finished.";
            doNothingRadio.Tag = "Once the synchronization is finished, do nothing.";

            forwardSlashesCheck.Tag = "Replaces forward slashes in your MC playlist names with a backslash in iTunes.";
            forwardSlashesCheck.Checked = (Settings.Options & App.REPLACEFORWARDSLASHES) > 0;
            App.SetToolTip(forwardSlashesCheck, true);

            podcastArtworkCheck.Checked = (Settings.Options & App.ADDARTWORKTOPODCASTS) > 0;
            podcastArtworkCheck.Tag = "If \"Folder.jpg\" exists in the same folder as an audio podcast within iTunes and the podcast has no artwork, attach the JPEG image to the podcast.";
            minimiseiTunesCheck.Checked = (Settings.Options & App.MINIMISEITUNES) > 0;
            minimiseiTunesCheck.Tag = "Minimize iTunes before the synchronization starts.";
            syncFileTypesCheck.Checked = (Settings.Options & App.SYNCCERTAINFILETYPES) > 0;
            syncFileTypesCheck.Tag = "Only synchronize files in the MC library of the specified types.\nThe text should be a semi-colon delimited list of file types.";
            syncFileTypesText.Text = Settings.SyncFileTypesRaw;
            importFilesCheck.Checked = (Settings.Options & App.IMPORTFILESTOITUNES) > 0;
            importFilesCheck.Tag = "Import files into iTunes of the specified types if they exist in the MC library but not in the iTunes library.\nThe text should be a semi-colon delimited list of file types.";
            importFileTypesText.Text = Settings.ImportFileTypesRaw;
            brokenLinksCheck.Checked = (Settings.Options & App.REMOVEBROKENLINKS) > 0;
            brokenLinksCheck.Tag = "Remove entries in the iTunes library where the physical file no longer exists.";

            logFileNotFoundCheck.Checked = (Settings.Options & App.LOGFILENOTFOUND) > 0;
            logFileNotFoundCheck.Tag = "Log when the file is in the iTunes library but not in the MC library.";
            logFileImportSuccessCheck.Checked = (Settings.Options & App.LOGIMPORTFILESUCCESS) > 0;
            logFileImportSuccessCheck.Tag = "Log when a file is successfully imported into the iTunes library.";
            logAlternativeKeyWarningCheck.Checked = (Settings.Options & App.LOGFILEKEYWARNING) > 0;
            logAlternativeKeyWarningCheck.Tag = "Log when Use Alternative File Keys is checked and the pathname entry in the iTunesFileKey field in MC is not found on disk.";
            logBrokenLinkCheck.Checked = (Settings.Options & App.LOGREMOVEBROKENLINK) > 0;
            logBrokenLinkCheck.Tag = "Log when a broken link is removed from the iTunes library.";
            logFileImportNoSuccess.Checked = (Settings.Options & App.LOGIMPORTFILENOSUCCESS) > 0;
            logFileImportNoSuccess.Tag = "Log when an import file into the iTunes library action fails.";
            logPlaylistCreateCheck.Checked = (Settings.Options & App.LOGPLAYLISTCREATED) > 0;
            logPlaylistCreateCheck.Tag = "Log when a playlist is created in iTunes or MC.";
            logPlaylistDeleteCheck.Checked = (Settings.Options & App.LOGPLAYLISTDELETED) > 0;
            logPlaylistDeleteCheck.Tag = "Log when a playlist is deleted in iTunes or MC.";
            logPlaylistFileAddedCheck.Checked = (Settings.Options & App.LOGPLAYLISTADDFILE) > 0;
            logPlaylistFileAddedCheck.Tag = "Log when a file is added to a playlist in iTunes or MC.";
            logPlaylistFileDeletedCheck.Checked = (Settings.Options & App.LOGPLAYLISTDELETEFILE) > 0;
            logPlaylistFileDeletedCheck.Tag = "Log when a file is deleted from a playlist in iTunes or MC.";
            logTrackChangesCheck.Checked = (Settings.Options & App.LOGTRACKCHANGES) > 0;
            logTrackChangesCheck.Tag = "Log individual track changes.";

            if ((Settings.Options & App.NOLOG) > 0)
            {
                noLogButton.Checked = true;
                logToClipboardButton.Checked = false;
                logToFileButton.Checked = false;
            }
            else if ((Settings.Options & App.LOGTOCLIPBOARD) > 0)
            {
                noLogButton.Checked = false;
                logToClipboardButton.Checked = true;
                logToFileButton.Checked = false;
            }
            else
            {
                noLogButton.Checked = false;
                logToClipboardButton.Checked = false;
                logToFileButton.Checked = true;
            }

            noLogButton.Tag = "MCiS will not log anything.";
            logToClipboardButton.Tag = "MCiS will log to memory and paste the results to the clipboard once the synchronization is finished.";
            logToFileButton.Tag = string.Format("MCiS will log to file \"{0}\" in the MCiS folder.", App.LOGFILENAME);

            // Set up the tooltips
            App.SetToolTip(testModeCheck, true);
            App.SetToolTip(alternativeFileKeyCheck, true);
            App.SetToolTip(podcastCheck, true);
            App.SetToolTip(syncIpodRadio, true);
            App.SetToolTip(quitiTunesRadio, true);
            App.SetToolTip(doNothingRadio, true);
            App.SetToolTip(podcastArtworkCheck, true);
            App.SetToolTip(minimiseiTunesCheck, true);
            App.SetToolTip(syncFileTypesCheck, true);
            App.SetToolTip(importFilesCheck, true);
            App.SetToolTip(brokenLinksCheck, true);
            App.SetToolTip(syncIpodCheck, true);
            App.SetToolTip(logFileNotFoundCheck, true);
            App.SetToolTip(logFileImportSuccessCheck, true);
            App.SetToolTip(logAlternativeKeyWarningCheck, true);
            App.SetToolTip(logBrokenLinkCheck, true);
            App.SetToolTip(logFileImportNoSuccess, true);
            App.SetToolTip(logPlaylistCreateCheck, true);
            App.SetToolTip(logPlaylistDeleteCheck, true);
            App.SetToolTip(logPlaylistFileAddedCheck, true);
            App.SetToolTip(logPlaylistFileDeletedCheck, true);
            App.SetToolTip(logTrackChangesCheck, true);
            App.SetToolTip(noLogButton, true);
            App.SetToolTip(logToClipboardButton, true);
            App.SetToolTip(logToFileButton, true);

            // Set menus
            SetTweakMenu();

            EnableControls(!syncRunning);
        }

        /// <summary>
        /// Enables/disables form controls
        /// </summary>
        /// <param name="enable">true to enable, false to disable</param>
        void EnableControls(bool enable)
        {
            testModeCheck.Enabled = enable;
            logFileNotFoundCheck.Enabled = enable;
            alternativeFileKeyCheck.Enabled = enable;
            podcastCheck.Enabled = enable;
            syncIpodCheck.Enabled = enable;
            syncIpodRadio.Enabled = enable;
            quitiTunesRadio.Enabled = enable;
            doNothingRadio.Enabled = enable;
            podcastArtworkCheck.Enabled = enable;
            minimiseiTunesCheck.Enabled = enable;
            syncFileTypesCheck.Enabled = enable;
            syncFileTypesText.Enabled = (enable && syncFileTypesCheck.Checked);
            importFilesCheck.Enabled = enable;
            importFileTypesText.Enabled = (enable && importFilesCheck.Checked);
            brokenLinksCheck.Enabled = enable;
            logFileNotFoundCheck.Enabled = enable;
            logFileImportSuccessCheck.Enabled = enable;
            logAlternativeKeyWarningCheck.Enabled = enable;
            logBrokenLinkCheck.Enabled = enable;
            logFileImportNoSuccess.Enabled = enable;
            logPlaylistCreateCheck.Enabled = enable;
            logPlaylistDeleteCheck.Enabled = enable;
            logPlaylistFileAddedCheck.Enabled = enable;
            logPlaylistFileDeletedCheck.Enabled = enable;
            logTrackChangesCheck.Enabled = enable;
            logToFileButton.Enabled = enable;
            noLogButton.Enabled = enable;
            logToClipboardButton.Enabled = enable;
            settingsGroup.Enabled = enable;
            beforeGroup.Enabled = enable;
            duringGroup.Enabled = enable;
            afterGroup.Enabled = enable;
            loggingGrp.Enabled = enable;
            tweakMenuItem.Enabled = enable;
        }

        #region Assembly Attribute Accessors

        public static string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public static string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public static string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public static string AssemblyConfiguration
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyConfigurationAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyConfigurationAttribute)attributes[0]).Configuration;
            }
        }

        public static string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public static string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public static string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion

        /// <summary>
        /// On checking sync file types
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void syncFileTypesCheck_CheckedChanged(object sender, EventArgs e)
        {
            syncFileTypesText.Enabled = syncFileTypesCheck.Checked;
            if (syncFileTypesCheck.Checked)
            {
                syncFileTypesText.Focus();
                syncFileTypesText.SelectAll();
            }
        }

        /// <summary>
        /// On checking import file types
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void importFilesCheck_CheckedChanged(object sender, EventArgs e)
        {
            importFileTypesText.Enabled = importFilesCheck.Checked;
            if (importFilesCheck.Checked)
            {
                importFileTypesText.Focus();
                importFileTypesText.SelectAll();
            }
        }

        /// <summary>
        /// On closing the form
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void AboutBox_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_syncRunning)
                return;

            // Clear about box settings first
            Settings.Options &= App.FIELDFLAGS;
            // Add the about box settings
            Settings.Options |= ((testModeCheck.Checked) ? App.TESTMODE : 0) |
                                ((alternativeFileKeyCheck.Checked) ? App.USEALTERNATIVEFILEKEY : 0) |
                                ((podcastCheck.Checked) ? App.REFRESHPODCASTS : 0) |
                                ((syncIpodRadio.Checked) ? App.SYNCHRONIZEIPOD : 0) |
                                ((syncIpodCheck.Checked) ? App.SYNCHRONIZEIPODBEFORE : 0) |
                                ((quitiTunesRadio.Checked) ? App.QUITITUNES : 0) |
                                ((podcastArtworkCheck.Checked) ? App.ADDARTWORKTOPODCASTS : 0) |
                                ((forwardSlashesCheck.Checked) ? App.REPLACEFORWARDSLASHES : 0) |
                                ((minimiseiTunesCheck.Checked) ? App.MINIMISEITUNES : 0) |
                                ((importFilesCheck.Checked) ? App.IMPORTFILESTOITUNES : 0) |
                                ((syncFileTypesCheck.Checked) ? App.SYNCCERTAINFILETYPES : 0) |
                                ((brokenLinksCheck.Checked) ? App.REMOVEBROKENLINKS : 0) |
                                ((logFileNotFoundCheck.Checked) ? App.LOGFILENOTFOUND : 0) |
                                ((logFileImportSuccessCheck.Checked) ? App.LOGIMPORTFILESUCCESS : 0) |
                                ((logAlternativeKeyWarningCheck.Checked) ? App.LOGFILEKEYWARNING : 0) |
                                ((logBrokenLinkCheck.Checked) ? App.LOGREMOVEBROKENLINK : 0) |
                                ((logFileImportNoSuccess.Checked) ? App.LOGIMPORTFILENOSUCCESS : 0) |
                                ((logPlaylistCreateCheck.Checked) ? App.LOGPLAYLISTCREATED : 0) |
                                ((logPlaylistDeleteCheck.Checked) ? App.LOGPLAYLISTDELETED : 0) |
                                ((logPlaylistFileAddedCheck.Checked) ? App.LOGPLAYLISTADDFILE : 0) |
                                ((logPlaylistFileDeletedCheck.Checked) ? App.LOGPLAYLISTDELETEFILE : 0) |
                                ((logTrackChangesCheck.Checked) ? App.LOGTRACKCHANGES : 0) |
                                ((noLogButton.Checked) ? App.NOLOG : 0) |
                                ((logToClipboardButton.Checked) ? App.LOGTOCLIPBOARD : 0);

            Settings.SyncFileTypesRaw = syncFileTypesText.Text;
            Settings.ImportFileTypesRaw = importFileTypesText.Text;
        }

        /// <summary>
        /// Executes on selecting an item from the tweak menu
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void tweakMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            int value = 0;

            if(item.Name.StartsWith("noneMenuItem"))
                value = 0;
            else if(item.Name.StartsWith("oneSecondMenuItem"))
                value = 1000;
            else if(item.Name.StartsWith("twoSecondsMenuItem"))
                value = 2000;
            else if(item.Name.StartsWith("fiveSecondsMenuItem"))
                value = 5000;
            else if(item.Name.StartsWith("tenSecondsMenuItem"))
                value = 10000;
            else if (item.Name.StartsWith("fifteenSecondsMenuItem"))
                value = 15000;
            else if (item.Name.StartsWith("twentySecondsMenuItem"))
                value = 20000;
            else if(item.Name.StartsWith("thirtySecondsMenuItem"))
                value = 30000;
            else if (item.Name.StartsWith("oneMinuteMenuItem"))
                value = 60000;
            else if (item.Name.StartsWith("twoMinutesMenuItem"))
                value = 120000;

            if(item.Name.EndsWith("1"))
                Settings.WaitTillSyncIPodBeforeSync = value;
            else if (item.Name.EndsWith("2"))
                Settings.WaitTillRefreshPodcastFeeds = value;
            else if (item.Name.EndsWith("3"))
                Settings.WaitAfterRefreshPodcastFeeds = value;
            else if (item.Name.EndsWith("4"))
                Settings.WaitBetweenPodcastRefreshIterations = value;
            else if (item.Name.EndsWith("5"))
                Settings.WaitOnSyncTrackError = value;
            else if (item.Name.EndsWith("6"))
                Settings.WaitOnLibraryLoadError = value;

            SetTweakMenu();
        }

        /// <summary>
        /// Sets the tweak menu up
        /// </summary>
        private void SetTweakMenu()
        {
            // Set the menu items
            noneMenuItem1.Checked = false;
            oneSecondMenuItem1.Checked = false;
            twoSecondsMenuItem1.Checked = false;
            fiveSecondsMenuItem1.Checked = false;
            tenSecondsMenuItem1.Checked = false;
            fifteenSecondsMenuItem1.Checked = false;
            twentySecondsMenuItem1.Checked = false;
            thirtySecondsMenuItem1.Checked = false;
            oneMinuteMenuItem1.Checked = false;
            twoMinutesMenuItem1.Checked = false;
            customMenuItem1.Checked = false;
            customTextBox1.Text = Settings.WaitTillSyncIPodBeforeSync.ToString();

            noneMenuItem2.Checked = false;
            oneSecondMenuItem2.Checked = false;
            twoSecondsMenuItem2.Checked = false;
            fiveSecondsMenuItem2.Checked = false;
            tenSecondsMenuItem2.Checked = false;
            fifteenSecondsMenuItem2.Checked = false;
            twentySecondsMenuItem2.Checked = false;
            thirtySecondsMenuItem2.Checked = false;
            oneMinuteMenuItem2.Checked = false;
            twoMinutesMenuItem2.Checked = false;
            customMenuItem2.Checked = false;
            customTextBox2.Text = Settings.WaitTillRefreshPodcastFeeds.ToString();

            noneMenuItem3.Checked = false;
            oneSecondMenuItem3.Checked = false;
            twoSecondsMenuItem3.Checked = false;
            fiveSecondsMenuItem3.Checked = false;
            tenSecondsMenuItem3.Checked = false;
            fifteenSecondsMenuItem3.Checked = false;
            twentySecondsMenuItem3.Checked = false;
            thirtySecondsMenuItem3.Checked = false;
            oneMinuteMenuItem3.Checked = false;
            twoMinutesMenuItem3.Checked = false;
            customMenuItem3.Checked = false;
            customTextBox3.Text = Settings.WaitAfterRefreshPodcastFeeds.ToString();

            noneMenuItem4.Checked = false;
            oneSecondMenuItem4.Checked = false;
            twoSecondsMenuItem4.Checked = false;
            fiveSecondsMenuItem4.Checked = false;
            tenSecondsMenuItem4.Checked = false;
            fifteenSecondsMenuItem4.Checked = false;
            twentySecondsMenuItem4.Checked = false;
            thirtySecondsMenuItem4.Checked = false;
            oneMinuteMenuItem4.Checked = false;
            twoMinutesMenuItem4.Checked = false;
            customMenuItem4.Checked = false;
            customTextBox4.Text = Settings.WaitBetweenPodcastRefreshIterations.ToString();

            noneMenuItem5.Checked = false;
            oneSecondMenuItem5.Checked = false;
            twoSecondsMenuItem5.Checked = false;
            fiveSecondsMenuItem5.Checked = false;
            tenSecondsMenuItem5.Checked = false;
            fifteenSecondsMenuItem5.Checked = false;
            twentySecondsMenuItem5.Checked = false;
            thirtySecondsMenuItem5.Checked = false;
            oneMinuteMenuItem5.Checked = false;
            twoMinutesMenuItem5.Checked = false;
            customMenuItem5.Checked = false;
            customTextBox5.Text = Settings.WaitOnSyncTrackError.ToString();

            noneMenuItem6.Checked = false;
            oneSecondMenuItem6.Checked = false;
            twoSecondsMenuItem6.Checked = false;
            fiveSecondsMenuItem6.Checked = false;
            tenSecondsMenuItem6.Checked = false;
            fifteenSecondsMenuItem6.Checked = false;
            twentySecondsMenuItem6.Checked = false;
            thirtySecondsMenuItem6.Checked = false;
            oneMinuteMenuItem6.Checked = false;
            twoMinutesMenuItem6.Checked = false;
            customMenuItem6.Checked = false;
            customTextBox6.Text = Settings.WaitOnLibraryLoadError.ToString();

            switch (Settings.WaitTillSyncIPodBeforeSync)
            {
                case 0:
                    noneMenuItem1.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem1.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem1.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem1.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem1.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem1.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem1.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem1.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem1.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem1.Checked = true;
                    break;
                default:
                    customMenuItem1.Checked = true;
                    break;
            }

            switch (Settings.WaitTillRefreshPodcastFeeds)
            {
                case 0:
                    noneMenuItem2.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem2.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem2.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem2.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem2.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem2.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem2.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem2.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem2.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem2.Checked = true;
                    break;
                default:
                    customMenuItem2.Checked = true;
                    break;
            }

            switch (Settings.WaitAfterRefreshPodcastFeeds)
            {
                case 0:
                    noneMenuItem3.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem3.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem3.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem3.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem3.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem3.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem3.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem3.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem3.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem3.Checked = true;
                    break;
                default:
                    customMenuItem3.Checked = true;
                    break;
            }

            switch (Settings.WaitBetweenPodcastRefreshIterations)
            {
                case 0:
                    noneMenuItem4.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem4.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem4.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem4.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem4.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem4.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem4.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem4.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem4.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem4.Checked = true;
                    break;
                default:
                    customMenuItem4.Checked = true;
                    break;
            }

            switch (Settings.WaitOnSyncTrackError)
            {
                case 0:
                    noneMenuItem5.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem5.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem5.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem5.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem5.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem5.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem5.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem5.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem5.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem5.Checked = true;
                    break;
                default:
                    customMenuItem5.Checked = true;
                    break;
            }

            switch (Settings.WaitOnLibraryLoadError)
            {
                case 0:
                    noneMenuItem6.Checked = true;
                    break;
                case 1000:
                    oneSecondMenuItem6.Checked = true;
                    break;
                case 2000:
                    twoSecondsMenuItem6.Checked = true;
                    break;
                case 5000:
                    fiveSecondsMenuItem6.Checked = true;
                    break;
                case 10000:
                    tenSecondsMenuItem6.Checked = true;
                    break;
                case 15000:
                    fifteenSecondsMenuItem6.Checked = true;
                    break;
                case 20000:
                    twentySecondsMenuItem6.Checked = true;
                    break;
                case 30000:
                    thirtySecondsMenuItem6.Checked = true;
                    break;
                case 60000:
                    oneMinuteMenuItem6.Checked = true;
                    break;
                case 120000:
                    twoMinutesMenuItem6.Checked = true;
                    break;
                default:
                    customMenuItem6.Checked = true;
                    break;
            }

            podcastRefreshIterationsBox.Text = Settings.PodcastRefreshIterations.ToString();
            syncTrackErrorRetriesBox.Text = Settings.SyncTrackErrorRetries.ToString();
            libraryLoadRetriesBox.Text = Settings.LibraryLoadRetries.ToString();
        }

        /// <summary>
        /// Executes on key down in a custom text box
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void customTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // Stops anything but numbers getting through
            switch (e.KeyValue)
            {
                case 8:
                case 13:
                    {
                        ToolStripTextBox item = (ToolStripTextBox)sender;
                        // Save setting
                        int value = 0;
                        try { value = int.Parse(item.Text); }
                        catch {/* Do nothing */}

                        item.Text = value.ToString();

                        if (item.Name == "podcastRefreshIterationsBox")
                            Settings.PodcastRefreshIterations = (value > 0) ? value : 1;
                        else if (item.Name == "syncTrackErrorRetriesBox")
                            Settings.SyncTrackErrorRetries = value;
                        else if (item.Name == "libraryLoadRetriesBox")
                            Settings.LibraryLoadRetries = value;
                        else
                        {
                            if (item.Name.EndsWith("1"))
                                Settings.WaitTillSyncIPodBeforeSync = value;
                            else if (item.Name.EndsWith("2"))
                                Settings.WaitTillRefreshPodcastFeeds = value;
                            else if (item.Name.EndsWith("3"))
                                Settings.WaitAfterRefreshPodcastFeeds = value;
                            else if (item.Name.EndsWith("4"))
                                Settings.WaitBetweenPodcastRefreshIterations = value;
                            else if (item.Name.EndsWith("5"))
                                Settings.WaitOnSyncTrackError = value;
                            else if (item.Name.EndsWith("6"))
                                Settings.WaitOnLibraryLoadError = value;
                        }

                        SetTweakMenu();
                        break;
                    }
                case 37:
                case 38:
                case 39:
                case 40:
                case 46:
                case 48:
                case 49:
                case 50:
                case 51:
                case 52:
                case 53:
                case 54:
                case 55:
                case 56:
                case 57:
                    break;
                default:
                    e.SuppressKeyPress = true;
                    break;
            }
        }

        /// <summary>
        /// Executes when the menu closes
        /// </summary>
        /// <param name="sender">The control that triggered the event</param>
        /// <param name="e">Event data</param>
        private void tweakMenuItem_DropDownClosed(object sender, EventArgs e)
        {
            int value;

            value = 1;
            try {value = int.Parse(podcastRefreshIterationsBox.Text);}
            catch {/* Do nothing */}

            Settings.PodcastRefreshIterations = (value > 0) ? value : 1;
            podcastRefreshIterationsBox.Text = Settings.PodcastRefreshIterations.ToString();

            value = 0;
            try {value = int.Parse(syncTrackErrorRetriesBox.Text);}
            catch {/* Do nothing */}

            Settings.SyncTrackErrorRetries = value;
            syncTrackErrorRetriesBox.Text = Settings.SyncTrackErrorRetries.ToString();

            value = 0;
            try {value = int.Parse(libraryLoadRetriesBox.Text);}
            catch {/* Do nothing */}

            Settings.LibraryLoadRetries = value;
            libraryLoadRetriesBox.Text = Settings.LibraryLoadRetries.ToString();
        }
    }
}
