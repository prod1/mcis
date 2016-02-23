using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MCiTunesSynchronizer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool startSync = false, exitAfterSync = false, runMinimised = false, exportOnlyMCInfo = false, exportOnlyiTunesInfo = false, run = true;
            string libraryName = null, viewScheme = null, playList = null;

            // Go through the arguments
            if (args != null)
            {
                foreach (string arg in args)
                {
                    string comp = arg.ToUpper();

                    switch (comp)
                    {
                        case "/STARTSYNC":
                            startSync = true;
                            break;
                        case "/EXITAFTERSYNC":
                            exitAfterSync = true;
                            break;
                        case "/RUNMINIMIZED":
                            runMinimised = true;
                            break;
                        case "/EXPORTAGGREGATESONLYFROM:MC":
                            exportOnlyMCInfo = true;
                            break;
                        case "/EXPORTAGGREGATESONLYFROM:ITUNES":
                            exportOnlyiTunesInfo = true;
                            break;
                        case "/H":
                        case "/HELP":
                            {
                                run = false;
                                break;
                            }
                        default:
                            {
                                if (comp.StartsWith("/LIBRARYNAME:"))
                                {
                                    libraryName = arg.Substring(13);
                                    if (libraryName.Trim() == string.Empty)
                                        libraryName = null;
                                }
                                else if (comp.StartsWith("/VIEWSCHEME:"))
                                {
                                    viewScheme = arg.Substring(12);
                                    if (viewScheme.Trim() == string.Empty)
                                        viewScheme = null;
                                }
                                else if (comp.StartsWith("/PLAYLIST:"))
                                {
                                    playList = arg.Substring(10);
                                    if (playList.Trim() == string.Empty)
                                        playList = null;
                                }
                                else if (comp.StartsWith("/SETTINGS:"))
                                {
                                    string s = arg.Substring(10).Trim();
                                    if (s != string.Empty)
                                        App.SETTINGSFILENAME = s + ".settings";
                                }
                                else
                                {
                                    // Something invalid is in there, don't run the app
                                    run = false;
                                }
                                break;
                            }
                    }
                }
            }

            // Only allow the playlist OR the viewscheme to be specified
            if (playList != null && viewScheme != null)
                run = false;

            // Only allow export only to be specified on one app
            if (exportOnlyMCInfo && exportOnlyiTunesInfo)
                run = false;

            // If we're not running, display a message box indicating the valid arguments
            if (!run)
            {
                MessageBox.Show("Usage:\n\nMCiTunesSynchronizer.exe [/StartSync] [/ExitAfterSync] [/RunMinimized] [/ExportAggregatesOnlyFrom:[MC|iTunes]] [/LibraryName:\"sss\"] [[/ViewScheme:\"vvv\"] | [/Playlist:\"ppp\"]] [/Settings:\"xxx\"]\n\nWhere sss is the name of the library to use, vvv is the name of the viewscheme to limit the scope of the synchronize operation to, ppp is the name of the playlist to limit the scope of the synchronize operation to, and xxx is the settings file to use/create.", "MC & iTunes Synchronizer", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                return;
            }

            // Package up the given library information
            object[] libraryInfo = null;
            if(libraryName != null || viewScheme != null || playList != null)
                libraryInfo = new object[] { libraryName, viewScheme, playList };

            // Our main form from which everything happens
            MCiTunesSynchronizerForm mainForm = new MCiTunesSynchronizerForm(startSync, exitAfterSync, exportOnlyMCInfo, exportOnlyiTunesInfo, libraryInfo);
            if (runMinimised)
                mainForm.WindowState = FormWindowState.Minimized;

            // The user agreed or this is not the first run of the app, run the app!
            Application.Run(mainForm);
        }
    }
}
