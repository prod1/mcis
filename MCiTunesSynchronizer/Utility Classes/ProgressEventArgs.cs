using System;

namespace MCiTunesSynchronizer
{
    /// <summary>
    /// Contains information about progress of the synchronization
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        /// <summary>
        /// The minimum value for the progress bar
        /// </summary>
        public int ProgressMin = -1;
        /// <summary>
        /// The maximum value for the progress bar
        /// </summary>
        public int ProgressMax = -1;
        /// <summary>
        /// The current progress value
        /// </summary>
        public int ProgressValue = -1;
        /// <summary>
        /// Any text information regarding current progress
        /// </summary>
        public string ProgressText = string.Empty;
        /// <summary>
        /// Whether the process has finished
        /// </summary>
        public bool Finished = false;
        /// <summary>
        /// The final status once finished 
        /// </summary>
        public StatusEnum FinalStatus = StatusEnum.Idle;
        /// <summary>
        /// The resulting XML
        /// </summary>
        public string ResultingXML = null;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressMin">The minimum value for the progress bar</param>
        /// <param name="progressMax">The maximum value for the progress bar</param>
        /// <param name="progressValue">The current progress value</param>
        public ProgressEventArgs(int progressMin, int progressMax, int progressValue)
        {
            ProgressMin = progressMin;
            ProgressMax = progressMax;
            ProgressValue = progressValue;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressMin">The minimum value for the progress bar</param>
        /// <param name="progressMax">The maximum value for the progress bar</param>
        /// <param name="progressValue">The current progress value</param>
        /// <param name="progressText">Any text information regarding current progress</param>
        public ProgressEventArgs(int progressMin, int progressMax, int progressValue, string progressText)
        {
            ProgressMin = progressMin;
            ProgressMax = progressMax;
            ProgressValue = progressValue;
            ProgressText = progressText;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressValue">The current progress value</param>
        public ProgressEventArgs(int progressValue)
        {
            ProgressValue = progressValue;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressValue">The current progress value</param>
        /// <param name="progressText">Any text information regarding current progress</param>
        public ProgressEventArgs(int progressValue, string progressText)
        {
            ProgressValue = progressValue;
            ProgressText = progressText;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressValue">The current progress value</param>
        /// <param name="progressText">Any text information regarding current progress</param>
        /// <param name="arg">A System.Object to format</param>
        public ProgressEventArgs(int progressValue, string progressText, params object[] arg)
        {
            ProgressValue = progressValue;
            ProgressText = string.Format(progressText, arg);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressText">Any text information regarding current progress</param>
        public ProgressEventArgs(string progressText)
        {
            ProgressText = progressText;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="progressText">Any text information regarding current progress</param>
        /// <param name="arg">A System.Object to format</param>
        public ProgressEventArgs(string progressText, params object[] arg)
        {
            ProgressText = string.Format(progressText, arg);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="finished">Whether the process has finished</param>
        /// <param name="finalStatus">The final status of the process</param>
        /// <param name="resultingXml">The resulting Xml</param>
        public ProgressEventArgs(bool finished, StatusEnum finalStatus, string resultingXml)
        {
            Finished = finished;
            FinalStatus = finalStatus;
            ResultingXML = resultingXml;
        }
    }
}
