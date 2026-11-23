using System;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>
    /// What the add-in's entry point tells the update system about this process.
    ///
    /// Kept out of ThisAddIn so the entry point stays one line per hook, and
    /// every hook here swallows its own failures: none of this is allowed to be
    /// the reason Excel fails to load DeepExcel.
    /// </summary>
    public static class UpdateHostHooks
    {
        private static readonly object Gate = new object();

        // Held for the life of the process on purpose and never disposed. The
        // add-in's DLL stays loaded until the process exits even after
        // OnDisconnection, so the files stay locked that long and the lease has
        // to say so. A static field also keeps the stream from being finalized.
        private static UpdateLease _lease;

        /// <summary>
        /// Called from OnConnection, before anything else can fail: every
        /// process that loaded the DLL holds its files, automation instances
        /// included.
        /// </summary>
        public static void OnAddInLoaded()
        {
            OnAddInLoaded(UpdateStage.DefaultRoot(), UpdateService.InstalledVersion);
        }

        internal static void OnAddInLoaded(string stageRoot, string version)
        {
            try
            {
                lock (Gate)
                {
                    if (_lease == null)
                    {
                        _lease = UpdateLease.Acquire(stageRoot, version);
                    }
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
