using System;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace Saturn_RemoveSuo
{
    [ProvideBindingPath]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [Guid("7A2F0672-7B43-4C35-94CD-1677B5548232")] //package id
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class SuoRemover : Package, IVsSolutionLoadEvents, IVsSolutionEvents
    {
        uint _cookie;
        private IVsSolution _solution;

        protected override void Initialize()
        {
            base.Initialize();

            _solution = (IVsSolution)GetService(typeof(SVsSolution));
            ErrorHandler.ThrowOnFailure(_solution.AdviseSolutionEvents(this, out _cookie));
            Remove_Suo();
        }


        /// <summary>
        /// Fired before a solution open begins. Extenders can activate a solution load manager by setting <see cref="F:Microsoft.VisualStudio.Shell.Interop.__VSPROPID4.VSPROPID_ActiveSolutionLoadManager" />.
        /// </summary>
        /// <param name="pszSolutionFilename">The name of the solution file.</param>
        /// <returns>
        /// If the method succeeds, it returns <see cref="F:Microsoft.VisualStudio.VSConstants.S_OK" />. If it fails, it returns an error code.
        /// </returns>
        public int OnBeforeOpenSolution(string pszSolutionFilename)
        {
            return Remove_Suo(pszSolutionFilename);
        }

        private int Remove_Suo(string pszSolutionFilename = null)
        {
            string solutionfileDirectory;
            DTE dte = (DTE)GetService(typeof(DTE));

            ActivityLog.LogWarning(pszSolutionFilename, "Solution File Name");

            if (!string.IsNullOrEmpty(pszSolutionFilename))
            {
                solutionfileDirectory = Path.GetDirectoryName(pszSolutionFilename);
            }
            else
            {
                solutionfileDirectory = Path.GetDirectoryName(dte.Solution.FullName);

                ActivityLog.LogWarning(solutionfileDirectory, "suo File Name from DTE");
            }

            if (solutionfileDirectory != null)
            {
                ;
                if (dte.Solution?.FileName != null)
                {
                    var suoPath = Path.Combine(solutionfileDirectory, @".vs", 
                        Path.GetFileNameWithoutExtension(dte.Solution.FileName), "v14");


                    if (suoPath != null)
                        foreach (var suoFile in Directory.EnumerateFiles(suoPath, "*.suo"))
                        {
                            try
                            {
                                ActivityLog.LogWarning(suoFile, "suo File Name");

                                File.Delete(suoFile);

                            }
                            catch (IOException)
                            {
                                // Never fail, no matter what.
                            }
                        }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
        {
            Remove_Suo();
            return VSConstants.S_OK;
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeBackgroundSolutionLoadBegins()
        {
            return VSConstants.S_OK;
        }

        public int OnAfterBackgroundSolutionLoadComplete()
        {
            // Load was successfull, clear the flag file.
            return VSConstants.S_OK;
        }



        #region Unused

        public int OnAfterCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            return VSConstants.S_OK;
        }

        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
        {
            return VSConstants.S_OK;
        }


        public int OnBeforeCloseSolution(object pUnkReserved)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeLoadProjectBatch(bool fIsBackgroundIdleBatch)
        {
            return VSConstants.S_OK;
        }

        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryBackgroundLoadProjectBatch(out bool pfShouldDelayLoadToNextIdle)
        {
            pfShouldDelayLoadToNextIdle = false;
            return VSConstants.S_OK;
        }

        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
        {
            return VSConstants.S_OK;
        }

        public int WriteUserOptsFile()
        {
            return VSConstants.S_OK;
        }

        public int IsBackgroundSolutionLoadEnabled(out bool pfIsEnabled)
        {
            pfIsEnabled = true;
            return VSConstants.S_OK;
        }

        public int EnsureProjectsAreLoaded(uint cProjects, Guid[] guidProjects, uint grfFlags)
        {
            return VSConstants.S_OK;
        }

        public int EnsureProjectIsLoaded(ref Guid guidProject, uint grfFlags)
        {
            return VSConstants.S_OK;
        }

        public int EnsureSolutionIsLoaded(uint grfFlags)
        {
            return VSConstants.S_OK;
        }

        public int ReloadProject(ref Guid guidProjectId)
        {
            return VSConstants.S_OK;
        }

        public int UnloadProject(ref Guid guidProjectID, uint dwUnloadStatus)
        {
            return VSConstants.S_OK;
        }

        #endregion
    }
}