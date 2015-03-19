using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.IO;
using EnvDTE80;
using EnvDTE;
using System.Linq;

namespace tanaka_733.OpenForPSCmdlet
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidOpenForPSCmdletPkgString)]
    public sealed class OpenForPSCmdletPackage : Package
    {
        private static DTE2 dte;

        internal static DTE2 DTE
        {
            get
            {
                if (dte == null)
                {
                    dte = ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE2;
                }

                return dte;
            }
        }

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public OpenForPSCmdletPackage()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidOpenForPSCmdletCmdSet, (int)PkgCmdIDList.cmdidOpenForPSCmdletCommand);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var proj = DTE.ActiveDocument.ProjectItem.ContainingProject;

            DTE.Solution.SolutionBuild.BuildProject(DTE.Solution.SolutionBuild.ActiveConfiguration.Name,
                proj.UniqueName, true);

            var outputPath = GetOutputPath(proj);
            if (!File.Exists(outputPath))
                return;
            var dir = Directory.GetParent(outputPath).ToString();
            try
            {
                var p = new System.Diagnostics.Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe",
                        WorkingDirectory = dir,
                        Arguments = $"-NoExit -Command \"& {{Import-Module '{outputPath}'}}\""
                    }
                };
                p.Start();

                var dteProcess = DTE.Debugger.LocalProcesses.Cast<EnvDTE.Process>().FirstOrDefault(prop => prop.ProcessID == p.Id);

                if (dteProcess != null)
                {
                    dteProcess.Attach();
                    DTE.Debugger.CurrentProcess = dteProcess;
                }
            }
            catch (Exception)
            {
                //TODO log
            }
            
            // Show a Message Box to prove we were here
            //IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            //Guid clsid = Guid.Empty;
            //int result;
            //Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
            //           0,
            //           ref clsid,
            //           "OpenForPSCmdlet",
            //           string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.ToString()),
            //           string.Empty,
            //           0,
            //           OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //           OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
            //           OLEMSGICON.OLEMSGICON_INFO,
            //           0,        // false
            //           out result));
        }

        static string GetOutputPath(Project proj)
        {
            var properties = proj.ConfigurationManager?.ActiveConfiguration?.Properties;
            if (properties == null)
            {
                throw new InvalidOperationException("project does not contain valid properties.");
            }
            var outputPath = GetProperty<string>(proj, "OutputPath");
            var assemblyName = GetProperty<string>(proj, "AssemblyName");
            //' The output folder can have these patterns:
            //' 1) "\\server\folder"
            //' 2) "drive:\folder"
            //' 3) "..\..\folder"
            //' 4) "folder"
            //see http://www.mztools.com/articles/2009/MZ2009015.aspx
            string absoluteOutputPath;
            if (outputPath.StartsWith(new string(Enumerable.Repeat(Path.DirectorySeparatorChar, 2).ToArray())))
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.Length >= 2 && outputPath[1] == Path.VolumeSeparatorChar)
            {
                absoluteOutputPath = outputPath;
            }
            else if (outputPath.Contains(@"..\"))
            {
                var projectFolder = Path.GetDirectoryName(proj.FullName);
                do
                {
                    outputPath = outputPath.Substring(3);
                    projectFolder = Path.GetDirectoryName(projectFolder);
                } while (outputPath.StartsWith(@"..\"));
                absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }
            else
            {
                var projectFolder = Path.GetDirectoryName(proj.FullName);
                absoluteOutputPath = Path.Combine(projectFolder, outputPath);
            }
            return Path.Combine(absoluteOutputPath, assemblyName + ".dll");
        }

        static T GetProperty<T>(Project project, string index) where T : class
        {
            try
            {
                var assemblyName = project.Properties.Item(index).Value as T;
                return assemblyName;
            }
            catch (Exception)
            {
                try
                {
                    var properties = project.ConfigurationManager?.ActiveConfiguration?.Properties;
                    return properties?.Item(index).Value as T;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }
    }
}
