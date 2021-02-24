//------------------------------------------------------------------------------
// <copyright file="Command1Package.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.Linq;

namespace AttachToApppool
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(AttachCommandPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class AttachCommandPackage : Package
    {
        /// <summary>
        /// Command1Package GUID string.
        /// </summary>
        public const string PackageGuidString = "4dfb291e-230d-4cdc-9d27-caf7f4bdfa52";

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachCommand"/> class.
        /// </summary>
        public AttachCommandPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            AttachCommand.Initialize(this);
            base.Initialize();
            OleMenuCommandService mcs1 = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs1)
            {
                CommandID comboCommandId = new CommandID(GuidList.guidAttachApppoolCommandPackageCmdSet, (int)PkgCmdIDList.cmdidAttachToProcessComboList);
                OleMenuCommand menuMyDynamicComboCommand = new OleMenuCommand(new EventHandler(OnWorkerProcessFillItems), comboCommandId);
                mcs1.AddCommand(menuMyDynamicComboCommand);
            }

            OleMenuCommandService mcs2 = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs2)
            {
                CommandID comboCommandId = new CommandID(GuidList.guidAttachApppoolCommandPackageCmdSet, (int)PkgCmdIDList.cmdidAttachToProcessCombo);
                OleMenuCommand menuMyDynamicComboCommand = new OleMenuCommand(new EventHandler(OnWorkerProcessSelection), comboCommandId);
                mcs2.AddCommand(menuMyDynamicComboCommand);
            }

            OleMenuCommandService mcs3 = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs3)
            {
                CommandID buttonCommandId = new CommandID(GuidList.guidAttachApppoolCommandPackageCmdSet, (int)PkgCmdIDList.cmdidAttachDotnetButton);
                OleMenuCommand menuMyDynamicComboCommand = new OleMenuCommand(new EventHandler(AttachToDotNetProcesses), buttonCommandId);
                mcs2.AddCommand(menuMyDynamicComboCommand);
            }
        }

        #endregion

        private void OnWorkerProcessFillItems(object sender, EventArgs e)
        {
            var args = e as OleMenuCmdEventArgs;
            using (var serverManager = new Microsoft.Web.Administration.ServerManager())
            {
                var workerProcesses = serverManager.WorkerProcesses.Select(wp => wp.AppPoolName).OrderBy(pname => pname).ToArray();
                Marshal.GetNativeVariantForObject(workerProcesses, args.OutValue);
            }
        }

        private void OnWorkerProcessSelection(object sender, EventArgs e)
        {
            var args = e as OleMenuCmdEventArgs;
            if (args.InValue != null)
            {
                var selectedProcess = args.InValue.ToString();
                using (var serverManager = new Microsoft.Web.Administration.ServerManager())
                {
                    var workerProcess = serverManager.WorkerProcesses.FirstOrDefault(wp => wp.AppPoolName == selectedProcess);
                    if (workerProcess != null)
                    {
                        var dte2 = (EnvDTE80.DTE2)Package.GetGlobalService(typeof(SDTE));
                        var process = dte2.Debugger.LocalProcesses.Cast<EnvDTE.Process>().FirstOrDefault(p => p.ProcessID == workerProcess.ProcessId);
                        if (process != null)
                            process.Attach();
                    }
                }
            }
        }

        private void AttachToDotNetProcesses(object sender, EventArgs e)
        {
            var dotNetProcesses = Process.GetProcessesByName("dotnet");
            
            foreach (var dotNetProcess in dotNetProcesses)
            {
                var dte2 = (EnvDTE80.DTE2)Package.GetGlobalService(typeof(SDTE));
                var process = dte2.Debugger.LocalProcesses.Cast<EnvDTE.Process>().FirstOrDefault(p => p.ProcessID == dotNetProcess.Id);
                try
                {
                    if (process != null)
                        process.Attach();
                }
                catch
                {
                }
            }
        }

        static class GuidList
        {
            public const string guidAttachApppoolCommandPackageString = "4dfb291e-230d-4cdc-9d27-caf7f4bdfa52";
            public const string guidAttachApppoolCommandPackageCmdSetString = "3a6897c9-a077-477b-8864-a31cf0aaa5b0";

            public static readonly Guid guidAttachApppoolCommandPackageCmdSet = new Guid(guidAttachApppoolCommandPackageCmdSetString);
        };

        static class PkgCmdIDList
        {
            public const uint cmdidAttachToProcessCombo = 0x1051;
            public const uint cmdidAttachToProcessComboList = 0x1052;
            public const uint cmdidAttachDotnetButton = 0x1053;
        }
    }
}
