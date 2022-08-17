using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AttachToStartupProject
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AttachToStartupProjectCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        public const int ToolbarCommandId = 0x0200;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("bc0f468b-9cc4-46fe-aedc-363f1cadd6d8");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachToStartupProjectCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private AttachToStartupProjectCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);

            var toolbarCommandID = new CommandID(CommandSet, ToolbarCommandId);
            var toolbarItem = new MenuCommand(this.Execute, toolbarCommandID);
            commandService.AddCommand(toolbarItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AttachToStartupProjectCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in AttachToStartupProjectCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new AttachToStartupProjectCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            if(await ServiceProvider.GetServiceAsync(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE)) is DTE2 dte)
            {
                var startupProjects = new List<object>((object[])dte.Solution.SolutionBuild.StartupProjects);
                foreach(Project project in dte.Solution.Projects)
                {
                    if(project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    {
                        var projectItems = GetSolutionFolderProjects(project);
                        foreach(Project projectItem in projectItems)
                        {
                            AttachToProcess(startupProjects, projectItem, dte);
                        }
                    }
                    else
                        AttachToProcess(startupProjects, project, dte);
                }
            }
        }

        private IEnumerable<Project> GetSolutionFolderProjects(Project project)
        {
            List<Project> projects = new List<Project>();
            var y = (project.ProjectItems as ProjectItems).Count;
            for(var i = 1; i <= y; i++)
            {
                var x = project.ProjectItems.Item(i).SubProject;
                var subProject = x as Project;
                if (subProject != null)
                {
                    projects.Add(subProject);
                }
            }

            return projects;
        }

        private static void AttachToProcess(List<object> startupProjects, Project project, DTE2 dte)
        {
            if(startupProjects.Find(p => ((string)p).StartsWith(project.Name)) != null)
            {
                foreach(Property property in project.Properties)
                {
                    if(property.Name == "AssemblyName")
                    {
                        foreach(Process process in dte.Debugger.LocalProcesses)
                        {
                            if(process.Name.Contains(property.Value.ToString()))
                            {
                                process.Attach();
                                break;
                            }
                        }

                        break;
                    }
                }
            }
        }
    }
}
