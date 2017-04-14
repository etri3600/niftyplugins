// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Linq;

namespace Aurora
{
	namespace NiftyPerforce
	{
		// Create a class to retrieve the OnBeforeSave event from VS
		// http://schmalls.com/2015/01/19/adventures-in-visual-studio-extension-development-part-2
		internal class RunningDocTableEvents : IVsRunningDocTableEvents3
		{
			private AutoCheckoutOnSave autoCheckoutOnSave;

			public RunningDocTableEvents(AutoCheckoutOnSave autoCheckoutOnSave)
			{
				this.autoCheckoutOnSave = autoCheckoutOnSave;
			}

			public int OnBeforeSave(uint docCookie)
			{
				RunningDocumentInfo runningDocumentInfo = autoCheckoutOnSave._rdt.Value.GetDocumentInfo(docCookie);
				autoCheckoutOnSave.OnBeforeSave(runningDocumentInfo.Moniker);
				return VSConstants.S_OK;
			}

			////////////////////////////////////////////////////////////////////
			// default implementation for the pure methods, return OK
			public int OnAfterAttributeChange(uint docCookie, uint grfAttribs){return VSConstants.S_OK;}
			public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew){return VSConstants.S_OK;}
			public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame){return VSConstants.S_OK;}
			public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining){return VSConstants.S_OK;}
			public int OnAfterSave(uint docCookie){return VSConstants.S_OK;}
			public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame){return VSConstants.S_OK;}
			public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining){return VSConstants.S_OK;}
		}

		class AutoCheckoutOnSave : PreCommandFeature
		{
			internal DTE _dte;
			internal readonly Lazy<RunningDocumentTable> _rdt;
			internal readonly Lazy<Microsoft.VisualStudio.OLE.Interop.IServiceProvider> _sp;

			public AutoCheckoutOnSave(Plugin plugin)
				: base(plugin, "AutoCheckoutOnSave", "Automatically checks out files on save")
			{
				if(!Singleton<Config>.Instance.autoCheckoutOnSave)
					return;

				Log.Info("Adding handlers for automatically checking out dirty files when you save");
				_dte = (DTE)Package.GetGlobalService(typeof(SDTE));
				_sp = new Lazy<Microsoft.VisualStudio.OLE.Interop.IServiceProvider>(() => Package.GetGlobalService(typeof(Microsoft.VisualStudio.OLE.Interop.IServiceProvider)) as Microsoft.VisualStudio.OLE.Interop.IServiceProvider);
				_rdt = new Lazy<RunningDocumentTable>(() => new RunningDocumentTable(new ServiceProvider(_sp.Value)));
				_rdt.Value.Advise(new RunningDocTableEvents(this));
			}

			internal void OnBeforeSave(string filename)
			{
				P4Operations.EditFileImmediate(mPlugin.OutputPane, filename);
			}

			private void OnSaveSelected(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
			{
				foreach(SelectedItem sel in mPlugin.App.SelectedItems)
				{
					if(sel.Project != null)
						P4Operations.EditFileImmediate(mPlugin.OutputPane, sel.Project.FullName);
					else if(sel.ProjectItem != null)
						P4Operations.EditFileImmediate(mPlugin.OutputPane, sel.ProjectItem.Document.FullName);
					else
						P4Operations.EditFileImmediate(mPlugin.OutputPane, mPlugin.App.Solution.FullName);
				}
			}

			private void OnSaveAll(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
			{
				if(!mPlugin.App.Solution.Saved)
					P4Operations.EditFileImmediate(mPlugin.OutputPane, mPlugin.App.Solution.FullName);

				foreach(Document doc in mPlugin.App.Documents)
				{
					if(doc.Saved)
						continue;
					P4Operations.EditFileImmediate(mPlugin.OutputPane, doc.FullName);
				}

				if(mPlugin.App.Solution.Projects == null)
					return;

				foreach(Project p in mPlugin.App.Solution.Projects)
				{
					EditProjectRecursive(p);
				}
			}

			private void EditProjectRecursive(Project p)
			{
				if(!p.Saved)
					P4Operations.EditFileImmediate(mPlugin.OutputPane, p.FullName);

				if(p.ProjectItems == null)
					return;

				foreach(ProjectItem pi in p.ProjectItems)
				{
					if(pi.SubProject != null)
					{
						EditProjectRecursive(pi.SubProject);
					}
					else if(!pi.Saved)
					{
						for(short i = 0; i <= pi.FileCount; i++)
							P4Operations.EditFileImmediate(mPlugin.OutputPane, pi.get_FileNames(i));
					}
				}
			}
		}
	}
}
