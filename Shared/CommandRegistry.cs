// Copyright (C) 2006-2010 Jim Tilander. See COPYING for and README for more details.
using System;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using Microsoft.VisualStudio.CommandBars;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;

namespace Aurora
{
	// Holds a dictionary between local command names and the instance that holds 
	// the logic to execute and update the command itself.
	public class CommandRegistry
	{
		private Dictionary<string, CommandBase> mCommands;
		private Dictionary<uint, CommandBase> mCommandsById;
		private Plugin mPlugin;
		private CommandBar mCommandBar;
		private Guid mPackageGuid;
		private Guid mCmdGroupGuid;

		public CommandRegistry(Plugin plugin, CommandBar commandBar, Guid packageGuid, Guid cmdGroupGuid)
		{
			mCommands = new Dictionary<string, CommandBase>();
			mCommandsById = new Dictionary<uint, CommandBase>();
			mPlugin = plugin;
			mCommandBar = commandBar;
			mPackageGuid = packageGuid;
			mCmdGroupGuid = cmdGroupGuid;
		}

		public void RegisterCommand(bool doBindings, CommandBase commandHandler)
		{
			RegisterCommand(doBindings, commandHandler, true);
		}

		public void RegisterCommand(bool doBindings, CommandBase commandHandler, bool onlyToolbar)
		{
			OleMenuCommand command = RegisterCommandPrivate(commandHandler, onlyToolbar);

			if(command != null && doBindings)
			{
				try
				{
                    Command cmd = mPlugin.Commands.Item(commandHandler.CanonicalName, -1);
					commandHandler.BindToKeyboard(cmd);
				}
				catch(ArgumentException e)
				{
					Log.Error("Failed to register keybindings for {0}: {1}", commandHandler.CanonicalName, e.ToString());
				}
			}

			mCommands.Add(commandHandler.CanonicalName, commandHandler);
		}
		
		private OleMenuCommand RegisterCommandPrivate(CommandBase commandHandler, bool toolbarOnly)
		{
			OleMenuCommand vscommand = null;
			uint cmdId = 0;
			try
			{
				Command existingCmd = mPlugin.Commands.Item(commandHandler.CanonicalName, -1);
                cmdId = (uint)existingCmd.ID;
			}
			catch(System.ArgumentException)
			{
			}

			if (cmdId == 0)
			{
				Log.Info("Registering the command {0} from scratch", commandHandler.Name);
				int result = mPlugin.ProfferCommands.AddNamedCommand(mPackageGuid, mCmdGroupGuid, commandHandler.CanonicalName, out cmdId, commandHandler.CanonicalName, commandHandler.CanonicalName, commandHandler.Tooltip, null, 0, (uint)commandHandler.IconIndex, 0, 0, null);
			}

			if (cmdId != 0)
			{
				OleMenuCommandService menuCommandService = mPlugin.MenuCommandService;
				CommandID commandID = new CommandID(mCmdGroupGuid, (int)cmdId);

				vscommand = new OleMenuCommand(OleMenuCommandCallback, commandID);
				vscommand.BeforeQueryStatus += this.OleMenuCommandBeforeQueryStatus;
				menuCommandService.AddCommand(vscommand);
				mCommandsById[cmdId] = commandHandler;
			}
			// Register the graphics controls for this command as well.
			// First let the command itself have a stab at register whatever it needs.
			// Then by default we always register ourselves in the main toolbar of the application.
			if(!commandHandler.RegisterGUI(vscommand, mCommandBar, toolbarOnly))
			{
			}

			return vscommand;
		}

		private void OleMenuCommandBeforeQueryStatus(object sender, EventArgs e)
		{

			try
			{
				OleMenuCommand oleMenuCommand = sender as OleMenuCommand;

				if (oleMenuCommand != null)
				{
					CommandID commandId = oleMenuCommand.CommandID;

					if (commandId != null)
					{
						if(mCommandsById.ContainsKey((uint)commandId.ID))
						{
							oleMenuCommand.Supported = true;
							oleMenuCommand.Enabled = mCommandsById[(uint)commandId.ID].IsEnabled();
							oleMenuCommand.Visible = true;
						}
					}
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine(ex.ToString());
			}
		}

		private void OleMenuCommandCallback(object sender, EventArgs e)
		{
			try
			{
				OleMenuCommand oleMenuCommand = sender as OleMenuCommand;

				if (oleMenuCommand != null)
				{
					CommandID commandId = oleMenuCommand.CommandID;
					if (commandId != null)
					{
						Log.Debug("Trying to execute command id \"{0}\"", commandId.ID);

						if(mCommandsById.ContainsKey((uint)commandId.ID))
						{
							bool dispatched = mCommandsById[(uint)commandId.ID].OnCommand();
							if (dispatched)
							{
								Log.Debug("{0} was dispatched", commandId.ID);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine(ex.ToString());
			}
		}
	}
}
