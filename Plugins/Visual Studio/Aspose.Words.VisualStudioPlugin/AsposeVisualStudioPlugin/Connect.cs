using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using AsposeVisualStudioPluginWords.GUI;
using AsposeVisualStudioPluginWords.Properties;

namespace AsposeVisualStudioPluginWords
{
    public class Connect : Extensibility.IDTExtensibility2, IDTCommandTarget
    {
        // Constants for command properties
        private const string ASPOSE_MENU_NAME = "MenuAsposeVSPlugin";
        private const string ASPOSE_MENU_CAPTION = "Aspose";
        private const string ASPOSE_MENU_TOOLTIP = "Download and run Aspose .NET API Examples";

        private const string WORDS_COMMAND_NAME = "AsposeWordsVSPlugin";
        private const string WORDS_COMMAND_CAPTION = "New Aspose.Words Example Project";
        private const string WORDS_COMMAND_TOOLTIP = "Download and run Aspose.Words for .NET API Examples";

        // Variables for IDE and add-in instances
        private DTE2 applicationObject;
        private EnvDTE.AddIn addInInstance;

        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode,
           object addInInst, ref System.Array custom)
        {
            try
            {
                applicationObject = (DTE2)application;
                addInInstance = (EnvDTE.AddIn)addInInst;

                switch (connectMode)
                {
                    case ext_ConnectMode.ext_cm_UISetup:

                        // Initialize the UI of the add-in
                        AddPermanentUI();
                        break;

                    case ext_ConnectMode.ext_cm_Startup:

                        // The add-in was marked to load on startup
                        // Do nothing at this point because the IDE may not be fully initialized
                        // Visual Studio will call OnStartupComplete when fully initialized
                        break;

                    case ext_ConnectMode.ext_cm_AfterStartup:

                        // The add-in was loaded by hand after startup using the Add-In Manager
                        // Initialize it in the same way that when is loaded on startup
                        InitializeAddIn();
                        break;
                }
            }
            catch (System.Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }
        }

        public void OnStartupComplete(ref System.Array custom)
        {
            InitializeAddIn();
        }

        private void InitializeAddIn()
        {
            // Initialize non-UI add-in
        }

        private void AddPermanentUI()
        {
            object[] contextGUIDS = new object[] { };
            Commands2 commands = (Commands2)applicationObject.Commands;

            Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar =
              ((Microsoft.VisualStudio.CommandBars.CommandBars)
              applicationObject.CommandBars)["MenuBar"];

            //Find the Tools command bar on the MenuBar command bar:
            CommandBarControl toolsControl = menuBarCommandBar.Controls["File"];
            CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;
            CommandBar oBar = null;
            CommandBarButton oBtn = null;
            //This try/catch block can be duplicated if you wish to add multiple commands
            //  to be handled by your Add-in, just make sure you also update the
            //  QueryStatus/Exec method to include the new command names.
            try
            {
                //User Code Start
                //searhing if submenu already exists
                for (int iloop = 1; iloop <= toolsPopup.CommandBar.Controls.Count; iloop++)
                {
                    if (toolsPopup.CommandBar.Controls[iloop].Caption == ASPOSE_MENU_CAPTION)
                    {
                        oBar = ((CommandBarPopup)toolsPopup.CommandBar.Controls[iloop]).CommandBar;
                        foreach (CommandBarButton cmdbtn in oBar.Controls)
                        {
                            if (cmdbtn.Caption == WORDS_COMMAND_CAPTION)
                            {
                                oBtn = cmdbtn;
                            }
                        }
                        break;
                    }
                }

                //if required submenu doesn't exist create a new one
                if (oBar == null)
                    oBar = (CommandBar)commands.AddCommandBar(ASPOSE_MENU_CAPTION,
                            vsCommandBarType.vsCommandBarTypeMenu, toolsPopup.CommandBar, 1);

                if (oBtn == null)
                {
                    //Add a command to the Commands collection:
                    Command myCommand = commands.AddNamedCommand2(addInInstance, WORDS_COMMAND_NAME,
                        WORDS_COMMAND_CAPTION, WORDS_COMMAND_TOOLTIP, false, Resources.pnglogosmall, ref contextGUIDS,
                        (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled,
                        (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

                    //Add a control for the command to the tools menu:
                    if ((myCommand != null) && (toolsPopup != null))
                        myCommand.AddControl(oBar, 1);

                }

                //User Code End
            }
            catch (System.ArgumentException)
            {
                // If we are here, then the exception is probably
                // because a command with that name already exists.
                // If so there is no need to recreate the command and we can 
                // safely ignore the exception.
            }
        }

        public void OnDisconnection(Extensibility.ext_DisconnectMode RemoveMode, ref System.Array custom)
        {
        }

        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        public void Exec(string cmdName, vsCommandExecOption executeOption, ref object varIn,
           ref object varOut, ref bool handled)
        {
            try
            {
                handled = false;

                if ((executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault))
                {
                    if (cmdName == addInInstance.ProgID + "." + WORDS_COMMAND_NAME)
                    {
                        SampleWizardPage page = new SampleWizardPage(applicationObject);
                        if (page != null && !page.IsDisposed) page.ShowDialog();
                        handled = true;
                        return;
                    }
                }
            }
            catch (Exception) { }

        }


        public void QueryStatus(string cmdName, vsCommandStatusTextWanted neededText,
           ref vsCommandStatus statusOption, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (cmdName == addInInstance.ProgID + "." + WORDS_COMMAND_NAME)
                {
                    statusOption = (vsCommandStatus)(vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported);
                    return;
                }
                else
                {
                    statusOption = vsCommandStatus.vsCommandStatusUnsupported;
                    return;
                }
            }
        }
    }
}
