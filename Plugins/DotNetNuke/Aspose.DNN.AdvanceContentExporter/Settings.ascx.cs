/*
' Copyright (c) 2015  Aspose.com
'  All rights reserved.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
' TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
' CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' 
*/

using System;
using DotNetNuke.Entities.Modules;
using DotNetNuke.Services.Exceptions;
using Aspose.Modules.AsposeDotNetNukeContentExport;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Aspose.Modules.DotNetNukeContentExport
{
    /// -----------------------------------------------------------------------------
    /// <summary>
    /// The Settings class manages Module Settings
    /// 
    /// Typically your settings control would be used to manage settings for your module.
    /// There are two types of settings, ModuleSettings, and TabModuleSettings.
    /// 
    /// ModuleSettings apply to all "copies" of a module on a site, no matter which page the module is on. 
    /// 
    /// TabModuleSettings apply only to the current module on the current page, if you copy that module to
    /// another page the settings are not transferred.
    /// 
    /// If you happen to save both TabModuleSettings and ModuleSettings, TabModuleSettings overrides ModuleSettings.
    /// 
    /// Below we have some examples of how to access these settings but you will need to uncomment to use.
    /// 
    /// Because the control inherits from DNNModule1SettingsBase you have access to any custom properties
    /// defined there, as well as properties from DNN such as PortalId, ModuleId, TabId, UserId and many more.
    /// </summary>
    /// -----------------------------------------------------------------------------
    public partial class Settings : AsposeDotNetNukeContentExportModuleSettingsBase
    {
        #region Base Method Implementations

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// LoadSettings loads the settings from the Database and displays them
        /// </summary>
        /// -----------------------------------------------------------------------------
        public override void LoadSettings()
        {
            try
            {
                if (Page.IsPostBack == false)
                {
                    //Check for existing settings and use those on this page
                    //Settings["SettingName"]

                    /* uncomment to load saved settings in the text boxes */

                    if (Settings.Contains("ExportTypeDropDownCssClass"))
                        ExportTypeDropDownCssClassTextBox.Text = Settings["ExportTypeDropDownCssClass"].ToString();

                    if (Settings.Contains("ExportButtonCssClass"))
                        ExportButtonCssClassTextBox.Text = Settings["ExportButtonCssClass"].ToString();

                    if (Settings.Contains("PaneSelectionDropDownCssClass"))
                        PaneSelectionDropDownCssClassTextBox.Text = Settings["PaneSelectionDropDownCssClass"].ToString();

                    if (Session["PanesDropDown_" + TabId.ToString()] != null)
                    {
                        ListItemCollection items = (ListItemCollection)Session["PanesDropDown_" + TabId.ToString()];

                        foreach (ListItem item in items)
                        {
                            PanesDropDownList.Items.Add(item);
                        }
                    }

                    if (Settings.Contains("DefaultPane"))
                        DefaultPaneTextBox.Text = Settings["DefaultPane"].ToString().Replace("dnn_", string.Empty);

                    if (PanesDropDownList.Items.Count <= 0)
                    {
                        DefaultPaneTextBox.Visible = true;
                        PanesDropDownList.Visible = false;
                    }

                    if (Settings.Contains("HideDefaultPane"))
                        HideDefaultPaneCheckBox.Checked = Convert.ToBoolean(Settings["HideDefaultPane"].ToString());
                }
            }
            catch (Exception exc) //Module failed to load
            {
                Exceptions.ProcessModuleLoadException(this, exc);
            }
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// UpdateSettings saves the modified settings to the Database
        /// </summary>
        /// -----------------------------------------------------------------------------
        public override void UpdateSettings()
        {
            try
            {
                var modules = new ModuleController();

                //the following are two sample Module Settings, using the text boxes that are commented out in the ASCX file.
                //module settings
                modules.UpdateModuleSetting(ModuleId, "ExportTypeDropDownCssClass", ExportTypeDropDownCssClassTextBox.Text);
                modules.UpdateModuleSetting(ModuleId, "ExportButtonCssClass", ExportButtonCssClassTextBox.Text);
                modules.UpdateModuleSetting(ModuleId, "PaneSelectionDropDownCssClass", PaneSelectionDropDownCssClassTextBox.Text);

                if (DefaultPaneTextBox.Visible)
                {
                    DefaultPaneTextBox.Text = Settings["DefaultPane"].ToString().StartsWith("dnn_") ? string.Empty : "dnn_" + DefaultPaneTextBox.Text.Trim();
                    modules.UpdateModuleSetting(ModuleId, "DefaultPane", DefaultPaneTextBox.Text);
                    modules.UpdateTabModuleSetting(TabModuleId, "DefaultPane", DefaultPaneTextBox.Text);
                }
                else
                {
                    modules.UpdateModuleSetting(ModuleId, "DefaultPane", PanesDropDownList.SelectedValue);
                    modules.UpdateTabModuleSetting(TabModuleId, "DefaultPane", PanesDropDownList.SelectedValue);
                }

                modules.UpdateModuleSetting(ModuleId, "HideDefaultPane", HideDefaultPaneCheckBox.Checked.ToString());

                //tab module settings
                modules.UpdateTabModuleSetting(TabModuleId, "ExportTypeDropDownCssClass", ExportTypeDropDownCssClassTextBox.Text);
                modules.UpdateTabModuleSetting(TabModuleId, "ExportButtonCssClass", ExportButtonCssClassTextBox.Text);
                modules.UpdateTabModuleSetting(TabModuleId, "PaneSelectionDropDownCssClass", PaneSelectionDropDownCssClassTextBox.Text);
                modules.UpdateTabModuleSetting(TabModuleId, "HideDefaultPane", HideDefaultPaneCheckBox.Checked.ToString());
            }
            catch (Exception exc) //Module failed to load
            {
                Exceptions.ProcessModuleLoadException(this, exc);
            }
        }

        #endregion
    }
}