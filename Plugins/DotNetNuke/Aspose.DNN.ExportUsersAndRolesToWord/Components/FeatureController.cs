/*
' Copyright (c) 2015 Aspose.com
'  All rights reserved.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
' TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
' CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
' 
*/

using System.Collections.Generic;
//using System.Xml;
using DotNetNuke.Entities.Modules;
using DotNetNuke.Services.Search;

namespace Aspose.DotNetNuke.Modules.Aspose.DNN.ExportUsersAndRolesToWord.Components
{

    /// -----------------------------------------------------------------------------
    /// <summary>
    /// The Controller class for Aspose.DNN.ExportUsersAndRolesToWord
    /// 
    /// The FeatureController class is defined as the BusinessController in the manifest file (.dnn)
    /// DotNetNuke will poll this class to find out which Interfaces the class implements. 
    /// 
    /// The IPortable interface is used to import/export content from a DNN module
    /// 
    /// The ISearchable interface is used by DNN to index the content of a module
    /// 
    /// The IUpgradeable interface allows module developers to execute code during the upgrade 
    /// process for a module.
    /// 
    /// Below you will find stubbed out implementations of each, uncomment and populate with your own data
    /// </summary>
    /// -----------------------------------------------------------------------------

    //uncomment the interfaces to add the support.
    public class FeatureController //: IPortable, ISearchable, IUpgradeable
    {


        #region Optional Interfaces

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// ExportModule implements the IPortable ExportModule Interface
        /// </summary>
        /// <param name="ModuleID">The Id of the module to be exported</param>
        /// -----------------------------------------------------------------------------
        //public string ExportModule(int ModuleID)
        //{
        //string strXML = "";

        //List<Aspose.DNN.ExportUsersAndRolesToWordInfo> colAspose.DNN.ExportUsersAndRolesToWords = GetAspose.DNN.ExportUsersAndRolesToWords(ModuleID);
        //if (colAspose.DNN.ExportUsersAndRolesToWords.Count != 0)
        //{
        //    strXML += "<Aspose.DNN.ExportUsersAndRolesToWords>";

        //    foreach (Aspose.DNN.ExportUsersAndRolesToWordInfo objAspose.DNN.ExportUsersAndRolesToWord in colAspose.DNN.ExportUsersAndRolesToWords)
        //    {
        //        strXML += "<Aspose.DNN.ExportUsersAndRolesToWord>";
        //        strXML += "<content>" + DotNetNuke.Common.Utilities.XmlUtils.XMLEncode(objAspose.DNN.ExportUsersAndRolesToWord.Content) + "</content>";
        //        strXML += "</Aspose.DNN.ExportUsersAndRolesToWord>";
        //    }
        //    strXML += "</Aspose.DNN.ExportUsersAndRolesToWords>";
        //}

        //return strXML;

        //	throw new System.NotImplementedException("The method or operation is not implemented.");
        //}

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// ImportModule implements the IPortable ImportModule Interface
        /// </summary>
        /// <param name="ModuleID">The Id of the module to be imported</param>
        /// <param name="Content">The content to be imported</param>
        /// <param name="Version">The version of the module to be imported</param>
        /// <param name="UserId">The Id of the user performing the import</param>
        /// -----------------------------------------------------------------------------
        //public void ImportModule(int ModuleID, string Content, string Version, int UserID)
        //{
        //XmlNode xmlAspose.DNN.ExportUsersAndRolesToWords = DotNetNuke.Common.Globals.GetContent(Content, "Aspose.DNN.ExportUsersAndRolesToWords");
        //foreach (XmlNode xmlAspose.DNN.ExportUsersAndRolesToWord in xmlAspose.DNN.ExportUsersAndRolesToWords.SelectNodes("Aspose.DNN.ExportUsersAndRolesToWord"))
        //{
        //    Aspose.DNN.ExportUsersAndRolesToWordInfo objAspose.DNN.ExportUsersAndRolesToWord = new Aspose.DNN.ExportUsersAndRolesToWordInfo();
        //    objAspose.DNN.ExportUsersAndRolesToWord.ModuleId = ModuleID;
        //    objAspose.DNN.ExportUsersAndRolesToWord.Content = xmlAspose.DNN.ExportUsersAndRolesToWord.SelectSingleNode("content").InnerText;
        //    objAspose.DNN.ExportUsersAndRolesToWord.CreatedByUser = UserID;
        //    AddAspose.DNN.ExportUsersAndRolesToWord(objAspose.DNN.ExportUsersAndRolesToWord);
        //}

        //	throw new System.NotImplementedException("The method or operation is not implemented.");
        //}

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// GetSearchItems implements the ISearchable Interface
        /// </summary>
        /// <param name="ModInfo">The ModuleInfo for the module to be Indexed</param>
        /// -----------------------------------------------------------------------------
        //public DotNetNuke.Services.Search.SearchItemInfoCollection GetSearchItems(DotNetNuke.Entities.Modules.ModuleInfo ModInfo)
        //{
        //SearchItemInfoCollection SearchItemCollection = new SearchItemInfoCollection();

        //List<Aspose.DNN.ExportUsersAndRolesToWordInfo> colAspose.DNN.ExportUsersAndRolesToWords = GetAspose.DNN.ExportUsersAndRolesToWords(ModInfo.ModuleID);

        //foreach (Aspose.DNN.ExportUsersAndRolesToWordInfo objAspose.DNN.ExportUsersAndRolesToWord in colAspose.DNN.ExportUsersAndRolesToWords)
        //{
        //    SearchItemInfo SearchItem = new SearchItemInfo(ModInfo.ModuleTitle, objAspose.DNN.ExportUsersAndRolesToWord.Content, objAspose.DNN.ExportUsersAndRolesToWord.CreatedByUser, objAspose.DNN.ExportUsersAndRolesToWord.CreatedDate, ModInfo.ModuleID, objAspose.DNN.ExportUsersAndRolesToWord.ItemId.ToString(), objAspose.DNN.ExportUsersAndRolesToWord.Content, "ItemId=" + objAspose.DNN.ExportUsersAndRolesToWord.ItemId.ToString());
        //    SearchItemCollection.Add(SearchItem);
        //}

        //return SearchItemCollection;

        //	throw new System.NotImplementedException("The method or operation is not implemented.");
        //}

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// UpgradeModule implements the IUpgradeable Interface
        /// </summary>
        /// <param name="Version">The current version of the module</param>
        /// -----------------------------------------------------------------------------
        //public string UpgradeModule(string Version)
        //{
        //	throw new System.NotImplementedException("The method or operation is not implemented.");
        //}

        #endregion

    }

}
