using umbraco.cms.presentation.Trees;
using umbraco.BusinessLogic.Actions;

namespace Aspose.UmbracoMemberExportToWord.Library
{
    public class MemberExportToWordTree : BaseTree
    {
        public MemberExportToWordTree(string application) : base(application) { }

        protected override void CreateRootNode(ref XmlTreeNode rootNode)
        {
            rootNode.NodeID = System.Guid.NewGuid().ToString();
            rootNode.Action = "javascript:openExportToWordPage()";
            rootNode.Menu.Clear();
            rootNode.Menu.Add(ActionRefresh.Instance);
            rootNode.Icon = "../../plugins/AsposeMemberExportToWord/Images/aspose.ico";
            rootNode.OpenIcon = "../../plugins/AsposeMemberExportToWord/Images/aspose.ico";
        }

        public override void Render(ref XmlTree tree)
        {
        }

        public override void RenderJS(ref System.Text.StringBuilder Javascript)
        {
            Javascript.Append(
               @"function openExportToWordPage() {
                 UmbClientMgr.contentFrame('/umbraco/plugins/AsposeMemberExportToWord/ExportToWord.aspx');
                }");
        }
    }
}