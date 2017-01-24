
using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using System.Collections;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;

namespace DocumentExplorerExample
{
    /// <summary>
    /// Base class to provide GUI representation for document nodes.
    /// </summary>
    public class Item 
    {
        /// <summary>
        /// Creates Item for the document node.
        /// </summary>
        /// <param name="aNode">Document node which this Item will represent.</param>
        public Item(Node aNode)
        {
            mNode = aNode;
        }

        /// <summary>
        /// Document node that this Item represents.
        /// </summary>
        public Node Node
        {
            get { return mNode; }
        }

        /// <summary>
        ///  DisplayName for this Item. Can be customized by overriding in inheriting classes.
        /// </summary>
        public virtual string Name
        {
            get { return mNode.NodeType.ToString(); }
        }

        /// <summary>
        /// Text contained by the corresponding document node.
        /// </summary>
        public string Text
        {
            get
            {
                StringBuilder result = new StringBuilder();

                // All control characters are converted to human readable form.
                // E.g. [!PageBreak!], [!ParagraphBreak!], etc.
                foreach (char c in mNode.GetText())
                {
                    string controlCharDisplay = (string)gControlCharacters[c];
                    if (controlCharDisplay == null)
                        result.Append(c);
                    else 
                        result.Append(controlCharDisplay);
                }

                return result.ToString();
            }
        }

        /// <summary>
        /// Creates TreeNode for this item to be displayed in Document Explorer TreeView control.
        /// </summary>
        public TreeNode TreeNode
        {
            get 
            {
                if (mTreeNode == null)
                {
                    mTreeNode = new TreeNode(Name);
                    if (!gIconNames.Contains(IconName))
                    {
                        gIconNames.Add(IconName);
                        ImageList.Images.Add(Icon);
                    }
                    int index = gIconNames.IndexOf(IconName);
                    mTreeNode.ImageIndex = index;
                    mTreeNode.SelectedImageIndex = index;
                    mTreeNode.Tag = this;
                    if (mNode is CompositeNode && ((CompositeNode)mNode).ChildNodes.Count > 0)
                    {
                        mTreeNode.Nodes.Add("#dummy");
                    }
                }
                return mTreeNode;
            }
        }

        public static ImageList ImageList 
        {
            get 
            {
                if (mImageList == null) 
                {
                    mImageList = new ImageList();
                    mImageList.ColorDepth = ColorDepth.Depth32Bit;
                    mImageList.ImageSize = new Size(16, 16);
                }
                return mImageList;
            }
        }

        /// <summary>
        /// Icon to display in the Document Explorer TreeView control.
        /// </summary>
        public Icon Icon 
        {
            get 
            {
                if (mIcon == null) 
                {
                    mIcon = LoadIcon(IconName);
                    if (mIcon == null)
                        mIcon = LoadIcon("Node");
                }
                return mIcon;
            }
        }

        /// <summary>
        /// Icon for this node can be customized by overriding this property in the inheriting classes.
        /// The name represents name of .ico file without extension located in the Icons folder of the project.
        /// </summary>
        protected virtual string IconName 
        {
            get { return GetType().Name.Replace("Item", ""); }
        }

        /// <summary>
        /// Provides lazy on-expand loading of underlying tree nodes.
        /// </summary>
        public void OnExpand()
        {
            // Optimized to allow automatic conversion to VB.NET
            if (TreeNode.Nodes[0].Text.Equals("#dummy"))
            {
                TreeNode.Nodes.Clear();
                foreach (Node n in ((CompositeNode)mNode).ChildNodes)
                {
                    TreeNode.Nodes.Add(CreateItem(n).TreeNode);
                }
            }
        }
        
        /// <summary>
        /// Loads icon from assembly resource stream.
        /// </summary>
        /// <param name="anIconName">Name of the icon to load.</param>
        /// <returns>Icon object or null if icon was not found in the resources.</returns>
        private static Icon LoadIcon(string anIconName)
        {
            string resourceName = "Aspose.Words.Examples.CSharp.Viewers_Visualizers.Document_Explorer.Icons." + anIconName + ".ico";
            Stream iconStream = FetchResourceStream(resourceName);

            if (iconStream != null)
                return new Icon(iconStream);
            else
                return null;
        }

        /// <summary>
        /// Returns a resource stream from the executing assembly or throws if the resource cannot be found.
        /// </summary>
        /// <param name="resourceName">The name of the resource without the name of the assembly.</param>
        /// <returns>The stream. Don't forget to close it when finished.</returns>
        internal static Stream FetchResourceStream(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string fullName = string.Format("{0}Example.{1}", asm.GetName().Name, resourceName);
            Stream stream = asm.GetManifestResourceStream(fullName);

            // Ugly optimization so conversion to VB.NET can work.
            while (stream == null)
            {
                int dotPos = fullName.IndexOf(".");
                if (dotPos < 0)
                    return null;

                fullName = fullName.Substring(dotPos + 1);
                stream = asm.GetManifestResourceStream(fullName);
            }

            return stream;
        }

        public void Remove()
        {
            if (IsRemovable)
            {
                mNode.Remove();
                mTreeNode.Remove();
            }
        }

        public virtual bool IsRemovable
        {
            get { return true; }
        }

        /// <summary>
        /// Static ctor.
        /// </summary>
        static Item()
        {
            // Fill set of typenames of Item inheritors for Item class fabric.
            gItemSet = new ArrayList();
            foreach (Type type in Assembly.GetExecutingAssembly().GetTypes()) 
            {
                if (type.IsSubclassOf(typeof(Item)) && !type.IsAbstract) 
                {
                    gItemSet.Add(type.Name);
                }
            }

            // Fill control chars fields set
            gControlCharacters.Add(ControlChar.CellChar, "[!Cell!]");
            gControlCharacters.Add(ControlChar.ColumnBreakChar, "[!ColumnBreak!]\r\n");
            gControlCharacters.Add(ControlChar.FieldEndChar, "[!FieldEnd!]");
            gControlCharacters.Add(ControlChar.FieldSeparatorChar, "[!FieldSeparator!]");
            gControlCharacters.Add(ControlChar.FieldStartChar, "[!FieldStart!]");
            gControlCharacters.Add(ControlChar.LineBreakChar, "[!LineBreak!]\r\n");
            gControlCharacters.Add(ControlChar.LineFeedChar, "[!LineFeed!]");
            gControlCharacters.Add(ControlChar.NonBreakingHyphenChar, "[!NonBreakingHyphen!]");
            gControlCharacters.Add(ControlChar.NonBreakingSpaceChar, "[!NonBreakingSpace!]");
            gControlCharacters.Add(ControlChar.OptionalHyphenChar, "[!OptionalHyphen!]");
            gControlCharacters.Add(ControlChar.ParagraphBreakChar, "¶\r\n");
            gControlCharacters.Add(ControlChar.SectionBreakChar, "[!SectionBreak!]\r\n");
            gControlCharacters.Add(ControlChar.TabChar, "[!Tab!]");
        }

        /// <summary>
        /// Item class factory implementation.
        /// </summary>
        public static Item CreateItem(Node aNode)
        {
            string typeName = aNode.NodeType.ToString() + "Item";
            if (gItemSet.Contains(typeName))
                return (Item)Activator.CreateInstance(Type.GetType("DocumentExplorerExample." + typeName), new object[] {aNode});
            else
                return new Item(aNode);
        }

        private readonly Node mNode;
        private TreeNode mTreeNode;
        private static ImageList mImageList;
        private Icon mIcon;

        private static readonly ArrayList gItemSet;
        private static readonly ArrayList gIconNames = new ArrayList();
        /// <summary>
        /// Map of character to string that we use to display control MS Word control characters.
        /// </summary>
        private static readonly Hashtable gControlCharacters = new Hashtable();
    }
}