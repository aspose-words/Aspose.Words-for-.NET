
using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class DocumentLayoutHelper
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            string fileName = "TestFile.docx";
            Document doc = new Document(dataDir + fileName);

            // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
            // the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.

            // Create a new RenderedDocument class from a Document object.
            RenderedDocument layoutDoc = new RenderedDocument(doc);

            // The following examples demonstrate how to use the wrapper API. 
            // This snippet returns the third line of the first page and prints the line of text to the console.
            RenderedLine line = layoutDoc.Pages[0].Columns[0].Lines[2];
            Console.WriteLine("Line: " + line.Text);

            // With a rendered line the original paragraph in the document object model can be returned.
            Paragraph para = line.Paragraph;
            Console.WriteLine("Paragraph text: " + para.Range.Text);

            // Retrieve all the text that appears of the first page in plain text format (including headers and footers).
            string pageText = layoutDoc.Pages[0].Text;
            Console.WriteLine();

            // Loop through each page in the document and print how many lines appear on each page.
            foreach (RenderedPage page in layoutDoc.Pages)
            {
                LayoutCollection<LayoutEntity> lines = page.GetChildEntities(LayoutEntityType.Line, true);
                Console.WriteLine("Page {0} has {1} lines.", page.PageIndex, lines.Count);
            }

            // This method provides a reverse lookup of layout entities for any given node (with the exception of runs and nodes in the
            // header and footer).
            Console.WriteLine();
            Console.WriteLine("The lines of the second paragraph:");
            foreach (RenderedLine paragraphLine in layoutDoc.GetLayoutEntitiesOfNode(doc.FirstSection.Body.Paragraphs[1]))
            {
                Console.WriteLine(string.Format("\"{0}\"", paragraphLine.Text.Trim()));
                Console.WriteLine(paragraphLine.Rectangle.ToString());
                Console.WriteLine();
            }

            Console.WriteLine("\nDocument layout helper example ran successfully.");
        }
    }

    /// <summary>
    /// Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in
    /// a object model like design.
    /// </summary>
    public class RenderedDocument : LayoutEntity
    {
        /// <summary>
        /// Creates a new instance from the supplied Document class.
        /// </summary>
        /// <param name="document">A document whose page layout model to enumerate.</param>
        /// <remarks><para>If page layout model of the document hasn't been built the enumerator calls <see cref="Document.UpdatePageLayout"/> to build it.</para>
        /// <para>Whenever document is updated and new page layout model is created, a new RenderedDocument instance must be used to access the changes.</para></remarks>
        public RenderedDocument(Document doc)
        {
            mLayoutCollector = new LayoutCollector(doc);
            mEnumerator = new LayoutEnumerator(doc);
            ProcessLayoutElements(this);
            LinkLayoutMarkersToNodes(doc);
            CollectLinesAndAddToMarkers();
        }

        /// <summary>
        /// Provides access to the pages of a document.
        /// </summary>
        public LayoutCollection<RenderedPage> Pages
        {
            get
            {
                return GetChildNodes<RenderedPage>();
            }
        }

        /// <summary>
        /// Returns all the layout entities of the specified node.
        /// </summary>
        /// <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
        public LayoutCollection<LayoutEntity> GetLayoutEntitiesOfNode(Node node)
        {
            if (!mLayoutCollector.Document.Equals(node.Document))
                throw new ArgumentException("Node does not belong to the same document which was rendered.");

            if (node.NodeType == NodeType.Document)
                return new LayoutCollection<LayoutEntity>(mChildEntities);

            List<LayoutEntity> entities = new List<LayoutEntity>();

            // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
            foreach (LayoutEntity entity in GetChildEntities(~LayoutEntityType.None, true))
            {
                if (entity.ParentNode == node)
                    entities.Add(entity);

                // There is no table entity in rendered output so manually check if rows belong to a table node.
                if (entity.Type == LayoutEntityType.Row)
                {
                    RenderedRow row = (RenderedRow)entity;
                    if (row.Table == node)
                        entities.Add(entity);
                }
            }

            return new LayoutCollection<LayoutEntity>(entities);
        }

        private void ProcessLayoutElements(LayoutEntity current)
        {
            do
            {
                LayoutEntity child = current.AddChildEntity(mEnumerator);

                if (mEnumerator.MoveFirstChild())
                {
                    current = child;

                    ProcessLayoutElements(current);
                    mEnumerator.MoveParent();

                    current = current.Parent;
                }
            } while (mEnumerator.MoveNext());
        }

        private void CollectLinesAndAddToMarkers()
        {
            CollectLinesOfMarkersCore(LayoutEntityType.Column);
            CollectLinesOfMarkersCore(LayoutEntityType.Comment);
        }

        private void CollectLinesOfMarkersCore(LayoutEntityType type)
        {
            List<RenderedLine> collectedLines = new List<RenderedLine>();

            foreach (RenderedPage page in Pages)
            {
                foreach (LayoutEntity story in page.GetChildEntities(type, false))
                {
                    foreach (RenderedLine line in story.GetChildEntities(LayoutEntityType.Line, true))
                    {
                        collectedLines.Add(line);
                        foreach (RenderedSpan span in line.Spans)
                        {
                            if (mLayoutToNodeLookup.ContainsKey(span.LayoutObject))
                            {
                                if (span.Kind == "PARAGRAPH" || span.Kind == "ROW" || span.Kind == "CELL" || span.Kind == "SECTION")
                                {
                                    Node node = mLayoutToNodeLookup[span.LayoutObject];

                                    if (node.NodeType == NodeType.Row)
                                        node = ((Row)node).LastCell.LastParagraph;

                                    foreach (RenderedLine collectedLine in collectedLines)
                                        collectedLine.SetParentNode(node);

                                    collectedLines = new List<RenderedLine>();
                                }
                                else
                                {
                                    span.SetParentNode(mLayoutToNodeLookup[span.LayoutObject]);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void LinkLayoutMarkersToNodes(Document doc)
        {
            foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
            {
                object entity = mLayoutCollector.GetEntity(node);

                if (entity != null)
                    mLayoutToNodeLookup.Add(entity, node);
            }
        }

        private LayoutCollector mLayoutCollector;
        private LayoutEnumerator mEnumerator;
        private static Dictionary<object, Node> mLayoutToNodeLookup =
                new Dictionary<object, Node>();
    }

    /// <summary>
    /// Provides the base class for rendered elements of a document.
    /// </summary>
    public abstract class LayoutEntity
    {
        protected LayoutEntity() { }

        /// <summary>
        /// Gets the 1-based index of a page which contains the rendered entity.
        /// </summary>
        public int PageIndex
        {
            get
            {
                return mPageIndex;
            }
        }

        /// <summary>
        /// Returns bounding rectangle of the entity relative to the page top left corner (in points).
        /// </summary>
        public RectangleF Rectangle
        {
            get
            {
                return mRectangle;
            }
        }

        /// <summary>
        /// Gets the type of this layout entity.
        /// </summary>
        public LayoutEntityType Type
        {
            get
            {
                return mType;
            }
        }

        /// <summary>
        /// Exports the contents of the entity into a string in plain text format.
        /// </summary>
        public virtual string Text
        {
            get
            {
                StringBuilder builder = new StringBuilder();
                foreach (LayoutEntity entity in mChildEntities)
                {
                    builder.Append(entity.Text);
                }

                return builder.ToString();
            }
        }

        /// <summary>
        /// Gets the immediate parent of this entity.
        /// </summary>
        public LayoutEntity Parent
        {
            get
            {
                return mParent;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        /// <remarks>This property may return null for spans that originate from Run nodes or nodes that are inside the header or footer.</remarks>
        public virtual Node ParentNode
        {
            get
            {
                return mParentNode;
            }
        }

        /// <summary>
        /// Internal method separate from ParentNode property to make code autoportable to VB.NET.
        /// </summary>
        internal virtual void SetParentNode(Node value)
        {
            mParentNode = value;
        }

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        internal object LayoutObject
        {
            get;
            set;
        }

        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        internal LayoutEntity AddChildEntity(LayoutEnumerator it)
        {
            LayoutEntity child = CreateLayoutEntityFromType(it);
            mChildEntities.Add(child);

            return child;
        }

        private LayoutEntity CreateLayoutEntityFromType(LayoutEnumerator it)
        {
            LayoutEntity childEntity;
            switch (it.Type)
            {
                case LayoutEntityType.Cell:
                    childEntity = new RenderedCell();
                    break;
                case LayoutEntityType.Column:
                    childEntity = new RenderedColumn();
                    break;
                case LayoutEntityType.Comment:
                    childEntity = new RenderedComment();
                    break;
                case LayoutEntityType.Endnote:
                    childEntity = new RenderedEndnote();
                    break;
                case LayoutEntityType.Footnote:
                    childEntity = new RenderedFootnote();
                    break;
                case LayoutEntityType.HeaderFooter:
                    childEntity = new RenderedHeaderFooter();
                    break;
                case LayoutEntityType.Line:
                    childEntity = new RenderedLine();
                    break;
                case LayoutEntityType.NoteSeparator:
                    childEntity = new RenderedNoteSeparator();
                    break;
                case LayoutEntityType.Page:
                    childEntity = new RenderedPage();
                    break;
                case LayoutEntityType.Row:
                    childEntity = new RenderedRow();
                    break;
                case LayoutEntityType.Span:
                    childEntity = new RenderedSpan(it.Text);
                    break;
                case LayoutEntityType.TextBox:
                    childEntity = new RenderedTextBox();
                    break;
                default:
                    throw new InvalidOperationException("Unknown layout type");
            }

            childEntity.mKind = it.Kind;
            childEntity.mPageIndex = it.PageIndex;
            childEntity.mRectangle = it.Rectangle;
            childEntity.mType = it.Type;
            childEntity.LayoutObject = it.Current;
            childEntity.mParent = this;

            return childEntity;
        }

        /// <summary>
        /// Returns a collection of child entities which match the specified type.
        /// </summary>
        /// <param name="type">Specifies the type of entities to select.</param>
        /// <param name="isDeep">True to select from all child entities recursively. False to select only among immediate children</param>
        public LayoutCollection<LayoutEntity> GetChildEntities(LayoutEntityType type, bool isDeep)
        {
            List<LayoutEntity> childList = new List<LayoutEntity>();

            foreach (LayoutEntity entity in mChildEntities)
            {
                if ((entity.Type & type) == entity.Type)
                    childList.Add(entity);

                if (isDeep)
                    childList.AddRange((IEnumerable<LayoutEntity>)entity.GetChildEntities(type, true));
            }

            return new LayoutCollection<LayoutEntity>(childList);
        }

        protected LayoutCollection<T> GetChildNodes<T>() where T : LayoutEntity, new()
        {
            T obj = new T();
            List<T> childList = new List<T>();

            foreach (LayoutEntity entity in mChildEntities)
            {
                if (entity.GetType() == obj.GetType())
                    childList.Add((T)entity);
            }

            return new LayoutCollection<T>(childList);
        }

        protected string mKind;
        protected int mPageIndex;
        protected Node mParentNode;
        protected RectangleF mRectangle;
        protected LayoutEntityType mType;
        protected LayoutEntity mParent;
        protected List<LayoutEntity> mChildEntities = new List<LayoutEntity>();
    }

    /// <summary>
    /// Represents a generic collection of layout entity types.
    /// </summary>
    public class LayoutCollection<T> : IEnumerable<T> where T : LayoutEntity
    {
        /// <summary>
        /// Reserved for internal use.
        /// </summary>
        internal LayoutCollection(List<T> baseList)
        {
            mBaseList = baseList;
        }

        /// <summary>
        /// Provides a simple "foreach" style iteration over the collection of nodes. 
        /// </summary>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return mBaseList.GetEnumerator();
        }

        /// <summary>
        /// Provides a simple "foreach" style iteration over the collection of nodes. 
        /// </summary>
        IEnumerator<T> IEnumerable<T>.GetEnumerator()
        {
            return mBaseList.GetEnumerator();
        }

        /// <summary>
        /// Returns the first entity in the collection.
        /// </summary>
        public T First
        {
            get
            {
                if (mBaseList.Count > 0)
                    return mBaseList[0];
                else
                    return default(T);
            }
        }

        /// <summary>
        /// Returns the last entity in the collection.
        /// </summary>
        public T Last
        {
            get
            {
                if (mBaseList.Count > 0)
                    return mBaseList[mBaseList.Count - 1];
                else
                    return default(T);
            }
        }

        /// <summary>
        /// Retrieves the entity at the given index. 
        /// </summary>
        /// <remarks><para>The index is zero-based.</para>
        /// <para>If index is greater than or equal to the number of items in the list, this returns a null reference.</para></remarks>
        public T this[int index]
        {
            get
            {
                if (index < mBaseList.Count)
                    return mBaseList[index];
                else
                    return default(T);
            }
        }

        /// <summary>
        /// Gets the number of entities in the collection.
        /// </summary>
        public int Count
        {
            get
            {
                return mBaseList.Count;
            }
        }

        private List<T> mBaseList;
    }

    /// <summary>
    /// Represents an entity that contains lines and rows.
    /// </summary>
    public abstract class StoryLayoutEntity : LayoutEntity
    {
        /// <summary>
        /// Provides access to the lines of a story.
        /// </summary>
        public LayoutCollection<RenderedLine> Lines { get { return GetChildNodes<RenderedLine>(); } }

        /// <summary>
        /// Provides access to the row entities of a table.
        /// </summary>
        public LayoutCollection<RenderedRow> Rows { get { return GetChildNodes<RenderedRow>(); } }
    }

    /// <summary>
    /// Represents line of characters of text and inline objects.
    /// </summary>
    public class RenderedLine : LayoutEntity
    {
        /// <summary>
        /// Exports the contents of the entity into a string in plain text format.
        /// </summary>
        public override string Text
        {
            get
            {
                return base.Text + Environment.NewLine;
            }
        }

        /// <summary>
        /// Returns the paragraph that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some lines such as those inside the header or footer.</remarks>
        public Paragraph Paragraph
        {
            get
            {
                return (Paragraph)ParentNode;
            }
        }

        /// <summary>
        /// Provides access to the spans of the line.
        /// </summary>
        public LayoutCollection<RenderedSpan> Spans { get { return GetChildNodes<RenderedSpan>(); } }
    }

    /// <summary>
    /// Represents one or more characters in a line.
    /// This include special characters like field start/end markers, bookmarks, shapes and comments.
    /// </summary>
    public class RenderedSpan : LayoutEntity
    {
        public RenderedSpan() { }

        internal RenderedSpan(string text)
        {
            // Assign empty text if the span text is null (this can happen with shape spans).
            mText = text != null ? text : string.Empty;
        }

        /// <summary>
        /// Gets kind of the span. This cannot be null.
        /// </summary>
        /// <remarks>This is a more specific type of the current entity, e.g. bookmark span has Span type and
        /// may have either a BOOKMARKSTART or BOOKMARKEND kind.</remarks>
        public string Kind
        {
            get
            {
                return mKind;
            }
        }

        /// <summary>
        /// Exports the contents of the entity into a string in plain text format.
        /// </summary>
        public override string Text
        {
            get
            {
                return mText;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        /// <remarks>This property returns null for spans that originate from Run nodes or nodes that are inside the header or footer.</remarks>
        public override Node ParentNode
        {
            get
            {
                return mParentNode;
            }
        }

        private string mText;
    }

    /// <summary>
    /// Represents the header/footer content on a page.
    /// </summary>
    public class RenderedHeaderFooter : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the type of the header or footer.
        /// </summary>
        public string Kind
        {
            get
            {
                return mKind;
            }
        }
    }

    /// <summary>
    /// Represents page of a document.
    /// </summary>
    public class RenderedPage : LayoutEntity
    {
        /// <summary>
        /// Provides access to the columns of the page.
        /// </summary>
        public LayoutCollection<RenderedColumn> Columns { get { return GetChildNodes<RenderedColumn>(); } }

        /// <summary>
        /// Provides access to the header and footers of the page.
        /// </summary>
        public LayoutCollection<RenderedHeaderFooter> HeaderFooters { get { return GetChildNodes<RenderedHeaderFooter>(); } }

        /// <summary>
        /// Provides access to the comments of the page.
        /// </summary>
        public LayoutCollection<RenderedComment> Comments { get { return GetChildNodes<RenderedComment>(); } }

        /// <summary>
        /// Returns the section that corresponds to the layout entity.  
        /// </summary>
        public Section Section
        {
            get
            {
                return (Section)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return Columns.First.GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Section);
            }
        }
    }

    /// <summary>
    /// Represents a table row.
    /// </summary>
    public class RenderedRow : LayoutEntity
    {
        /// <summary>
        /// Provides access to the cells of the row.
        /// </summary>
        public LayoutCollection<RenderedCell> Cells { get { return GetChildNodes<RenderedCell>(); } }

        /// <summary>
        /// Returns the row that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some rows such as those inside the header or footer.</remarks>
        public Row Row
        {
            get
            {
                return (Row)ParentNode;
            }
        }

        /// <summary>
        /// Returns the table that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some tables such as those inside the header or footer.</remarks>
        public Table Table
        {
            get
            {
                return Row != null ? Row.ParentTable : null;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
        public override Node ParentNode
        {
            get
            {
                Paragraph para = Cells.First.Lines.First != null ? Cells.First.Lines.First.Paragraph : null;
                return para != null ? para.GetAncestor(NodeType.Row) : para;
            }
        }
    }

    /// <summary>
    /// Represents a column of text on a page.
    /// </summary>
    public class RenderedColumn : StoryLayoutEntity
    {
        /// <summary>
        /// Provides access to the footnotes of the page.
        /// </summary>
        public LayoutCollection<RenderedFootnote> Footnotes { get { return GetChildNodes<RenderedFootnote>(); } }

        /// <summary>
        /// Provides access to the endnotes of the page.
        /// </summary>
        public LayoutCollection<RenderedEndnote> Endnotes { get { return GetChildNodes<RenderedEndnote>(); } }

        /// <summary>
        /// Provides access to the note separators of the page.
        /// </summary>
        public LayoutCollection<RenderedNoteSeparator> NoteSeparators { get { return GetChildNodes<RenderedNoteSeparator>(); } }

        /// <summary>
        /// Returns the body that corresponds to the layout entity.  
        /// </summary>
        public Body Body
        {
            get
            {
                return (Body)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Body);
            }
        }
    }

    /// <summary>
    /// Represents a table cell.
    /// </summary>
    public class RenderedCell : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the cell that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some cells such as those inside the header or footer.</remarks>
        public Cell Cell
        {
            get
            {
                return (Cell)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        /// <remarks>This property may return null for nodes that are inside the header or footer.</remarks>
        public override Node ParentNode
        {
            get
            {
                if (Lines.First == null)
                    return null;
                else
                    return Lines.First.Paragraph != null ? Lines.First.Paragraph.GetAncestor(NodeType.Cell) : null;
            }
        }
    }

    /// <summary>
    /// Represents placeholder for footnote content.
    /// </summary>
    public class RenderedFootnote : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the footnote that corresponds to the layout entity.  
        /// </summary>
        public Footnote Footnote
        {
            get
            {
                return (Footnote)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Footnote);
            }
        }
    }

    /// <summary>
    /// Represents placeholder for endnote content.
    /// </summary>
    public class RenderedEndnote : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the endnote that corresponds to the layout entity.  
        /// </summary>
        public Footnote Endnote
        {
            get
            {
                return (Footnote)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Footnote);
            }
        }
    }

    /// <summary>
    /// Represents text area inside of a shape.
    /// </summary>
    public class RenderedTextBox : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the Shape or DrawingML that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some Shapes or DrawingML such as those inside the header or footer.</remarks>
        public override Node ParentNode
        {
            get
            {
                LayoutCollection<LayoutEntity> lines = GetChildEntities(LayoutEntityType.Line, true);
                Node shape = lines.First.ParentNode.GetAncestor(NodeType.Shape);

                if (shape != null)
                    return shape;
                else
                    return lines.First.ParentNode.GetAncestor(NodeType.Shape);
            }
        }
    }

    /// <summary>
    /// Represents placeholder for comment content.
    /// </summary>
    public class RenderedComment : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the comment that corresponds to the layout entity.  
        /// </summary>
        public Comment Comment
        {
            get
            {
                return (Comment)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Comment);
            }
        }
    }

    /// <summary>
    /// Represents footnote/endnote separator.
    /// </summary>
    public class RenderedNoteSeparator : StoryLayoutEntity
    {
        /// <summary>
        /// Returns the footnote/endnote that corresponds to the layout entity.  
        /// </summary>
        public Footnote Footnote
        {
            get
            {
                return (Footnote)ParentNode;
            }
        }

        /// <summary>
        /// Returns the node that corresponds to this layout entity.  
        /// </summary>
        public override Node ParentNode
        {
            get
            {
                return GetChildEntities(LayoutEntityType.Line, true).First.ParentNode.GetAncestor(NodeType.Footnote);
            }
        }
    }
}
