//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

using Aspose.Words.Tables;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;

namespace Aspose.Words.Layout
{
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