using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

using Aspose.Words.Tables;
using Aspose.Words.Layout;

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
                if (entity.Type == type)
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
                return mBaseList[index];
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
    /// Represents an entity that has a reference to a document node.
    /// </summary>
    public abstract class NodeReferenceLayoutEntity : LayoutEntity
    {
        /// <summary>
        /// Returns the document paragraph that corresponds to the layout entity.  
        /// </summary>
        /// <remarks>This property may return null for some lines such as those inside the header or footer.</remarks>
        public Paragraph Paragraph { get; set; }
    }

    /// <summary>
    /// Represents line of characters of text and inline objects.
    /// </summary>
    public class RenderedLine : NodeReferenceLayoutEntity
    {
        public override string Text
        {
            get
            {
                return base.Text + Environment.NewLine;
            }
        }

        /// <summary>
        /// Provides access to the spans of the line.
        /// </summary>
        public LayoutCollection<RenderedSpan> Spans { get { return GetChildNodes<RenderedSpan>(); } }
    }

    /// <summary>
    /// Represents one or more characters in a line.
    /// This include special characters like field start/end markers, bookmarks and comments.
    /// </summary>
    public class RenderedSpan : LayoutEntity
    {
        public RenderedSpan() { }

        internal RenderedSpan(string text)
        {
            mText = text;
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
    }

    /// <summary>
    /// Represents a table cell.
    /// </summary>
    public class RenderedCell : StoryLayoutEntity { }

    /// <summary>
    /// Represents placeholder for footnote content.
    /// </summary>
    public class RenderedFootnote : StoryLayoutEntity { }

    /// <summary>
    /// Represents placeholder for endnote content.
    /// </summary>
    public class RenderedEndnote : StoryLayoutEntity { }

    /// <summary>
    /// Represents text area inside of a shape.
    /// </summary>
    public class RenderedTextBox : StoryLayoutEntity { }

    /// <summary>
    /// Represents placeholder for comment content.
    /// </summary>
    public class RenderedComment : StoryLayoutEntity { }

    /// <summary>
    /// Represents footnote/endnote separator.
    /// </summary>
    public class RenderedNoteSeparator : StoryLayoutEntity { }
}
