using System;
using System.Collections.Generic;
using System.Text;

using Aspose.Words;
using Aspose.Words.Layout;

namespace Aspose.Words.Layout
{
    /// <summary>
    /// Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in
    /// a object model like design.
    /// </summary>
    public class RenderedDocument : LayoutEntity
    {
        /// <summary>
        /// Creates a new instance from the supplied Aspose.Words.Document class.
        /// </summary>
        /// <param name="document">A document whose page layout model to enumerate.</param>
        /// <remarks><para>If page layout model of the document hasn't been built the enumerator calls <see cref="Document.UpdatePageLayout"/> to build it.</para>
        /// <para>Whenever document is updated and new page layout model is created, a new enumerator must be used to access it.</para></remarks>
        public RenderedDocument(Document doc)
        {
            mEnumerator = new LayoutEnumerator(doc);
            ProcessLayoutElements(this);
            CollectLinesAndAddToMarkers();
            LinkLinesToNodes(doc);
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
        /// Returns all lines of the specified document paragraph.
        /// </summary>
        /// <remarks>Note that this method sometimes won't return all lines if the paragraph is inside a table.</remarks>
        public LayoutCollection<RenderedLine> GetLinesOfParagraph(Paragraph para)
        {
            mEnumerator.MoveNode(para);
            List<RenderedLine> lines = new List<RenderedLine>();

            if (mLayoutToLinesLookup.ContainsKey(mEnumerator.Current))
                lines = mLayoutToLinesLookup[mEnumerator.Current];

            return new LayoutCollection<RenderedLine>(lines);
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
            CollectLinesOfMarkersCore(LayoutEntityType.HeaderFooter);
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
                            if (span.Kind == "PARAGRAPH" || span.Kind == "ROW" || span.Kind == "CELL" || span.Kind == "SECTION")
                            {
                                mLayoutToLinesLookup.Add(span.LayoutObject, collectedLines);
                                collectedLines = new List<RenderedLine>();
                            }
                        }
                    }
                }
            }
        }

        private void LinkLinesToNodes(Document doc)
        {
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                mEnumerator.MoveNode(para);

                if (mLayoutToLinesLookup.ContainsKey(mEnumerator.Current))
                {
                    foreach (LayoutEntity entity in mLayoutToLinesLookup[mEnumerator.Current])
                    {
                        ((NodeReferenceLayoutEntity)entity).Paragraph = para;
                    }
                }
            }
        }

        private LayoutEnumerator mEnumerator;
        private static Dictionary<object, List<RenderedLine>> mLayoutToLinesLookup =
                new Dictionary<object, List<RenderedLine>>();
    }
}
