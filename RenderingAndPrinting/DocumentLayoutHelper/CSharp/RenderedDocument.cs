// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
using System;
using System.Collections.Generic;
using System.Text;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

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
        /// <para>Whenever document is updated and new page layout model is created, a new RenderedDocument instance must be used to access the changes.</para></remarks>
        public RenderedDocument(Document doc)
        {
            mLayoutCollector = new LayoutCollector(doc);
            mEnumerator = new LayoutEnumerator(doc);
            ProcessLayoutElements(this);
            CollectLinesAndAddToMarkers();
            LinkLayoutMarkersToNodes(doc);
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
                    if(row.Table == node)
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
                            if (span.Kind == "PARAGRAPH" || span.Kind == "ROW" || span.Kind == "CELL" || span.Kind == "SECTION")
                            {
                                mLayoutToLinesLookup.Add(span.LayoutObject, collectedLines);
                                collectedLines = new List<RenderedLine>();
                            }
                            else
                            {
                                mLayoutToSpanLookup.Add(span.LayoutObject, span);
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
                switch (node.NodeType)
                {
                    case NodeType.Paragraph:
                        foreach (RenderedLine line in GetLinesOfNode(node))
                            line.SetParentNode(node);
                        break;

                    case NodeType.Row:
                        foreach (RenderedLine line in GetLinesOfNode(node))
                            line.SetParentNode(((Row)node).LastCell.LastParagraph);
                        break;

                    default:
                        if (mLayoutCollector.GetEntity(node) != null)
                            mLayoutToSpanLookup[mLayoutCollector.GetEntity(node)].SetParentNode(node);
                        break;
                }
            }
        }

        private List<RenderedLine> GetLinesOfNode(Node node)
        {
            List<RenderedLine> lines = new List<RenderedLine>();
            object nodeEntity = mLayoutCollector.GetEntity(node);

            if (nodeEntity != null && mLayoutToLinesLookup.ContainsKey(nodeEntity))
                lines = mLayoutToLinesLookup[nodeEntity];

            return lines;
        }

        private LayoutCollector mLayoutCollector;
        private LayoutEnumerator mEnumerator;
        private static Dictionary<object, List<RenderedLine>> mLayoutToLinesLookup =
                new Dictionary<object, List<RenderedLine>>();
        private static Dictionary<object, RenderedSpan> mLayoutToSpanLookup =
                new Dictionary<object, RenderedSpan>();
    }
}
