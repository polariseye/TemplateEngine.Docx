using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
    internal class TableProcessor : IProcessor
    {
        private bool _isNeedToRemoveContentControls;
        private readonly ProcessContext _context;

        public TableProcessor(ProcessContext context)
        {
            _context = context;
        }

        public IProcessor SetRemoveContentControls(bool isNeedToRemove)
        {
            _isNeedToRemoveContentControls = isNeedToRemove;
            return this;
        }

        public ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items)
        {
            var processResult = ProcessResult.NotHandledResult;
            var handled = false;

            foreach (var contentItem in items)
            {
                var itemProcessResult = FillContent(contentControl, contentItem);
                if (!itemProcessResult.Handled) continue;

                handled = true;

                processResult.Merge(itemProcessResult);
            }

            if (!handled) return ProcessResult.NotHandledResult;

            if (processResult.Success && _isNeedToRemoveContentControls)
            {
                // Remove the content control for the table and replace it with its contents.
                foreach (var xElement in contentControl.AncestorsAndSelf(W.sdt))
                {
                    xElement.RemoveContentControl();
                }
            }

            return processResult;
        }

        /// <summary>
        /// Fills content with one content item
        /// </summary>
        /// <param name="contentControl">Content control</param>
        /// <param name="item">Content item</param>
        private ProcessResult FillContent(XContainer contentControl, IContentItem item)
        {
            if (!(item is TableContent))
                return ProcessResult.NotHandledResult;

            var processResult = ProcessResult.NotHandledResult;

            var table = item as TableContent;

            // If there isn't a table with that name, add an error to the error string,
            // and continue with next table.
            if (contentControl == null)
            {
                processResult.AddError(new ContentControlNotFoundError(table));

                return processResult;
            }

            // If the table doesn't contain content controls in cells, then error and continue with next table.
            var cellContentControl = contentControl
                .Descendants(W.sdt)
                .FirstOrDefault();
            if (cellContentControl == null)
            {
                processResult.AddError(new CustomContentItemError(table,
                    string.Format("doesn't contain content controls in cells")));

                return processResult;
            }

            if (table.IsHidden || table.Rows == null)
            {
                contentControl.Descendants(W.tbl).Remove();
            }
            else
            {
                var fieldNames = table.FieldNames.ToList();
                var prototypeRows = GetPrototype(contentControl, fieldNames);

                //Select content controls tag names
                var contentControlTagNames = prototypeRows
                    .Descendants(W.sdt)
                    .Select(sdt => sdt.SdtTagName())
                    .Where(fieldNames.Contains)
                    .ToList();

                //If there are not content controls with the one of specified field name we need to add the warning
                if (contentControlTagNames.Intersect(fieldNames).Count() != fieldNames.Count())
                {
                    // 如果fieldNames只与其中某行完全匹配，则只返回对应行的内容
                    var isPerfectMatch = true;
                    foreach (var prototypeRow in prototypeRows)
                    {
                        foreach (var fieldName in fieldNames)
                        {
                            if (!prototypeRow
                            .FirstLevelDescendantsAndSelf(W.sdt)
                            .Any(sdt => sdt.SdtTagName() == fieldName))
                            {
                                isPerfectMatch = false;
                                break;
                            }
                        }

                        if (isPerfectMatch)
                        {
                            prototypeRows.Clear();
                            prototypeRows.Add(prototypeRow);
                            break;
                        }
                    }
                    if (!isPerfectMatch)
                    {
                        var invalidFileNames = fieldNames
                            .Where(fn => !contentControlTagNames.Contains(fn))
                            .ToList();

                        processResult.AddError(
                            new CustomContentItemError(table,
                                string.Format("doesn't contain rows with cell content {0} {1}",
                                    invalidFileNames.Count > 1 ? "controls" : "control",
                                    string.Join(", ", invalidFileNames.Select(fn => string.Format("'{0}'", fn))))));
                    }

                }

                // Create a list of new rows to be inserted into the document.  Because this
                // is a document centric transform, this is written in a non-functional
                // style, using tree modification.
                var newRows = new List<List<XElement>>();
                foreach (var row in table.Rows)
                {
                    // Clone the prototypeRows into newRowsEntry.
                    var newRowsEntry = prototypeRows.Select(prototypeRow => new XElement(prototypeRow)).ToList();

                    // Create new rows that will contain the data that was passed in to this
                    // method in the XML tree.
                    foreach (var sdt in newRowsEntry.FirstLevelDescendantsAndSelf(W.sdt).ToList())
                    {
                        // Get fieldName from the content control tag.
                        var fieldName = sdt.SdtTagName();

                        var content = row.GetContentItem(fieldName);

                        if (content == null) continue;
                        var contentProcessResult = new ContentProcessor(_context)
                            .SetRemoveContentControls(_isNeedToRemoveContentControls)
                            .FillContent(sdt, content);

                        processResult.Merge(contentProcessResult);
                    }

                    // Add the newRow to the list of rows that will be placed in the newly
                    // generated table.
                    newRows.Add(newRowsEntry);
                }

                prototypeRows.Last().AddAfterSelf(newRows);

                // Remove the prototype rows
                prototypeRows.Remove();

                // merge cell
                if (table.MergeCellContent != null && table.MergeCellContent.Count > 0)
                {
                    // merge cell                    
                    this.MergeCell(contentControl, table);
                }
            }

            processResult.AddItemToHandled(table);

            return processResult;
        }

        // Determine the elements that contains the content controls with specified names.
        // This is the prototype for the rows that the code will generate from data.
        private List<XElement> GetPrototype(XContainer tableContentControl, IEnumerable<string> fieldNames)
        {
            var rowsWithContentControl = tableContentControl
                .Descendants(W.tr)
                .Where(tr =>
                    tr.Descendants(W.sdt)
                        .Any(sdt =>
                            {
                                var names = fieldNames as string[] ?? fieldNames.ToArray();
                                return !names.Any() || names.Contains(
                                           sdt.SdtTagName());
                            }))
                .ToList();


            return GetIntermediateAndMergedRows(rowsWithContentControl.First(), rowsWithContentControl.Last(),
                tableContentControl);
        }

        private List<XElement> GetIntermediateAndMergedRows(XElement firstRow, XElement lastRow, XContainer tableContentControl)
        {
            var resultRows = new List<XElement>();

            var mergeVector = new bool[lastRow.Descendants(W.tc).Count()];

            var firstRowReached = false;
            var lastRowReached = false;

            //find merged rows and rows between first and last rows
            foreach (var tableRow in tableContentControl.Descendants(W.tr))
            {
                if (tableRow == firstRow)
                {
                    resultRows.Add(tableRow);
                    firstRowReached = true;
                }
                if (!firstRowReached) continue;

                if (!lastRowReached)
                {
                    if (tableRow == lastRow)
                    {
                        if (firstRow != lastRow)
                            resultRows.Add(tableRow);

                        var lastRowCells = lastRow.Descendants(W.tc).ToArray();
                        for (var i = 0; i < lastRowCells.Count(); i++)
                        {
                            var cell = lastRowCells[i];
                            var cellFormatting = cell.Element(W.tcPr);
                            if (cellFormatting != null && cellFormatting.Element(W.vMerge) != null)
                            {
                                mergeVector[i] = true;
                            }
                        }
                        lastRowReached = true;
                        continue;
                    }

                    if (tableRow != firstRow)
                        resultRows.Add(tableRow);
                }

                //if there are any maybe merged rows
                if (mergeVector.Any(r => r))
                {
                    var rowCells = tableRow.Descendants(W.tc).ToArray();
                    for (var i = 0; i < rowCells.Count(); i++)
                    {
                        var cell = rowCells[i];
                        var cellFormatting = cell.Element(W.tcPr);
                        if (cellFormatting != null && cellFormatting.Element(W.vMerge) != null &&
                            (cellFormatting.Element(W.vMerge).Attribute(W.val) == null ||
                             cellFormatting.Element(W.vMerge).Attribute(W.val).Value == "continue"))
                        {
                            resultRows.Add(tableRow);
                            mergeVector[i] = true;
                        }
                        else
                        {
                            mergeVector[i] = false;
                        }
                    }
                }
                else if (lastRowReached)
                    break;
            }


            return resultRows;
        }

        private void MergeCell(XContainer contentControl, TableContent table)
        {
            var tableElement = contentControl.Descendants(W.tbl).FirstOrDefault();
            if (tableElement == null)
            {
                return;
            }

            var rows = tableElement.Descendants(W.tr).ToList();
            var rowCells = new List<List<XElement>>();
            foreach (var item in rows)
            {
                rowCells.Add(item.Descendants(W.tc).ToList());
            }

            foreach (var item in table.MergeCellContent)
            {
                if (item.StartRowIndex == item.EndRowIndex)
                {
                    continue;
                }

                for (var columnIndex = item.StartColumnIndex; columnIndex <= item.EndColumnIndex; columnIndex++)
                {
                    var cellItem = rowCells[item.StartRowIndex][columnIndex];
                    this.SetMerge(cellItem, W.vMerge, "restart");

                    for (var rowIndex = item.StartRowIndex + 1; rowIndex <= item.EndRowIndex; rowIndex++)
                    {
                        cellItem = rowCells[rowIndex][columnIndex];
                        this.SetMerge(cellItem, W.vMerge, "");
                    }
                }
            }

            foreach (var item in table.MergeCellContent)
            {
                if (item.StartColumnIndex == item.EndColumnIndex)
                {
                    continue;
                }

                for (var rowIndex = item.StartRowIndex; rowIndex <= item.EndRowIndex; rowIndex++)
                {
                    var cellItem = rowCells[rowIndex][item.StartColumnIndex];
                    this.SetMerge(cellItem, W.gridSpan, (item.EndColumnIndex - item.StartColumnIndex + 1).ToString());
                    for (var columnIndex = item.EndColumnIndex; columnIndex >= item.StartColumnIndex + 1; columnIndex--)
                    {
                        rowCells[rowIndex][columnIndex].Remove();
                    }
                }
            }
        }

        void SetMerge(XElement cell, XName xName, string val)
        {
            var tcPr = cell.Descendants(W.tcPr).FirstOrDefault();
            if (tcPr != null)
            {
                var mergeElement = tcPr.Descendants(xName).FirstOrDefault();
                if (mergeElement != null)
                {
                    if (!String.IsNullOrEmpty(val))
                    {
                        mergeElement.SetAttributeValue(W.val, val);
                    }
                    else
                    {
                        var attr = mergeElement.Attribute(W.val);
                        if (attr != null)
                        {
                            attr.Remove();
                        }
                    }
                }
                else
                {
                    var child = new XElement(xName);
                    if (!String.IsNullOrEmpty(val))
                    {
                        child.SetAttributeValue(W.val, val);
                    }
                    tcPr.Add(child);
                }
            }
            else
            {
                var child = new XElement(xName);
                if (!String.IsNullOrEmpty(val))
                {
                    child.SetAttributeValue(W.val, val);
                }
                cell.Add(new XElement(W.tcPr, child));
            }

        }

        void RemoveMerge(XElement cell, XName xName)
        {
            var tcPr = cell.Descendants(W.tcPr).FirstOrDefault();
            if (tcPr != null)
            {
                var vMerge = tcPr.Descendants(xName).FirstOrDefault();
                if (vMerge != null)
                {
                    vMerge.Remove();
                }
            }
        }
    }
}
