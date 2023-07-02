using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateEngine.Docx
{
    [ContentItemName("MergeTableCell")]
    public class MergeTableCellContent
    {
        public Int32 StartRowIndex { get; set; }
        public Int32 EndRowIndex { get; set; }
        public Int32 StartColumnIndex { get; set; }
        public Int32 EndColumnIndex { get; set; }

        public MergeTableCellContent(int startRowIndex, Int32 endRowIndex, Int32 startColumnIndex, Int32 endColumnIndex)
        {
            this.StartRowIndex = startRowIndex;
            this.EndRowIndex = endRowIndex;
            this.StartColumnIndex = startColumnIndex;
            this.EndColumnIndex = endColumnIndex;
        }
    }
}
