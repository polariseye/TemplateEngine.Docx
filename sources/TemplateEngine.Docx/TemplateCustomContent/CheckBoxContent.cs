using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateEngine.Docx
{
    [ContentItemName("CheckBox")]
    public class CheckBoxContent : HiddenContent<CheckBoxContent>, IEquatable<CheckBoxContent>
    {
        public bool Checked { get; set; }


        public CheckBoxContent(string name)
        {
            base.Name = name;
        }

        public CheckBoxContent(string name, bool @checked)
        {
            this.Name = name;
            this.Checked = @checked;
        }

        #region Equals

        public bool Equals(CheckBoxContent other)
        {
            if (other == null) return false;

            return Checked == other.Checked
                   ;
        }

        public override bool Equals(IContentItem other)
        {
            if (!(other is CheckBoxContent)) return false;

            return Equals((CheckBoxContent)other);
        }

        public override int GetHashCode()
        {
            return new { Checked}.GetHashCode();
        }

        #endregion
    }
}
