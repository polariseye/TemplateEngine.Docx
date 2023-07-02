using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
    internal class CheckBoxProcessor : IProcessor
    {
        private bool _isNeedToRemoveContentControls;

        public IProcessor SetRemoveContentControls(bool isNeedToRemove)
        {
            _isNeedToRemoveContentControls = isNeedToRemove;
            return this;
        }

        public ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items)
        {
            var processResult = ProcessResult.NotHandledResult;

            foreach (var contentItem in items)
            {
                processResult.Merge(FillContent(contentControl, contentItem));
            }

            if (processResult.Success && _isNeedToRemoveContentControls)
                contentControl.RemoveContentControl();

            return processResult;
        }

        public ProcessResult FillContent(XElement contentControl, IContentItem item)
        {
            var processResult = ProcessResult.NotHandledResult;
            if (!(item is CheckBoxContent))
            {
                processResult = ProcessResult.NotHandledResult;
                return processResult;
            }

            var field = item as CheckBoxContent;

            // If there isn't a field with that name, add an error to the error string,
            // and continue with next field.
            if (contentControl == null)
            {
                processResult.AddError(new ContentControlNotFoundError(field));
                return processResult;
            }

            if (this.FillCheckBox(contentControl, field.Checked))
            {
                processResult.AddItemToHandled(item);
            }
            else
            {
                processResult = ProcessResult.NotHandledResult;
                return processResult;
            }

            return processResult;
        }

        /*
          private List<Run> GetCheckBox(CheckBoxContent wordCheckbox)
         {
             var result = new List<Run>();
             result.Add(new Run(
                 new FieldChar(
                     new FormFieldData(
                         //new FormFieldName() { Val = internalName },
                         new Enabled(),
                         new CalculateOnExit() { Val = OnOffValue.FromBoolean(false) },
                         new CheckBox(
                             wordCheckbox.AutoSize || wordCheckbox.Size <= 0 ? new AutomaticallySizeFormField() as OpenXmlElement : new FormFieldSize() { Val = wordCheckbox.Size.ToString() } as OpenXmlElement,
                             new DefaultCheckBoxFormFieldState() { Val = OnOffValue.FromBoolean(wordCheckbox.Checked) }))
                 )
                 {
                     FieldCharType = FieldCharValues.Begin
                 }
             ));
             result.Add(new Run(new FieldCode(" FORMCHECKBOX ") { Space = SpaceProcessingModeValues.Preserve }));
             result.Add(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
             if (!String.IsNullOrEmpty(wordCheckbox.Text))
             {
                 result.Add(new Run(new Text(wordCheckbox.Text)));
             }

             return result;
         }
         */

        bool FillCheckBox(XElement contentControl, bool isChecked)
        {
            var sdtContentElement = contentControl.Element(W.sdtContent);
            if (sdtContentElement == null)
            {
                return false;
            }
                        
            var chkBox = sdtContentElement.DescendantsAndSelf(W.checkBox).FirstOrDefault();
            if (chkBox == null)
            {
                return false;
            }

            var valElement = chkBox.Element(W.@checked);
            if (valElement != null)
            {
                valElement.SetAttributeValue(W.val, isChecked ? 1 : 0);
            }
            else
            {
                var val = new XElement(W.@checked);
                val.SetAttributeValue(W.val, isChecked ? 1 : 0);
                chkBox.Add(val);
            }

            var defaultElement = chkBox.Element(W.@default);
            if (defaultElement != null)
            {
                defaultElement.SetAttributeValue(W.val, isChecked ? 1 : 0);
            }
            else
            {
                var val = new XElement(W.@default);
                val.SetAttributeValue(W.val, isChecked ? 1 : 0);
                chkBox.Add(val);
            }

            return true;
        }
    }
}
