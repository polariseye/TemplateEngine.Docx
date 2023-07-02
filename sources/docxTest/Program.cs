using System;
using System.IO;
using TemplateEngine.Docx;

namespace docxTest
{
    class Program
    {
        static void Main(string[] args)
        {
            File.Delete("OutputDocument.docx");
            File.Copy("InputTemplate.docx", "OutputDocument.docx");

            var valuesToFill = new Content(
                //new FieldContent("Report date", DateTime.Now.ToString()),
                new RepeatContent("Report date", new Content(new CheckBoxContent("Report date")
                {
                    Checked = true,
                    Text = "nihao a",
                }), new Content(new CheckBoxContent("Report date")
                {
                    Checked = true,
                    Text = "nihao a222",
                }))
                );


            using (var outputDocument = new TemplateProcessor("OutputDocument.docx")
                .SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }
    }
}