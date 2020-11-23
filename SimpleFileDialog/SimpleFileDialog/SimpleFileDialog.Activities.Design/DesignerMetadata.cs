using System.Activities.Presentation.Metadata;
using System.ComponentModel;
using System.ComponentModel.Design;
using SimpleFileDialog.Activities.Design.Designers;
using SimpleFileDialog.Activities.Design.Properties;

namespace SimpleFileDialog.Activities.Design
{
    public class DesignerMetadata : IRegisterMetadata
    {
        public void Register()
        {
            var builder = new AttributeTableBuilder();
            builder.ValidateTable();

            var categoryAttribute = new CategoryAttribute($"{Resources.Category}");

            builder.AddCustomAttributes(typeof(SelectFolder), categoryAttribute);
            builder.AddCustomAttributes(typeof(SelectFolder), new DesignerAttribute(typeof(SelectFolderDesigner)));
            builder.AddCustomAttributes(typeof(SelectFolder), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(OpenFileDialog), categoryAttribute);
            builder.AddCustomAttributes(typeof(OpenFileDialog), new DesignerAttribute(typeof(OpenFileDialogDesigner)));
            builder.AddCustomAttributes(typeof(OpenFileDialog), new HelpKeywordAttribute(""));

            builder.AddCustomAttributes(typeof(SaveFileDialog), categoryAttribute);
            builder.AddCustomAttributes(typeof(SaveFileDialog), new DesignerAttribute(typeof(SaveFileDialogDesigner)));
            builder.AddCustomAttributes(typeof(SaveFileDialog), new HelpKeywordAttribute(""));


            MetadataStore.AddAttributeTable(builder.CreateTable());
        }
    }
}
