using System.Reflection;

namespace SowAutomationTool.Services
{
    public class TemplateProvider
    {
        private const string ExcelResourceName = "Sow_Automation.Templates.Excel_Template.xlsx";
        private const string WordResourceName = "Sow_Automation.Templates.Word_Template.docx";

        private readonly Assembly _assembly = Assembly.GetExecutingAssembly();

        public byte[] GetExcelTemplate()
        {
            return ReadResource(ExcelResourceName);
        }

        public byte[] GetWordTemplate()
        {
            return ReadResource(WordResourceName);
        }

        private byte[] ReadResource(string resourceName)
        {
            using var stream = _assembly.GetManifestResourceStream(resourceName)
                ?? throw new FileNotFoundException($"Embedded resource '{resourceName}' not found. Available: {string.Join(", ", _assembly.GetManifestResourceNames())}");

            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }
    }
}
