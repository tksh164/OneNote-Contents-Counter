using System;
using System.Xml;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.IO;

namespace OneNoteContentsCounter
{
    public class Program
    {
        private const String OneNoteXmlNamespace = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

        public static void Main(String[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine(@"Usage: ExeName.exe <NotebookNickName> [-x]");
                return;
            }

            String notebookName = args[0];

            Boolean isExportXml = false;
            if (args.Length >= 2)
            {
                if (args[1].Equals(@"-x", StringComparison.OrdinalIgnoreCase))
                {
                    isExportXml = true;
                }
            }

            //
            //

            OneNote.Application onenoteApp = new OneNote.Application();

            String hierarchyXml;
            onenoteApp.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out hierarchyXml);

            if (isExportXml)
            {
                exportXmlToFile(notebookName, hierarchyXml);
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(hierarchyXml);

            XmlNamespaceManager nsManager = new XmlNamespaceManager(xmlDoc.NameTable);
            nsManager.AddNamespace(@"one", OneNoteXmlNamespace);

            //
            //

            String notebookXpath = String.Format(@"//one:Notebook[@nickname='{0}']", notebookName);
            XmlNode notebookNode = xmlDoc.SelectSingleNode(notebookXpath, nsManager);

            // Sections
            XmlNodeList allSections = notebookNode.SelectNodes(@"(./one:Section|one:SectionGroup[@name!='OneNote_RecycleBin']//one:Section)", nsManager);

            // Section groups
            XmlNodeList allSectionGroups = notebookNode.SelectNodes(@"(./one:SectionGroup[@name!='OneNote_RecycleBin']|one:SectionGroup[@name!='OneNote_RecycleBin']//one:SectionGroup)", nsManager);

            // Pages
            XmlNodeList allPages = notebookNode.SelectNodes(@"(./*/one:Page|one:SectionGroup[@name!='OneNote_RecycleBin']//one:Page)", nsManager);

            //
            //

            Console.WriteLine(@"---- Summary ----");
            Console.WriteLine(@"Section: {0}", allSections.Count);
            Console.WriteLine(@"SectionGroup: {0}", allSectionGroups.Count);
            Console.WriteLine(@"Page: {0}", allPages.Count);

            Console.WriteLine(@"---- Sections ----");
            foreach (XmlNode section in allSections)
            {
                Console.WriteLine(@"{0} [LastModified:{1}]", section.Attributes[@"name"].Value, section.Attributes[@"lastModifiedTime"].Value);
            }

            Console.WriteLine(@"---- Section Groups ----");
            foreach (XmlNode sectionGroup in allSectionGroups)
            {
                Console.WriteLine(@"{0} [LastModified:{1}]", sectionGroup.Attributes[@"name"].Value, sectionGroup.Attributes[@"lastModifiedTime"].Value);
            }

            Console.WriteLine(@"---- Pages ----");
            foreach (XmlNode page in allPages)
            {
                Console.WriteLine(@"{0} [Created:{1}, LastModified:{2}]", page.Attributes[@"name"].Value, page.Attributes[@"dateTime"].Value, page.Attributes[@"lastModifiedTime"].Value);
            }


            return;
        }

        private static void exportXmlToFile(String notebookName, String xmlText)
        {
            String xmlFileOutputPath = String.Format(@"{0}{1}{2}.xml", Directory.GetCurrentDirectory(), Path.DirectorySeparatorChar, notebookName);

            using (FileStream stream = new FileStream(xmlFileOutputPath, FileMode.Create, FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(stream))
            {
                writer.WriteLine(xmlText);
            }
        }
    }
}
