using SolidWorks.Interop.swdocumentmgr;
using System;

namespace PropertyTypeChanger
{
    internal class Program
    {
        #region Private Methods

        private static void Main(string[] args)
        {
            // docManKey.txt is included in project but is not tracked by source control 

            string licenseKey = System.IO.File.ReadAllText("docManKey.txt");

            var testPart = $@"part.sldprt";
            var testDrawing = $@"drawing.slddrw";

            if (!System.IO.File.Exists(testPart)) return;
            if (!System.IO.File.Exists(testDrawing)) return;

            var docman = new SOLIDWORKSDocumentManager(licenseKey);

            var part = docman.OpenDocument(testPart);
            part.SetCustomPropertyType("Vendor Date", SwDmCustomInfoType.swDmCustomInfoText);
            docman.SaveDocument(part);
            docman.CloseDoc(part);

            var drawing = docman.OpenDocument(testDrawing);
            drawing.SetCustomPropertyType("Vendor Date", SwDmCustomInfoType.swDmCustomInfoText);
            docman.SaveDocument(drawing);
            docman.CloseDoc(drawing);

            Console.ReadLine();
        }

        #endregion
    }

 
}