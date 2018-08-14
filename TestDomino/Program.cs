using Domino;
using System.Runtime.InteropServices;

namespace TestDomino
{
    class Program
    {
        static void Main(string[] args)
        {
            NotesSession session = null;
            try
            {
                session = new NotesSession();
                session.Initialize("MonPetitJulien1");
                var database = session.GetDatabase("CancerS_A1/Serveurs/SSSS", @"TEST\PQDCS\PRINCIP.NSF", false);
                var collection = database.AllDocuments;
                var document = collection.GetFirstDocument();
                while (document != null)
                {
                    var form = document.GetItemValue("form");
                    foreach (NotesItem item in document.Items)
                    {
                    }
                    document = collection.GetNextDocument(document);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(session);
            }
        }
    }
}
