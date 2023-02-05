using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KompasAPI7;
using Kompas6API5;
using System.Runtime.InteropServices.ComTypes;

namespace Selection.Classes
{
    internal class KompasEvents : ksKompasObjectNotify
    {
        public KompasObject kompas { get; set; }
        public KompasEvents(KompasObject _kompas)
        {
            kompas = _kompas;
        }

        public bool CreateDocument(object newDoc, int docType)
        {
            return true;
        }

        public bool BeginOpenDocument(string fileName)
        {
            return true;
        }

        public bool OpenDocument(object newDoc, int docType)
        {
            if (docType == 1)
            {
                ksDocument2D documentAPI5 = (ksDocument2D)newDoc;
                IKompasDocument kompasDocument = kompas.TransferReference(documentAPI5.reference, 0);
                
                IDocumentFrames documentFrames = kompasDocument.DocumentFrames;
                IDocumentFrame documentFrame = documentFrames[0];
                DocumentFrameNotify documentFrameNotify = new DocumentFrameNotify(kompasDocument);
                IConnectionPointContainer cpContainer = documentFrame as IConnectionPointContainer;
                IConnectionPoint m_ConnPt;
                int m_Cookie;
                if (cpContainer != null)
                {
                    cpContainer.FindConnectionPoint(typeof(ksDocumentFrameNotify).GUID, out m_ConnPt);
                    if (m_ConnPt != null)
                    {
                        m_ConnPt.Advise(documentFrameNotify, out m_Cookie);
                    }
                }
                
            }
            return true;
        }

        public bool ChangeActiveDocument(object newDoc, int docType)
        {
            return true;
        }

        public bool ApplicationDestroy()
        {
            return true;
        }

        public bool BeginCreate(int docType)
        {
            return true;
        }

        public bool BeginOpenFile()
        {
            return true;
        }

        public bool BeginCloseAllDocument()
        {
            return true;
        }

        public bool KeyDown(ref int key, int flags, bool systemKey)
        {
            return true;
        }

        public bool KeyUp(ref int key, int flags, bool systemKey)
        {
            return true;
        }

        public bool KeyPress(ref int key, bool systemKey)
        {
            return true;
        }

        public bool BeginReguestFiles(int requestID, ref object files)
        {
            return true;
        }

        public bool BeginChoiceMaterial(int MaterialPropertyId)
        {
            return true;
        }

        public bool ChoiceMaterial(int MaterialPropertyId, string material, double density)
        {
            return true;
        }

        public bool IsNeedConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam, ref bool needConvert)
        {
            return true;
        }

        public bool BeginConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
        {
            return true;
        }

        public bool EndConvertToSavePrevious(object pDoc, int docType, int saveVersion, object saveToPreviusParam)
        {
            return true;
        }

        public bool ChangeTheme(int newTheme)
        {
            return true;
        }
    }
}
