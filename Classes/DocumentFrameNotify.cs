using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Selection.Classes
{
    internal class DocumentFrameNotify : ksDocumentFrameNotify
    {
        public IKompasDocument kompasDocument { get; set; }
        public DocumentFrameNotify(IKompasDocument _kompasDocument)
        {
            kompasDocument = _kompasDocument;
        }

        public bool BeginPaint(IPaintObject PaintObj)
        {
            return true;
        }

        public bool ClosePaint(IPaintObject PaintObj)
        {
            return true;
        }

        public bool MouseDown(short NButton, short NShiftState, int X, int Y)
        {
            return true;
        }

        public bool MouseUp(short NButton, short NShiftState, int X, int Y)
        {
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kompasDocument;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selected = selectionManager.SelectedObjects;
            if (selected is object[])
            {
                foreach (IKompasAPIObject selectObject in selected)
                {
                    if (selectObject is ICircle)
                    {
                        selectionManager.Unselect(selectObject);
                    }
                }
            }
            
            return true;
        }

        public bool MouseDblClick(short NButton, short NShiftState, int X, int Y)
        {
            return true;
        }

        public bool BeginPaintGL(ksGLObject GlObj, int DrawMode)
        {
            return true;
        }

        public bool ClosePaintGL(ksGLObject GlObj, int DrawMode)
        {
            return true;
        }

        public bool AddGabarit(IGabaritObject GabObj)
        {
            return true;
        }

        public bool Activate()
        {
            return true;
        }

        public bool Deactivate()
        {
            return true;
        }

        public bool CloseFrame()
        {
            return true;
        }

        public bool MouseMove(short NShiftState, int X, int Y)
        {
            return true;
        }

        public bool ShowOcxTree(object Ocx, bool Show)
        {
            return true;
        }

        public bool BeginPaintTmpObjects()
        {
            return true;
        }

        public bool ClosePaintTmpObjects()
        {
            return true;
        }
    }
}
