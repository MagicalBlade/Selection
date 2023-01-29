using Kompas6API5;
using KompasAPI7;
using Kompas6Constants;
using Microsoft.Win32;
using Selection.Windows;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using W = Selection.Windows;

namespace Selection
{
    public class Main
    {
        KompasObject kompas;
        IApplication application;
        IKompasDocument activeDocument;



        // Имя библиотеки
        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "Выбор";
        }

        [return: MarshalAs(UnmanagedType.BStr)]

        #region Формируем меню команд
        public string ExternalMenuItem(short number, ref short itemType, ref short command)
        {
            string result = string.Empty;
            itemType = 1; // "MENUITEM"
            switch (number)
            {
                case 1:
                    result = "Основная линия";
                    command = 1;
                    break;
                case 2:
                    result = "Разрушить макроэлементы";
                    command = 1;
                    break;
                case 3:
                    result = "Выбрать по типу линии";
                    command = 1;
                    break;
                case 4:
                    result = "Закрыть не сохраняясь";
                    command = 1;
                    break;
                case 5:
                    result = "Разрыв вида";
                    command = 1;
                    break;
            }
            return result;
        }

        #endregion

        #region Команды


        /// <summary>
        /// Разрушить макроэлементы
        /// </summary>
        private void DestroyMacroElements()
        {
            ksDocument2D document2DAPI5 = kompas.ActiveDocument2D();
            document2DAPI5.ksEnableUndo(true);
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)activeDocument;
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)activeDocument;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selectdynamic = selectionManager.SelectedObjects;
            if (selectdynamic != null)
            {
                if (selectdynamic is object[])
                {
                    foreach (IKompasAPIObject kompasobject in selectdynamic)
                    {
                        if (kompasobject.Type == Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectMacroObject)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                    }
                }
                else
                {
                    document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                }
            }
            else
            {
                IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.ActiveView;
                IDrawingContainer drawingContainer = (IDrawingContainer)view;
                IMacroObjects macroObjects = drawingContainer.MacroObjects;
                while (macroObjects.Count != 0)
                {
                    foreach (IMacroObject macroObject in macroObjects)
                    {
                        document2DAPI5.ksDestroyObjects(macroObject.Reference);
                    }
                    macroObjects = drawingContainer.MacroObjects;
                }
            }
            document2DAPI5.ksEnableUndo(false);
        }

        /// <summary>
        /// Выбор типа линии
        /// </summary>
        private void SelecetType()
        {
            IStylesManager stylesManager = (IStylesManager)application;
            IStyles styles = stylesManager.CurvesStyles;
            W.CurvesStyle curveStyle = new W.CurvesStyle();
            string[] selected_lb = new string[styles.Count + 1];
            selected_lb[0] = "Выбрать всё";
            for (int i = 1; i < selected_lb.Length; i++)
            {
                selected_lb[i] = styles[i - 1].Name;
            }
            curveStyle.lb_CurvesStyle.Items.AddRange(selected_lb);
            curveStyle.lb_CurvesStyle.SetSelected(0, true);
            var result = curveStyle.ShowDialog();
            int[] selectedType;
            if (curveStyle.lb_CurvesStyle.SelectedIndices[0] == 0)
            {
                selectedType = new int[selected_lb.Length];
                for (int i = 0; i < selected_lb.Length; i++)
                {
                    selectedType[i] = i;
                }
            }
            else
            {
                selectedType = new int[curveStyle.lb_CurvesStyle.SelectedIndices.Count];
                for (int i = 0; i < selectedType.Length; i++)
                {
                    selectedType[i] = curveStyle.lb_CurvesStyle.SelectedIndices[i];
                }
            }
            if (result == DialogResult.OK)
            {
                Select(selectedType);
            }
        }

        #endregion
        /// <summary>
        /// Выбор лини по указаному типу
        /// </summary>
        /// <param name="typeLine"></param>
        private void Select(int[] typeLine)
        {
            List<IDrawingObject> selectobjects = new List<IDrawingObject>();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)activeDocument;

            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selectdynamic = selectionManager.SelectedObjects;
            if (selectdynamic is object[])
            {
                object[] array = selectdynamic;
                foreach (IKompasAPIObject kompasobject in array)
                {
                    switch (kompasobject.Type)
                    {
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectArc:
                            {
                                IArc temp = (IArc)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectBeziers:
                            {
                                IBezier temp = (IBezier)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectCircle:
                            {
                                ICircle temp = (ICircle)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectConicCurve:
                            {
                                IConicCurve temp = (IConicCurve)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectDrawingContour:
                            {
                                IDrawingContour temp = (IDrawingContour)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEllipse:
                            {
                                IEllipse temp = (IEllipse)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEllipseArc:
                            {
                                IEllipseArc temp = (IEllipseArc)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEquidistant:
                            {
                                IEquidistant temp = (IEquidistant)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectLineSegment:
                            {
                                ILineSegment temp = (ILineSegment)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectNurbs:
                            {
                                INurbs temp = (INurbs)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectPolyLine2D:
                            {
                                IPolyLine2D temp = (IPolyLine2D)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectRectangle:
                            {
                                IRectangle temp = (IRectangle)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectRegularPolygon:
                            {
                                IRegularPolygon temp = (IRegularPolygon)kompasobject;
                                if (Array.IndexOf(typeLine, temp.Style) == -1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        default:
                            {
                                selectionManager.Unselect(kompasobject);
                                break;
                            }
                    }
                }
                return;
            }

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)activeDocument;
            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView view = views.ActiveView;
            IDrawingContainer drawingContainer = (IDrawingContainer)view;

            #region Проверка графических элементов на тип линии

            IArcs arcs = drawingContainer.Arcs;
            foreach (IArc arc in arcs)
            {
                if (Array.IndexOf(typeLine, arc.Style) != -1)
                {
                    selectobjects.Add(arc);
                }
            }
            IBeziers beziers = drawingContainer.Beziers;
            foreach (IBezier bezier in beziers)
            {
                if (Array.IndexOf(typeLine, bezier.Style) != -1)
                {
                    selectobjects.Add(bezier);
                }
            }
            ICircles circles = drawingContainer.Circles;
            foreach (ICircle circle in circles)
            {
                if (Array.IndexOf(typeLine, circle.Style) != -1)
                {
                    selectobjects.Add(circle);
                }
            }
            IConicCurves conicCurves = drawingContainer.ConicCurves;
            foreach (IConicCurve conicCurve in conicCurves)
            {
                if (Array.IndexOf(typeLine, conicCurve.Style) != -1)
                {
                    selectobjects.Add(conicCurve);
                }
            }
            IDrawingContours drawingContours = drawingContainer.DrawingContours;
            foreach (IDrawingContour drawingContour in drawingContours)
            {
                if (Array.IndexOf(typeLine, drawingContour.Style) != -1)
                {
                    selectobjects.Add(drawingContour);
                }
            }
            IEllipses ellipses = drawingContainer.Ellipses;
            foreach (IEllipse ellipse in ellipses)
            {
                if (Array.IndexOf(typeLine, ellipse.Style) != -1)
                {
                    selectobjects.Add(ellipse);
                }
            }
            IEllipseArcs ellipseNurbses = drawingContainer.EllipseArcs;
            foreach (IEllipseArc ellipseNurbse in ellipseNurbses)
            {
                if (Array.IndexOf(typeLine, ellipseNurbse.Style) != -1)
                {
                    selectobjects.Add(ellipseNurbse);
                }
            }
            IEquidistants equidistants = drawingContainer.Equidistants;
            foreach (IEquidistant equidistant in equidistants)
            {
                if (Array.IndexOf(typeLine, equidistant.Style) != -1)
                {
                    selectobjects.Add(equidistant);
                }
            }
            ILineSegments lineSegments = drawingContainer.LineSegments;
            foreach (ILineSegment lineSegment in lineSegments)
            {
                if (Array.IndexOf(typeLine, lineSegment.Style) != -1)
                {
                    selectobjects.Add(lineSegment);
                }
            }
            INurbses Nurbses = drawingContainer.Nurbses;
            foreach (INurbs nurbse in Nurbses)
            {
                if (Array.IndexOf(typeLine, nurbse.Style) != -1)
                {
                    selectobjects.Add(nurbse);
                }
            }
            INurbses NurbsesPoint = drawingContainer.NurbsesByPoints;
            foreach (INurbs nurbse in NurbsesPoint)
            {
                if (Array.IndexOf(typeLine, nurbse.Style) != -1)
                {
                    selectobjects.Add(nurbse);
                }
            }
            IPolyLines2D polyLines2Ds = drawingContainer.PolyLines2D;
            foreach (IPolyLine2D polyLines2D in polyLines2Ds)
            {
                if (Array.IndexOf(typeLine, polyLines2D.Style) != -1)
                {
                    selectobjects.Add(polyLines2D);
                }
            }
            IRectangles rectangles = drawingContainer.Rectangles;
            foreach (IRectangle rectangle in rectangles)
            {
                if (Array.IndexOf(typeLine, rectangle.Style) != -1)
                {
                    selectobjects.Add(rectangle);
                }
            }
            IRegularPolygons regularPolygons = drawingContainer.RegularPolygons;
            foreach (IRegularPolygon regularPolygon in regularPolygons)
            {
                if (Array.IndexOf(typeLine, regularPolygon.Style) != -1)
                {
                    selectobjects.Add(regularPolygon);
                }
            }

            #endregion

            selectionManager.Select(selectobjects.ToArray());


        }

        /// <summary>
        /// Закрыть документ не сохраняясь
        /// </summary>
        private void CloseNoSave()
        {
            activeDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
        }
        private void BreakView()
        {
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)application.ActiveDocument;
            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView view = views.ActiveView;
            IBreakViewParam breakViewParam = (IBreakViewParam)view;
            if (breakViewParam.BreaksCount == 0)
            {
                return;
            }
            if (breakViewParam.BreaksVisible == false)
            {
                breakViewParam.BreaksVisible = true;
            }
            else
            {
                breakViewParam.BreaksVisible = false;
            }
            view.Update();
        }

        // Головная функция библиотеки
        public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {
            kompas = (KompasObject)kompas_;
            application = (IApplication)kompas.ksGetApplication7();
            activeDocument = application.ActiveDocument;
            //Вызываем команды
            switch (command)
            {
                case 1: Select(new int[] {1}); break; //Выбор линии с типом "Основная"
                case 2: DestroyMacroElements(); break;
                case 3: SelecetType(); break;
                case 4: CloseNoSave(); break;
                case 5: BreakView(); break;
            }
        }

        public object ExternalGetResourceModule()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        public int ExternalGetToolBarId(short barType, short index)
        {
            int result = 0;

            if (barType == 0)
            {
                result = -1;
            }
            else
            {
                switch (index)
                {
                    case 1:
                        result = 3001;
                        break;
                    case 2:
                        result = -1;
                        break;
                }
            }

            return result;
        }


        #region COM Registration
        // Эта функция выполняется при регистрации класса для COM
        // Она добавляет в ветку реестра компонента раздел Kompas_Library,
        // который сигнализирует о том, что класс является приложением Компас,
        // а также заменяет имя InprocServer32 на полное, с указанием пути.
        // Все это делается для того, чтобы иметь возможность подключить
        // библиотеку на вкладке ActiveX.
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas_Library");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }

        // Эта функция удаляет раздел Kompas_Library из реестра
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            subKey.DeleteSubKey("Kompas_Library");
            subKey.Close();
        }
        #endregion
    }
}
