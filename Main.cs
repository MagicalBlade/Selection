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
using Selection.Classes;
using System.Runtime.InteropServices.ComTypes;
using System.Diagnostics;
using System.IO;

namespace Selection
{
    public class Main
    {
        KompasObject Kompas;
        IApplication Application;
        IKompasDocument ActiveDocument;



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
                    result = "Выбрать по типу элемента";
                    command = 1;
                    break;
                case 5:
                    result = "Выбрать графику";
                    command = 1;
                    break;
                case 6:
                    result = "Выбрать оформление";
                    command = 1;
                    break;
                case 7:
                    command = -1;
                    itemType = 8; // "ENDMENU"
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
            ksDocument2D document2DAPI5 = Kompas.ActiveDocument2D();
            document2DAPI5.ksUndoContainer(true);
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)ActiveDocument;
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)ActiveDocument;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selectdynamic = selectionManager.SelectedObjects;
            if (selectdynamic != null)
            {
                if (selectdynamic is object[])
                {
                    foreach (IKompasAPIObject kompasobject in selectdynamic)
                    {
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectInsertionFragment)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectMacroObject)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectDrawingContour)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectEquidistant)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectMultiline)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectPolyLines2D)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectRectangle)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                        if (kompasobject.Type == KompasAPIObjectTypeEnum.ksObjectRegularPolygon)
                        {
                            document2DAPI5.ksDestroyObjects(kompasobject.Reference);
                        }
                    }
                    Application.MessageBoxEx("Элементы разрушены", "Сообщение", 64);
                }
                else
                {
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectInsertionFragment)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectMacroObject)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectDrawingContour)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectEquidistant)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectMultiline)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectPolyLines2D)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectRectangle)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    if (selectdynamic.Type == (int)KompasAPIObjectTypeEnum.ksObjectRegularPolygon)
                    {
                        document2DAPI5.ksDestroyObjects(selectdynamic.Reference);
                    }
                    Application.MessageBoxEx("Элементы разрушены", "Сообщение", 64);
                }
            }
            else
            {
                IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.ActiveView;
                IDrawingContainer drawingContainer = (IDrawingContainer)view;
                IMacroObjects macroObjects = drawingContainer.MacroObjects;
                InsertionObjects insertionObjects = drawingContainer.InsertionObjects;
                
                while (macroObjects.Count != 0)
                {
                    foreach (IMacroObject macroObject in macroObjects)
                    {
                        document2DAPI5.ksDestroyObjects(macroObject.Reference);
                    }
                    foreach (IInsertionObject insertionObject in insertionObjects)
                    {
                        IInsertionFragment insertionFragment = (IInsertionFragment)insertionObject;
                        if (insertionFragment != null)
                        {
                            document2DAPI5.ksDestroyObjects(insertionFragment.Reference);
                        }
                    }
                    insertionObjects = drawingContainer.InsertionObjects;
                    macroObjects = drawingContainer.MacroObjects;
                }

                foreach (IDrawingContour item in drawingContainer.DrawingContours)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                foreach (IEquidistant item in drawingContainer.Equidistants)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                foreach (IMultiline item in drawingContainer.Multilines)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                foreach (IPolyLine2D item in drawingContainer.PolyLines2D)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                foreach (IRectangle item in drawingContainer.Rectangles)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                foreach (IRegularPolygon item in drawingContainer.RegularPolygons)
                {
                    document2DAPI5.ksDestroyObjects(item.Reference);
                }
                Application.MessageBoxEx("Элементы разрушены", "Сообщение", 64);
            }
            document2DAPI5.ksUndoContainer(false);

        }

        /// <summary>
        /// Выбор типа линии
        /// </summary>
        private void SelecetTypeLine()
        {
            IStylesManager stylesManagerApp = (IStylesManager)Application;
            //IStylesManager stylesManagerADoc = (IStylesManager)activeDocument;
            IStyles styles = stylesManagerApp.CurvesStyles;
            W.TypeObject curveStyle = new W.TypeObject();
            string[] selected_lb = new string[styles.Count + 1];
            selected_lb[0] = "Выбрать всё";
            for (int i = 1; i < selected_lb.Length; i++)
            {
                selected_lb[i] = styles[i - 1].Name;
            }
            curveStyle.lb_Type.Items.AddRange(selected_lb);
            curveStyle.lb_Type.SetSelected(0, true);
            var result = curveStyle.ShowDialog();
            int[] selectedType;
            if (curveStyle.lb_Type.SelectedIndices[0] == 0)
            {
                selectedType = new int[selected_lb.Length];
                for (int i = 0; i < selected_lb.Length; i++)
                {
                    selectedType[i] = i;
                }
            }
            else
            {
                selectedType = new int[curveStyle.lb_Type.SelectedIndices.Count];
                for (int i = 0; i < selectedType.Length; i++)
                {
                    selectedType[i] = curveStyle.lb_Type.SelectedIndices[i];
                }
            }
            if (result == DialogResult.OK)
            {
                Select(selectedType);
            }
        }

        /// <summary>
        /// Выбор лини по указаному стилю
        /// </summary>
        /// <param name="typeLine"></param>
        private void Select(int[] typeLine)
        {
            List<IDrawingObject> selectobjects = new List<IDrawingObject>();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)ActiveDocument;
            ksDocument2D ksDocument2D = Kompas.ActiveDocument2D();
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selectdynamic = selectionManager.SelectedObjects;
            ksDocument2D.ksUndoContainer(true);
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

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)ActiveDocument;
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

            ksDocument2D.ksUndoContainer(false);
        }

        /// <summary>
        /// Выбрать по типу элемента
        /// </summary>
        private void SelectTypeObjectCommand()
        {
            List<object> kompasObjects = new List<object>();
            Dictionary<string, List<object>> objectDictionary = new Dictionary<string, List<object>>()
            {
                { "Геометрия",  new List<object>()},
                { "Оформление",  new List<object>()},
                { "Дуги",  new List<object>()},
                { "Кривые Безье",  new List<object>()},
                { "Окружности",  new List<object>()},
                { "Заливки",  new List<object>()},
                { "Конические кривые",  new List<object>()},
                { "Контуры",  new List<object>()},
                { "Текст",  new List<object>()},
                { "Элипсы",  new List<object>()},
                { "Дуги элипсов",  new List<object>()},
                { "Эквидистанты",  new List<object>()},
                { "Штриховки",  new List<object>()},
                { "Фрагменты видов другого чертежа",  new List<object>()},
                { "Прямые",  new List<object>()},
                { "Отрезки",  new List<object>()},
                { "Вставной фрагмент",  new List<object>()},
                { "Макроэлементы",  new List<object>()},
                { "Мультилинии",  new List<object>()},
                { "Nurbs сплайны",  new List<object>()},
                { "OLE объекты",  new List<object>()},
                { "Точки",  new List<object>()},
                { "2D ломанные",  new List<object>()},
                { "Растровые объекты",  new List<object>()},
                { "Прямоугольники",  new List<object>()},
                { "Многоугольники",  new List<object>()},
                { "Угловые размеры",  new List<object>()},
                { "Размеры дуг окружностей",  new List<object>()},
                { "Ассоциативные таблицы отчетов",  new List<object>()},
                { "Осевые линии",  new List<object>()},
                { "Обозначение базы",  new List<object>()},
                { "Линейные размеры с обрывом",  new List<object>()},
                { "Радиальные размеры с изломом",  new List<object>()},
                { "Линии обрыва с изломом",  new List<object>()},
                { "Обозначения центра",  new List<object>()},
                { "Круговые сетки центров",  new List<object>()},
                { "Условные пересечения",  new List<object>()},
                { "Линии разреза/сечения",  new List<object>()},
                { "Диаметральные размеры",  new List<object>()},
                { "Таблицы",  new List<object>()},
                { "Размеры высоты",  new List<object>()},
                { "Линии-выноски",  new List<object>()},
                { "Линия-выноска для обозначения позиции",  new List<object>()},
                { "Линия-выноска для для обозначения клеймения",  new List<object>()},
                { "Линия-выноска для обозначения маркирования",  new List<object>()},
                { "Знак изменения",  new List<object>()},
                { "Линейные размеры",  new List<object>()},
                { "Линейные сетки центров",  new List<object>()},
                { "Радиальные размеры",  new List<object>()},
                { "Выносные элементы",  new List<object>()},
                { "Обозначения шероховатости",  new List<object>()},
                { "Допуски формы",  new List<object>()},
                { "Стрелки взгляда",  new List<object>()},
                { "Волнистые линии",  new List<object>()},
                { "Фигурные скобки",  new List<object>()},
                { "Прямые строительные оси",  new List<object>()},
                { "Круговые строительные оси",  new List<object>()},
                { "Дуговые строительные оси",  new List<object>()},
                { "Линии разреза СПДС",  new List<object>()},
                { "Обозначения узла в сечении",  new List<object>()},
                { "Марки",  new List<object>()},
                { "Выносные надписи к многослойным конструкциям",  new List<object>()},
                { "Обозначения узлов",  new List<object>()},
                { "Номера узлов",  new List<object>()}
            };
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)Application.ActiveDocument;
            ksDocument2D activeDocumentAPI5 = Kompas.ActiveDocument2D();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kompasDocument2D;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selected = selectionManager.SelectedObjects;
            if (selected is object[])
            {
                foreach (IDrawingObject item in selected)
                {
                    switch (item.DrawingObjectType)
                    {
                        case DrawingObjectTypeEnum.ksDrLineSeg:
                            objectDictionary["Отрезки"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCircle:
                            objectDictionary["Окружности"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrArc:
                            objectDictionary["Дуги"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrDrawText:
                            objectDictionary["Текст"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrPoint:
                            objectDictionary["Точки"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrHatch:
                            objectDictionary["Штриховки"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBezier:
                            objectDictionary["Кривые Безье"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrLDimension:
                            objectDictionary["Линейные размеры"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrADimension:
                            objectDictionary["Угловые размеры"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrDDimension:
                            objectDictionary["Диаметральные размеры"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRDimension:
                            objectDictionary["Радиальные размеры"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRBreakDimension:
                            objectDictionary["Радиальные размеры с изломом"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRough:
                            objectDictionary["Обозначения шероховатости"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBase:
                            objectDictionary["Обозначение базы"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrWPointer:
                            objectDictionary["Стрелки взгляда"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCut:
                            objectDictionary["Линии разреза/сечения"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrLeader:
                            objectDictionary["Линии-выноски"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrPosLeader:
                            objectDictionary["Линия-выноска для обозначения позиции"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBrandLeader:
                            objectDictionary["Линия-выноска для для обозначения клеймения"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkerLeader:
                            objectDictionary["Линия-выноска для обозначения маркирования"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrChangeLeader:
                            objectDictionary["Знак изменения"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrTolerance:
                            objectDictionary["Допуски формы"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrTable:
                            objectDictionary["Таблицы"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrContour:
                            objectDictionary["Контуры"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrFragment:
                            objectDictionary["Вставной фрагмент"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMacro:
                            objectDictionary["Макроэлементы"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrLine:
                            objectDictionary["Прямые"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrPolyline:
                            objectDictionary["2D ломанные"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrEllipse:
                            objectDictionary["Элипсы"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrNurbs:
                            objectDictionary["Nurbs сплайны"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrEllipseArc:
                            objectDictionary["Дуги элипсов"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRectangle:
                            objectDictionary["Прямоугольники"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRegularPolygon:
                            objectDictionary["Многоугольники"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrEquid:
                            objectDictionary["Эквидистанты"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrLBreakDimension:
                            objectDictionary["Линейные размеры с обрывом"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrOrdinateDimension:
                            objectDictionary["Размеры высоты"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrColorFill:
                            objectDictionary["Заливки"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCentreMarker:
                            objectDictionary["Обозначения центра"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrArcDimension:
                            objectDictionary["Размеры дуг окружностей"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRaster:
                            objectDictionary["Растровые объекты"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrRemoteElement:
                            objectDictionary["Выносные элементы"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrAxisLine:
                            objectDictionary["Осевые линии"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrOLEObject:
                            objectDictionary["OLE объекты"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrUnitNumber:
                            objectDictionary["Номера узлов"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBrace:
                            objectDictionary["Фигурные скобки"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkOnLeader:
                            objectDictionary["Марки"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkOnLine:
                            objectDictionary["Марки"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkInsideForm:
                            objectDictionary["Марки"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrWaveLine:
                            objectDictionary["Волнистые линии"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrStraightAxis:
                            objectDictionary["Прямые строительные оси"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBrokenLine:
                            objectDictionary["Линии обрыва с изломом"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCircleAxis:
                            objectDictionary["Круговые строительные оси"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrArcAxis:
                            objectDictionary["Дуговые строительные оси"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCutUnitMarking:
                            objectDictionary["Обозначения узла в сечении"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrUnitMarking:
                            objectDictionary["Обозначения узлов"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMultiTextLeader:
                            objectDictionary["Выносные надписи к многослойным конструкциям"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrExternalView:
                            objectDictionary["Фрагменты видов другого чертежа"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrMultiLine:
                            objectDictionary["Мультилинии"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrBuildingCutLine:
                            objectDictionary["Линии разреза СПДС"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrConditionCrossing:
                            objectDictionary["Условные пересечения"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksReportTable:
                            objectDictionary["Ассоциативные таблицы отчетов"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrNurbsByPoints:
                            objectDictionary["Nurbs сплайны"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrConicCurve:
                            objectDictionary["Конические кривые"].Add(item);
                            objectDictionary["Геометрия"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrCircularCentres:
                            objectDictionary["Круговые сетки центров"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        case DrawingObjectTypeEnum.ksDrLinearCentres:
                            objectDictionary["Линейные сетки центров"].Add(item);
                            objectDictionary["Оформление"].Add(item);
                            break;
                        default:
                            break;
                    }
                }
            }
            else if (selected == null)
            {
                IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.ActiveView;
                IDrawingContainer drawingContainer = (IDrawingContainer)view;

                #region Заполняем словарь коллекциями элементов
                /// IDrawingContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArc] != null)
                {
                    objectDictionary["Дуги"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArc]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArc]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBezier] != null)
                {
                    objectDictionary["Кривые Безье"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBezier]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBezier]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircle] != null)
                {
                    objectDictionary["Окружности"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircle]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircle]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrColorFill] != null)
                {
                    objectDictionary["Заливки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrColorFill]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrColorFill]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConicCurve] != null)
                {
                    objectDictionary["Конические кривые"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConicCurve]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConicCurve]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrContour] != null)
                {
                    objectDictionary["Контуры"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrContour]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrContour]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipse] != null)
                {
                    objectDictionary["Элипсы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipse]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipse]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipseArc] != null)
                {
                    objectDictionary["Дуги элипсов"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipseArc]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipseArc]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEquid] != null)
                {
                    objectDictionary["Эквидистанты"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEquid]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEquid]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrHatch] != null)
                {
                    objectDictionary["Штриховки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrHatch]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrHatch]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrExternalView] != null)
                {
                    objectDictionary["Фрагменты видов другого чертежа"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrExternalView]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLine] != null)
                {
                    objectDictionary["Прямые"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLine]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLineSeg] != null)
                {
                    objectDictionary["Отрезки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLineSeg]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLineSeg]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrFragment] != null)
                {
                    objectDictionary["Вставной фрагмент"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrFragment]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrFragment]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMacro] != null)
                {
                    objectDictionary["Макроэлементы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMacro]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMacro]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiLine] != null)
                {
                    objectDictionary["Мультилинии"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiLine]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbs] != null)
                {
                    objectDictionary["Nurbs сплайны"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbs]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbs]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbsByPoints] != null)
                {
                    objectDictionary["Nurbs сплайны"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbsByPoints]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbsByPoints]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOLEObject] != null)
                {
                    objectDictionary["OLE объекты"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOLEObject]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPoint] != null)
                {
                    objectDictionary["Точки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPoint]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPoint]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPolyline] != null)
                {
                    objectDictionary["2D ломанные"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPolyline]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPolyline]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRaster] != null)
                {
                    objectDictionary["Растровые объекты"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRaster]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRectangle] != null)
                {
                    objectDictionary["Прямоугольники"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRectangle]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRectangle]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRegularPolygon] != null)
                {
                    objectDictionary["Многоугольники"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRegularPolygon]);
                    objectDictionary["Геометрия"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRegularPolygon]);
                }
                //ISymbols2DContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDrawText] != null)
                {
                    objectDictionary["Текст"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDrawText]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDrawText]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrADimension] != null)
                {
                    objectDictionary["Угловые размеры"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrADimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrADimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcDimension] != null)
                {
                    objectDictionary["Размеры дуг окружностей"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksReportTable] != null)
                {
                    objectDictionary["Ассоциативные таблицы отчетов"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksReportTable]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksReportTable]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrAxisLine] != null)
                {
                    objectDictionary["Осевые линии"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrAxisLine]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrAxisLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBase] != null)
                {
                    objectDictionary["Обозначение базы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBase]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBase]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLBreakDimension] != null)
                {
                    objectDictionary["Линейные размеры с обрывом"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLBreakDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLBreakDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRBreakDimension] != null)
                {
                    objectDictionary["Радиальные размеры с изломом"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRBreakDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRBreakDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrokenLine] != null)
                {
                    objectDictionary["Линии обрыва с изломом"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrokenLine]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrokenLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCentreMarker] != null)
                {
                    objectDictionary["Обозначения центра"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCentreMarker]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCentreMarker]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircularCentres] != null)
                {
                    objectDictionary["Круговые сетки центров"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircularCentres]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircularCentres]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConditionCrossing] != null)
                {
                    objectDictionary["Условные пересечения"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConditionCrossing]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConditionCrossing]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCut] != null)
                {
                    objectDictionary["Линии разреза/сечения"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCut]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCut]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDDimension] != null)
                {
                    objectDictionary["Диаметральные размеры"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTable] != null)
                {
                    objectDictionary["Таблицы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTable]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTable]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOrdinateDimension] != null)
                {
                    objectDictionary["Размеры высоты"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOrdinateDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOrdinateDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLeader] != null)
                {
                    objectDictionary["Линии-выноски"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPosLeader] != null)
                {
                    objectDictionary["Линия-выноска для обозначения позиции"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPosLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPosLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrandLeader] != null)
                {
                    objectDictionary["Линия-выноска для для обозначения клеймения"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrandLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrandLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkerLeader] != null)
                {
                    objectDictionary["Линия-выноска для обозначения маркирования"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkerLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkerLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrChangeLeader] != null)
                {
                    objectDictionary["Знак изменения"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrChangeLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrChangeLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLDimension] != null)
                {
                    objectDictionary["Линейные размеры"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLinearCentres] != null)
                {
                    objectDictionary["Линейные сетки центров"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLinearCentres]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLinearCentres]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRDimension] != null)
                {
                    objectDictionary["Радиальные размеры"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRDimension]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRemoteElement] != null)
                {
                    objectDictionary["Выносные элементы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRemoteElement]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRemoteElement]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRough] != null)
                {
                    objectDictionary["Обозначения шероховатости"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRough]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRough]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTolerance] != null)
                {
                    objectDictionary["Допуски формы"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTolerance]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTolerance]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWPointer] != null)
                {
                    objectDictionary["Стрелки взгляда"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWPointer]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWPointer]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWaveLine] != null)
                {
                    objectDictionary["Волнистые линии"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWaveLine]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWaveLine]);
                }
                //IBuildingContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrace] != null)
                {
                    objectDictionary["Фигурные скобки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrace]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrace]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrStraightAxis] != null)
                {
                    objectDictionary["Прямые строительные оси"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrStraightAxis]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrStraightAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircleAxis] != null)
                {
                    objectDictionary["Круговые строительные оси"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircleAxis]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircleAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcAxis] != null)
                {
                    objectDictionary["Дуговые строительные оси"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcAxis]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBuildingCutLine] != null)
                {
                    objectDictionary["Линии разреза СПДС"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBuildingCutLine]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBuildingCutLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCutUnitMarking] != null)
                {
                    objectDictionary["Обозначения узла в сечении"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCutUnitMarking]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCutUnitMarking]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLeader] != null)
                {
                    objectDictionary["Марки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLine] != null)
                {
                    objectDictionary["Марки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLine]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkInsideForm] != null)
                {
                    objectDictionary["Марки"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkInsideForm]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkInsideForm]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiTextLeader] != null)
                {
                    objectDictionary["Выносные надписи к многослойным конструкциям"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiTextLeader]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiTextLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitMarking] != null)
                {
                    objectDictionary["Обозначения узлов"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitMarking]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitMarking]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitNumber] != null)
                {
                    objectDictionary["Номера узлов"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitNumber]);
                    objectDictionary["Оформление"].AddRange(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitNumber]);
                }
                #endregion

            }
            TypeObject typeObject = new TypeObject();
            foreach (string key in objectDictionary.Keys)
            {
                if (objectDictionary[key].Count != 0)
                {
                    typeObject.lb_Type.Items.Add(key);
                }
            }
            if (typeObject.lb_Type.Items.Count == 0)
            {
                return;
            }
            typeObject.lb_Type.SetSelected(0, true);
            var result = typeObject.ShowDialog();

            if (result == DialogResult.Cancel)
            {
                return;
            }

            activeDocumentAPI5.ksUndoContainer(true);
            selectionManager.UnselectAll();
            foreach (var item in typeObject.lb_Type.SelectedItems)
            {
                selectionManager.Select(objectDictionary[item.ToString()].ToArray());
            }
            activeDocumentAPI5.ksUndoContainer(false);
        }

        /// <summary>
        /// Выбрать только геометрию
        /// </summary>
        private void SelectTypeGraphic()
        {
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)Application.ActiveDocument;
            ksDocument2D activeDocumentAPI5 = Kompas.ActiveDocument2D();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kompasDocument2D;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selected = selectionManager.SelectedObjects;
            activeDocumentAPI5.ksUndoContainer(true);
            if (selected is object[])
            {
                foreach (IDrawingObject item in selected)
                {
                    switch (item.DrawingObjectType)
                    {
                        case DrawingObjectTypeEnum.ksDrLineSeg:
                            break;
                        case DrawingObjectTypeEnum.ksDrCircle:
                            break;
                        case DrawingObjectTypeEnum.ksDrArc:
                            break;
                        case DrawingObjectTypeEnum.ksDrPoint:
                            break;
                        case DrawingObjectTypeEnum.ksDrHatch:
                            break;
                        case DrawingObjectTypeEnum.ksDrBezier:
                            break;
                        case DrawingObjectTypeEnum.ksDrContour:
                            break;
                        case DrawingObjectTypeEnum.ksDrFragment:
                            break;
                        case DrawingObjectTypeEnum.ksDrMacro:
                            break;
                        case DrawingObjectTypeEnum.ksDrLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrPolyline:
                            break;
                        case DrawingObjectTypeEnum.ksDrEllipse:
                            break;
                        case DrawingObjectTypeEnum.ksDrNurbs:
                            break;
                        case DrawingObjectTypeEnum.ksDrEllipseArc:
                            break;
                        case DrawingObjectTypeEnum.ksDrRectangle:
                            break;
                        case DrawingObjectTypeEnum.ksDrRegularPolygon:
                            break;
                        case DrawingObjectTypeEnum.ksDrEquid:
                            break;
                        case DrawingObjectTypeEnum.ksDrColorFill:
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkOnLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrMultiLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrNurbsByPoints:
                            break;
                        case DrawingObjectTypeEnum.ksDrConicCurve:
                            break;
                        default:
                            selectionManager.Unselect(item);
                            break;
                    }
                }
            }
            else if (selected == null)
            {
                IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.ActiveView;
                IDrawingContainer drawingContainer = (IDrawingContainer)view;

                #region Выделяем
                /// IDrawingContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArc] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArc]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBezier] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBezier]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircle] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircle]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrColorFill] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrColorFill]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConicCurve] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConicCurve]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrContour] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrContour]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipse] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipse]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipseArc] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEllipseArc]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEquid] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrEquid]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrHatch] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrHatch]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLineSeg] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLineSeg]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrFragment] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrFragment]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMacro] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMacro]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbs] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbs]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbsByPoints] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrNurbsByPoints]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPoint] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPoint]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPolyline] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPolyline]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRectangle] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRectangle]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRegularPolygon] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRegularPolygon]);
                }
                #endregion

            }

            activeDocumentAPI5.ksUndoContainer(false);
            Application.MessageBoxEx("Элементы выделены", "Сообщение", 64);
        }

        /// <summary>
        /// Выбрать только оформление
        /// </summary>
        private void SelectTypeDimension()
        {
            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)Application.ActiveDocument;
            ksDocument2D activeDocumentAPI5 = Kompas.ActiveDocument2D();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)kompasDocument2D;
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selected = selectionManager.SelectedObjects;
            activeDocumentAPI5.ksUndoContainer(true);
            if (selected is object[])
            {
                foreach (IDrawingObject item in selected)
                {
                    switch (item.DrawingObjectType)
                    {
                        case DrawingObjectTypeEnum.ksDrDrawText:
                            break;
                        case DrawingObjectTypeEnum.ksDrLDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrADimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrDDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrRDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrRBreakDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrRough:
                            break;
                        case DrawingObjectTypeEnum.ksDrBase:
                            break;
                        case DrawingObjectTypeEnum.ksDrWPointer:
                            break;
                        case DrawingObjectTypeEnum.ksDrCut:
                            break;
                        case DrawingObjectTypeEnum.ksDrLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrPosLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrBrandLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkerLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrChangeLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrTolerance:
                            break;
                        case DrawingObjectTypeEnum.ksDrTable:
                            break;
                        case DrawingObjectTypeEnum.ksDrLBreakDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrOrdinateDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrCentreMarker:
                            break;
                        case DrawingObjectTypeEnum.ksDrArcDimension:
                            break;
                        case DrawingObjectTypeEnum.ksDrRaster:
                            break;
                        case DrawingObjectTypeEnum.ksDrRemoteElement:
                            break;
                        case DrawingObjectTypeEnum.ksDrAxisLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrOLEObject:
                            break;
                        case DrawingObjectTypeEnum.ksDrUnitNumber:
                            break;
                        case DrawingObjectTypeEnum.ksDrBrace:
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkOnLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkOnLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrMarkInsideForm:
                            break;
                        case DrawingObjectTypeEnum.ksDrWaveLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrStraightAxis:
                            break;
                        case DrawingObjectTypeEnum.ksDrBrokenLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrCircleAxis:
                            break;
                        case DrawingObjectTypeEnum.ksDrArcAxis:
                            break;
                        case DrawingObjectTypeEnum.ksDrCutUnitMarking:
                            break;
                        case DrawingObjectTypeEnum.ksDrUnitMarking:
                            break;
                        case DrawingObjectTypeEnum.ksDrMultiTextLeader:
                            break;
                        case DrawingObjectTypeEnum.ksDrExternalView:
                            break;
                        case DrawingObjectTypeEnum.ksDrBuildingCutLine:
                            break;
                        case DrawingObjectTypeEnum.ksDrConditionCrossing:
                            break;
                        case DrawingObjectTypeEnum.ksReportTable:
                            break;
                        case DrawingObjectTypeEnum.ksDrCircularCentres:
                            break;
                        case DrawingObjectTypeEnum.ksDrLinearCentres:
                            break;
                        default:
                            selectionManager.Unselect(item);
                            break;
                    }
                }
            }
            else if (selected == null)
            {
                IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
                IViews views = viewsAndLayersManager.Views;
                IView view = views.ActiveView;
                IDrawingContainer drawingContainer = (IDrawingContainer)view;

                #region Выделяем
                //ISymbols2DContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDrawText] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDrawText]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrADimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrADimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksReportTable] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksReportTable]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrAxisLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrAxisLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBase] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBase]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLBreakDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLBreakDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRBreakDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRBreakDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrokenLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrokenLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCentreMarker] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCentreMarker]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircularCentres] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircularCentres]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConditionCrossing] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrConditionCrossing]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCut] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCut]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrDDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTable] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTable]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOrdinateDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrOrdinateDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPosLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrPosLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrandLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrandLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkerLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkerLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrChangeLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrChangeLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLinearCentres] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrLinearCentres]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRDimension] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRDimension]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRemoteElement] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRemoteElement]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRough] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrRough]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTolerance] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrTolerance]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWPointer] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWPointer]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWaveLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrWaveLine]);
                }
                //IBuildingContainer
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrace] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBrace]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrStraightAxis] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrStraightAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircleAxis] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCircleAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcAxis] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrArcAxis]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBuildingCutLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrBuildingCutLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCutUnitMarking] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrCutUnitMarking]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLine] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkOnLine]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkInsideForm] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMarkInsideForm]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiTextLeader] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrMultiTextLeader]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitMarking] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitMarking]);
                }
                if (drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitNumber] != null)
                {
                    selectionManager.Select(drawingContainer.Objects[DrawingObjectTypeEnum.ksDrUnitNumber]);
                }
                #endregion

            }

            activeDocumentAPI5.ksUndoContainer(false);
            Application.MessageBoxEx("Элементы выделены", "Сообщение", 64);
        }

        /// <summary>
        /// Выбрать основную линию и обозначение маркировки
        /// </summary>
        private void SelectMainTypeLineMarkLeader()
        {
            List<object> selectobjects = new List<object>();
            IKompasDocument2D1 kompasDocument2D1 = (IKompasDocument2D1)ActiveDocument;
            ksDocument2D ksDocument2D = Kompas.ActiveDocument2D();
            ISelectionManager selectionManager = kompasDocument2D1.SelectionManager;
            dynamic selectdynamic = selectionManager.SelectedObjects;
            ksDocument2D.ksUndoContainer(true);
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
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectBeziers:
                            {
                                IBezier temp = (IBezier)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectCircle:
                            {
                                ICircle temp = (ICircle)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectConicCurve:
                            {
                                IConicCurve temp = (IConicCurve)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectDrawingContour:
                            {
                                IDrawingContour temp = (IDrawingContour)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEllipse:
                            {
                                IEllipse temp = (IEllipse)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEllipseArc:
                            {
                                IEllipseArc temp = (IEllipseArc)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectEquidistant:
                            {
                                IEquidistant temp = (IEquidistant)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectLineSegment:
                            {
                                ILineSegment temp = (ILineSegment)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectNurbs:
                            {
                                INurbs temp = (INurbs)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectPolyLine2D:
                            {
                                IPolyLine2D temp = (IPolyLine2D)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectRectangle:
                            {
                                IRectangle temp = (IRectangle)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectRegularPolygon:
                            {
                                IRegularPolygon temp = (IRegularPolygon)kompasobject;
                                if (temp.Style != 1)
                                {
                                    selectionManager.Unselect(temp);
                                }
                                break;
                            }
                        case Kompas6Constants.KompasAPIObjectTypeEnum.ksObjectMarkLeader: //Не снимает выделение если это обозначение маркировки
                            {
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

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)ActiveDocument;
            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView view = views.ActiveView;
            IDrawingContainer drawingContainer = (IDrawingContainer)view;
            ISymbols2DContainer symbols2DContainer = (ISymbols2DContainer)view;

            #region Проверка графических элементов на тип линии

            IArcs arcs = drawingContainer.Arcs;
            foreach (IArc arc in arcs)
            {
                if (arc.Style==1)
                {
                    selectobjects.Add(arc);
                }
            }
            IBeziers beziers = drawingContainer.Beziers;
            foreach (IBezier bezier in beziers)
            {
                if (bezier.Style==1)
                {
                    selectobjects.Add(bezier);
                }
            }
            ICircles circles = drawingContainer.Circles;
            foreach (ICircle circle in circles)
            {
                if (circle.Style==1)
                {
                    selectobjects.Add(circle);
                }
            }
            IConicCurves conicCurves = drawingContainer.ConicCurves;
            foreach (IConicCurve conicCurve in conicCurves)
            {
                if (conicCurve.Style==1)
                {
                    selectobjects.Add(conicCurve);
                }
            }
            IDrawingContours drawingContours = drawingContainer.DrawingContours;
            foreach (IDrawingContour drawingContour in drawingContours)
            {
                if (drawingContour.Style==1)
                {
                    selectobjects.Add(drawingContour);
                }
            }
            IEllipses ellipses = drawingContainer.Ellipses;
            foreach (IEllipse ellipse in ellipses)
            {
                if (ellipse.Style==1)
                {
                    selectobjects.Add(ellipse);
                }
            }
            IEllipseArcs ellipseNurbses = drawingContainer.EllipseArcs;
            foreach (IEllipseArc ellipseNurbse in ellipseNurbses)
            {
                if (ellipseNurbse.Style==1)
                {
                    selectobjects.Add(ellipseNurbse);
                }
            }
            IEquidistants equidistants = drawingContainer.Equidistants;
            foreach (IEquidistant equidistant in equidistants)
            {
                if (equidistant.Style==1)
                {
                    selectobjects.Add(equidistant);
                }
            }
            ILineSegments lineSegments = drawingContainer.LineSegments;
            foreach (ILineSegment lineSegment in lineSegments)
            {
                if (lineSegment.Style==1)
                {
                    selectobjects.Add(lineSegment);
                }
            }
            INurbses Nurbses = drawingContainer.Nurbses;
            foreach (INurbs nurbse in Nurbses)
            {
                if (nurbse.Style==1)
                {
                    selectobjects.Add(nurbse);
                }
            }
            INurbses NurbsesPoint = drawingContainer.NurbsesByPoints;
            foreach (INurbs nurbse in NurbsesPoint)
            {
                if (nurbse.Style==1)
                {
                    selectobjects.Add(nurbse);
                }
            }
            IPolyLines2D polyLines2Ds = drawingContainer.PolyLines2D;
            foreach (IPolyLine2D polyLines2D in polyLines2Ds)
            {
                if (polyLines2D.Style==1)
                {
                    selectobjects.Add(polyLines2D);
                }
            }
            IRectangles rectangles = drawingContainer.Rectangles;
            foreach (IRectangle rectangle in rectangles)
            {
                if (rectangle.Style==1)
                {
                    selectobjects.Add(rectangle);
                }
            }
            IRegularPolygons regularPolygons = drawingContainer.RegularPolygons;
            foreach (IRegularPolygon regularPolygon in regularPolygons)
            {
                if (regularPolygon.Style==1)
                {
                    selectobjects.Add(regularPolygon);
                }
            }

            #endregion

            #region Добавление в выбор обозначения маркировки
            foreach (var item in symbols2DContainer.Leaders)
            {
                if (item is IMarkLeader markLeader)
                {
                    selectobjects.Add(markLeader);
                }
            } 
            #endregion

            selectionManager.Select(selectobjects.ToArray());

            ksDocument2D.ksUndoContainer(false);
        }


        /// <summary>
        /// Открытие файла помощи
        /// </summary>
        private void OpenHelp()
        {
            ILibraryManager libraryManager = Application.LibraryManager;
            string path = $"{Path.GetDirectoryName(libraryManager.CurrentLibrary.PathName)}\\Help\\index.html"; //Получить путь к папке библиотеки
            if (File.Exists(path))
            {
                Process.Start(path);
            }
            else
            {
                Application.MessageBoxEx("Файл помощи не найден. Обратитесь к разработчику", "Ошибка", 64);
            }
        }
        #endregion

        // Головная функция библиотеки
        public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {
            Kompas = (KompasObject)kompas_;
            Application = (IApplication)Kompas.ksGetApplication7();
            ActiveDocument = Application.ActiveDocument;
            //Вызываем команды
            switch (command)
            {
                case 1: Select(new int[] {1}); break; //Выбор линии с типом "Основная"
                case 2: DestroyMacroElements(); break;
                case 3: SelecetTypeLine(); break;
                case 4: SelectTypeObjectCommand(); break;
                case 5: SelectTypeGraphic(); break;
                case 6: SelectTypeDimension(); break;
                case 7: SelectMainTypeLineMarkLeader(); break;



                case 999: OpenHelp(); break;
            }
        }
        /*
        public bool LibInterfaceNotifyEntry(object application1)
        {
            bool result = true;

            KompasEvents aplEvent = new KompasEvents(((KompasObject)application1));
            IConnectionPointContainer cpContainer = application1 as IConnectionPointContainer;
            IConnectionPoint m_ConnPt;
            int m_Cookie;
            if (cpContainer != null)
            {
                cpContainer.FindConnectionPoint(typeof(ksKompasObjectNotify).GUID, out m_ConnPt);
                if (m_ConnPt != null)
                {
                    m_ConnPt.Advise(aplEvent, out m_Cookie);
                }
            }
            return result;
        }
        */
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
