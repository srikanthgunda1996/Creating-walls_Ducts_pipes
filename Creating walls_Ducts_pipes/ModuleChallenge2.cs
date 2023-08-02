#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace Creating_walls_Ducts_pipes
{
    [Transaction(TransactionMode.Manual)]
    public class ModuleChallenge2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

           UIDocument uidoc = uiapp.ActiveUIDocument;
           IList<Element> selectedElements = uidoc.Selection.PickElementsByRectangle("Select by rectangle");
            List<ModelCurve> modelcurves = new List<ModelCurve>();
            int count = 0;
            FilteredElementCollector lev = new FilteredElementCollector(doc).OfClass(typeof(Level));
            foreach(Element element in selectedElements)
            {
                if (element is CurveElement)
                {
                    CurveElement curve = (CurveElement)element;
                    if (curve.CurveElementType == CurveElementType.ModelCurve)
                    {
                        modelcurves.Add(curve as ModelCurve);
                        count++;
                    }
                }
            }
            Transaction t = new Transaction(doc);
            t.Start("create walls");
            FilteredElementCollector ducttype = new FilteredElementCollector(doc).OfClass(typeof (DuctType));
            FilteredElementCollector walltype = new FilteredElementCollector(doc).OfClass(typeof(WallType));
            FilteredElementCollector pipetype = new FilteredElementCollector(doc).OfClass(typeof(PipeType));
            FilteredElementCollector ductsystemtype = new FilteredElementCollector(doc).OfClass(typeof(MEPSystemType));
            FilteredElementCollector pipesystemtype = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType));
            WallType storefrontwall = null;
            WallType genwall = null;
           foreach(WallType genwall1 in walltype)
            {
                if (genwall1.Name == "Generic - 8\"")
                {
                    genwall = genwall1;
                    break;
                }
            }
            foreach (WallType genwall1 in walltype)
            {
                if (genwall1.Name == "Storefront")
                {
                    storefrontwall = genwall1;
                    break;
                }
            }

            foreach (CurveElement element in modelcurves)
            {
                Curve modelcurve = element.GeometryCurve;
                if (modelcurve.IsBound == true)
                {
                    GraphicsStyle curgra = element.LineStyle as GraphicsStyle;
                    XYZ startpoint = modelcurve.GetEndPoint(0);
                    XYZ endpoint = modelcurve.GetEndPoint(1);

                    if (curgra.Name == "A-GLAZ")
                    {
                        Wall.Create(doc, modelcurve, storefrontwall.Id, lev.FirstElementId(), 20, 0, false, false);
                    }
                    if (curgra.Name == "A-WALL")
                    {
                        Wall.Create(doc, modelcurve, genwall.Id, lev.FirstElementId(), 20, 0, false, false);
                    }
                    if (curgra.Name == "M-DUCT")
                    {
                        Duct.Create(doc, ductsystemtype.FirstElementId(), ducttype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                    }
                    if (curgra.Name == "P-PIPE")
                    {
                        Pipe.Create(doc, pipesystemtype.FirstElementId(), pipetype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                    }

                }

            }
            t.Commit();
            t.Dispose();
            TaskDialog.Show("Model Curves selected", count.ToString() + " " + "modelcurves got selected");
            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
