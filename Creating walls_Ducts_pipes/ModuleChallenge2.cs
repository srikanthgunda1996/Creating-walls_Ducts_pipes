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
        internal MEPSystemType getductsystemtype(Document doc, string typeName)
        {
            FilteredElementCollector ducttype = new FilteredElementCollector(doc).OfClass(typeof(MEPSystemType));
            MEPSystemType Ductsystem = null;
            foreach(MEPSystemType ductsystem in ducttype)
            {
                if(ductsystem.Name == typeName)
                {
                    Ductsystem = ductsystem;
                    break;
                }
            }
            return Ductsystem;
        }
        internal PipingSystemType getpipesystemtype(Document doc, string typeName)
        {
            FilteredElementCollector pipesystemtype = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType));
            PipingSystemType PipingSystem = null;
            foreach (PipingSystemType item in pipesystemtype)
            {
                if(item.Name == typeName)
                {
                    PipingSystem = item;
                    break;
                }
            }
            return PipingSystem;
        }

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
            string ductsystyp = "";
            TaskDialog tdg = new TaskDialog("Duct Systems");
            tdg.MainInstruction = "Select Duct System:";
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Return Air");
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Exhaust Air");
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Supply Air");
            TaskDialogResult selectedsystem = tdg.Show();
            if(selectedsystem == TaskDialogResult.CommandLink1)
            {
                ductsystyp = "Return Air";
            }
            if (selectedsystem == TaskDialogResult.CommandLink2)
            {
                ductsystyp = "Exhaust Air";
            }
            if (selectedsystem == TaskDialogResult.CommandLink3)
            {
                ductsystyp = "Supply Air";
            }
            FilteredElementCollector ducttype = new FilteredElementCollector(doc).OfClass(typeof (DuctType));
            FilteredElementCollector walltype = new FilteredElementCollector(doc).OfClass(typeof(WallType));
            FilteredElementCollector pipetype = new FilteredElementCollector(doc).OfClass(typeof(PipeType));
            MEPSystemType ductsystemtype = getductsystemtype(doc, ductsystyp);
            PipingSystemType pipesystemtype = getpipesystemtype(doc, "Domestic Cold Water");
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
                    

                    //if (curgra.Name == "A-GLAZ")
                    //{
                    //    Wall.Create(doc, modelcurve, storefrontwall.Id, lev.FirstElementId(), 20, 0, false, false);
                    //}
                    //if (curgra.Name == "A-WALL")
                    //{
                    //    Wall.Create(doc, modelcurve, genwall.Id, lev.FirstElementId(), 20, 0, false, false);
                    //}
                    //if (curgra.Name == "M-DUCT")
                    //{
                    //    Duct.Create(doc, ductsystemtype.Id, ducttype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                    //}
                    //if (curgra.Name == "P-PIPE")
                    //{
                    //    Pipe.Create(doc, pipesystemtype.Id, pipetype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                    //}
                    switch (curgra.Name)
                    {
                        case "A-GLAZ":
                            Wall.Create(doc, modelcurve, storefrontwall.Id, lev.FirstElementId(), 20, 0, false, false);
                            break;
                        case "A-WALL":
                            Wall.Create(doc, modelcurve, genwall.Id, lev.FirstElementId(), 20, 0, false, false);
                            break;
                        case "M-DUCT":
                            Duct.Create(doc, ductsystemtype.Id, ducttype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                            break;
                        case "P-PIPE":
                            Pipe.Create(doc, pipesystemtype.Id, pipetype.FirstElementId(), lev.FirstElementId(), startpoint, endpoint);
                            break;
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
