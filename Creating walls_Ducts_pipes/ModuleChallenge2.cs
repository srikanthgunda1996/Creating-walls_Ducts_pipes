#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Windows.Forms;
using System.Drawing;
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
        internal Parameter GetParameterByName(Element Elem, string paramName)
        {
            foreach (Parameter Param in Elem.Parameters)
            {
                if (Param.Definition.Name.ToString() == paramName)
                    return Param;
            }

            return null;
        }
        internal bool SetParameterValue(Element curElem, string paramName, double value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);
            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;

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
            string pipesystyp = "";

            TaskDialog tdg = new TaskDialog("Duct Systems");
            tdg.MainInstruction = "Select Duct System:";
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Return Air");
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Exhaust Air");
            tdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Supply Air");
            TaskDialogResult selectedsystem = tdg.Show();
            if (selectedsystem == TaskDialogResult.CommandLink1)
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

            TaskDialog tsdg = new TaskDialog("Pipe Systems");
            tsdg.MainInstruction = "Select Pipe System:";
            tsdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Domestic Cold Water");
            tsdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Fire Protection Dry");
            tsdg.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Hydronic Supply");
            TaskDialogResult selectedpipesystem = tsdg.Show();
            if (selectedpipesystem == TaskDialogResult.CommandLink1)
            {
                pipesystyp = "Domestic Cold Water";
            }
            if (selectedpipesystem == TaskDialogResult.CommandLink2)
            {
                pipesystyp = "Fire Protection Dry";
            }
            if (selectedpipesystem == TaskDialogResult.CommandLink3)
            {
                pipesystyp = "Hydronic Supply";
            }


            FilteredElementCollector ducttype = new FilteredElementCollector(doc).OfClass(typeof (DuctType));
            FilteredElementCollector walltype = new FilteredElementCollector(doc).OfClass(typeof(WallType));
            FilteredElementCollector pipetype = new FilteredElementCollector(doc).OfClass(typeof(PipeType));
            MEPSystemType ductsystemtype = getductsystemtype(doc, ductsystyp);
            PipingSystemType pipesystemtype = getpipesystemtype(doc, pipesystyp);
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
            FilteredElementCollector dct = new FilteredElementCollector(doc).OfClass(typeof(Duct));
            int wid = 0;
            int hgt = 0;
            foreach (Element currentduct in dct)
            { 
                    TaskDialog tsdg2 = new TaskDialog("Duct Size");
                    tsdg2.MainInstruction = "Select Duct Size:";
                    tsdg2.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, @"1'x1'");
                    tsdg2.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, @"2'x2'");
                    tsdg2.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, @"3'x3'");
                    TaskDialogResult DuctSize = tsdg2.Show();
                    if (DuctSize == TaskDialogResult.CommandLink1)
                    {
                        wid = 1;
                        hgt = 1;
                    }
                    if (DuctSize == TaskDialogResult.CommandLink2)
                    {
                        wid = 2;
                        hgt = 2;
                }
                    if (DuctSize == TaskDialogResult.CommandLink3)
                    {
                        wid = 3;
                        hgt = 3;
                }
                bool ductwidth = SetParameterValue(currentduct, "Width", wid);
                bool ductheight = SetParameterValue(currentduct, "Height", hgt);
            }
            FilteredElementCollector pip = new FilteredElementCollector(doc).OfClass(typeof(Pipe));
            int dia = 0;
            foreach (Element currentpipe in pip)
            {
                TaskDialog tsdg3 = new TaskDialog("Pipe Size");
                tsdg3.MainInstruction = "Select Pipe Size:";
                tsdg3.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, @"1'");
                tsdg3.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, @"2'");
                tsdg3.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, @"3'");
                TaskDialogResult PipeSize = tsdg3.Show();
                if (PipeSize == TaskDialogResult.CommandLink1)
                {
                    dia = 1;
                }
                if (PipeSize == TaskDialogResult.CommandLink2)
                {
                    dia = 2;
                }
                if (PipeSize == TaskDialogResult.CommandLink3)
                {
                    dia = 3;
                }
                bool pipedia = SetParameterValue(currentpipe, "Diameter", dia);
            }


            t.Commit();
            t.Dispose();
            return Result.Succeeded;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
