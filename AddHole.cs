using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc==null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();
            if (familySymbol==null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

             List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();

            if(view3D==null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector refIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            using (var ts = new Transaction(arDoc, "Create hole for ducts"))
            {
                ts.Start();

                foreach (var duct in ducts)
                {
                    Line line = (duct.Location as LocationCurve).Curve as Line;
                    XYZ point = line.GetEndPoint(0);
                    XYZ direction = line.Direction;

                    var intersections = refIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= line.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();
                    foreach (var refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        hole.LookupParameter("Ширина").Set(duct.Diameter);
                        hole.LookupParameter("Высота").Set(duct.Diameter);
                    }
                }
                ts.Commit();
            }

            using (var ts = new Transaction(arDoc, "Create hole for pipes"))
            {
                ts.Start();

                foreach (var pipe in pipes)
                {
                    Line line = (pipe.Location as LocationCurve).Curve as Line;
                    XYZ point = line.GetEndPoint(0);
                    XYZ direction = line.Direction;

                    var intersections = refIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= line.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();
                    foreach (var refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                        hole.LookupParameter("Ширина").Set(pipe.Diameter);
                        hole.LookupParameter("Высота").Set(pipe.Diameter);
                    }
                }
                ts.Commit();
            }
            return Result.Succeeded;
        }

    }
    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
