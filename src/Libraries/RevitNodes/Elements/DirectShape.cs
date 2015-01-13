using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.DesignScript.Geometry;
using Autodesk.Revit.DB;
using DSNodeServices;
using Revit.GeometryConversion;

using RevitServices.Materials;
using RevitServices.Persistence;
using RevitServices.Transactions;

namespace Revit.Elements
{
    [RegisterForTrace]
    public class DirectShape : Element
    {
        [Browsable(false)]
        public override Autodesk.Revit.DB.Element InternalElement
        {
            get { return InternalDirectShape; }
        }

        public string Path { get; private set; }

        internal Autodesk.Revit.DB.DirectShape InternalDirectShape { get; private set; }
        internal Autodesk.Revit.DB.ImportInstance InternalImportInstance { get; private set; }

        internal string _directShapeGUID;
        internal static string _directShapeAppGUID = Guid.NewGuid().ToString();

        internal DirectShape(Geometry geom, Category category)
        {
            var goptions = new Options
            {
                IncludeNonVisibleObjects = true,
                ComputeReferences = true
            };

            // transform geometry from dynamo unit system (m) to revit (ft)
            var geometry = geom.InHostUnits();

            var translation = Vector.ByCoordinates(0, 0, 0);
            Robustify(ref geometry, ref translation);

            // Export to temporary file
            var fn = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".sat";
            var satPath = geometry.ExportToSAT(fn);

            var options = new SATImportOptions()
            {
                Unit = ImportUnit.Foot
            };

            TransactionManager.Instance.EnsureInTransaction(Document);

            var id = Document.Import(satPath, options, Document.ActiveView);
            var element = Document.GetElement(id);
            var importInstance = element as Autodesk.Revit.DB.ImportInstance;

            if (importInstance == null)
            {
                throw new Exception("Could not obtain ImportInstance from imported Element");
            }

            InternalSetImportInstance(importInstance);
            InternalUnpinAndTranslateImportInstance(translation.ToXyz());

            this.Path = satPath;

            GeometryElement revitGeometry = InternalImportInstance.get_Geometry(goptions);

            _directShapeGUID = Guid.NewGuid().ToString();

            Autodesk.Revit.DB.DirectShape ds = Autodesk.Revit.DB.DirectShape.CreateElement(
                Document, category.InternalCategory.Id, _directShapeAppGUID, _directShapeGUID);

            ds.SetShape(CollectConcreteGeometry(revitGeometry).ToList());

            InternalSetDirectShape(ds);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(importInstance);
            ElementBinder.SetElementForTrace(ds);
        }

        internal DirectShape(ImportInstance importInstance, Category category)
        {
            TransactionManager.Instance.EnsureInTransaction(Document);

            var goptions = new Options
            {
                IncludeNonVisibleObjects = true,
                ComputeReferences = true
            };

            GeometryElement revitGeometry = importInstance.InternalImportInstance.get_Geometry(goptions);

            _directShapeGUID = Guid.NewGuid().ToString();

            Autodesk.Revit.DB.DirectShape ds = Autodesk.Revit.DB.DirectShape.CreateElement(
                Document, category.InternalCategory.Id, _directShapeAppGUID, _directShapeGUID);

            ds.SetShape(CollectConcreteGeometry(revitGeometry).ToList());

            InternalSetDirectShape(ds);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(ds);
        }

        internal DirectShape(InvisibleImportInstance importInstance, Category category)
        {
            TransactionManager.Instance.EnsureInTransaction(Document);

            var goptions = new Options
            {
                IncludeNonVisibleObjects = true,
                ComputeReferences = true
            };

            GeometryElement revitGeometry = importInstance.InternalImportInstance.get_Geometry(goptions);

            _directShapeGUID = Guid.NewGuid().ToString();

            Autodesk.Revit.DB.DirectShape ds = Autodesk.Revit.DB.DirectShape.CreateElement(
                Document, category.InternalCategory.Id, _directShapeAppGUID, _directShapeGUID);

            ds.SetShape(CollectConcreteGeometry(revitGeometry).ToList());

            InternalSetDirectShape(ds);

            TransactionManager.Instance.TransactionTaskDone();

            ElementBinder.SetElementForTrace(ds);
        }

        public static DirectShape ByGeometry(Geometry geom, Category category)
        {
            return new DirectShape(geom, category);
        }

        public static DirectShape ByImportInstance(ImportInstance importInstance, Category category)
        {
            if (importInstance == null)
                throw new ArgumentNullException("importInstance");

            return new DirectShape(importInstance, category);
        }

        public static DirectShape ByInvisibleImportInstance(InvisibleImportInstance importInstance, Category category)
        {
            if (importInstance == null)
                throw new ArgumentNullException("importInstance");

            return new DirectShape(importInstance, category);
        }

        private void InternalUnpinAndTranslateImportInstance(Autodesk.Revit.DB.XYZ translation)
        {
            TransactionManager.Instance.EnsureInTransaction(Document);

            // the element must be unpinned to translate
            InternalImportInstance.Pinned = false;

            if (!translation.IsZeroLength()) ElementTransformUtils.MoveElement(Document, InternalImportInstance.Id, translation);

            TransactionManager.Instance.TransactionTaskDone();
        }

        private void InternalSetImportInstance(Autodesk.Revit.DB.ImportInstance ele)
        {
            //this.InternalUniqueId = ele.UniqueId;
            //this.InternalElementId = ele.Id;
            this.InternalImportInstance = ele;
        }

        private void InternalSetDirectShape(Autodesk.Revit.DB.DirectShape ds)
        {
            this.InternalUniqueId = ds.UniqueId;
            this.InternalElementId = ds.Id;
            this.InternalDirectShape = ds;
        }

        /// <summary>
        /// This method contains workarounds for increasing the robustness of input geometry
        /// </summary>
        /// <param name="geometry"></param>
        /// <param name="translation"></param>
        private static void Robustify(ref Autodesk.DesignScript.Geometry.Geometry geometry,
            ref Autodesk.DesignScript.Geometry.Vector translation)
        {
            // translate centroid of the solid to the origin
            // export, then move back 
            if (geometry is Autodesk.DesignScript.Geometry.Solid)
            {
                var solid = geometry as Autodesk.DesignScript.Geometry.Solid;

                translation = solid.Centroid().AsVector();
                var tranGeo = solid.Translate(translation.Reverse());

                geometry = tranGeo;
            }
        }

        /// <summary>
        /// This method contains workarounds for increasing the robustness of input geometry
        /// </summary>
        /// <param name="geometry"></param>
        /// <param name="translation"></param>
        private static void Robustify(ref Autodesk.DesignScript.Geometry.Geometry[] geometry,
            ref Autodesk.DesignScript.Geometry.Vector translation)
        {
            // translate all geom to centroid of bbox, then translate back
            var bb = Autodesk.DesignScript.Geometry.BoundingBox.ByGeometry(geometry);

            // get center of bbox
            var trans = ((bb.MinPoint.ToXyz() + bb.MaxPoint.ToXyz()) / 2).ToVector().Reverse();

            // translate all geom so that it is centered by bb
            geometry = geometry.Select(x => x.Translate(trans)).ToArray();

            // so that we can move it all back
            translation = trans.Reverse();
        }
    }
}
