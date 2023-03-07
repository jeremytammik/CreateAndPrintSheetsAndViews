#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
#endregion

namespace CreateAndPrintSheetsAndViews
{
  /// <summary>
  /// Create a sheet with four views, right side, 
  /// front, top and isometric 3D, for a selected 
  /// part, then print the sheet to PDF and JPG.
  /// </summary>
  [Transaction(TransactionMode.Manual)]
  public class CmdCreateAndPrintSheetAndViews 
    : IExternalCommand
  {
    const string _title_block_name = "AIR CRO - Plano DIN A4";
    const int _view_scale = 20;
    const int _view_scale_3d = 25;

    /// <summary>
    /// Create a TextNote at the specified XYZ point,
    /// useful during debugging to attach a label to 
    /// points in space.
    /// </summary>
    static TextNote CreateTextNote(
      View view,
      string text, 
      XYZ p,
      double rotation = 0)
    {
      Document doc = view.Document;

      var options = new TextNoteOptions
      {
        Rotation = rotation,
        HorizontalAlignment = HorizontalTextAlignment.Center,
        VerticalAlignment = VerticalTextAlignment.Middle,
        TypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
      };

      return TextNote.Create(doc, view.Id, p, text, options);
    }

#if SAMPLE_CODE_SNIPPET_CREATE_DIMENSION
    /// <summary>
    /// Create dimension sample from 
    /// https://www.revitapidocs.com/2023/47b3977d-da93-e1a4-8bfa-f23a29e5c4c1.htm
    /// </summary>
    Dimension CreateNewDimensionAlongLine(
      Document doc, Line line)
    {
      // Use the Start and End points of our line as the references  
      // Line must come from something in Revit, such as a beam
      ReferenceArray references = new ReferenceArray();
      references.Append(line.GetEndPointReference(0));
      references.Append(line.GetEndPointReference(1));

      // create the new dimension
      Dimension dimension = doc.Create.NewDimension(
        doc.ActiveView, line, references);
      return dimension;
    }
#endif // SAMPLE_CODE_SNIPPET_CREATE_DIMENSION

#if SAMPLE_CODE_SNIPPET_VIEW_CREATION
    void f(Document doc, ElementId assembly_instance_id )
    {
      // https://forums.autodesk.com/t5/revit-api-forum/create-views-dynamically/m-p/9301816
      // assembly views are also recommended by  RPTHOMAS108 in:
      // https://forums.autodesk.com/t5/revit-api-forum/printing-many-individual-elements-create-new-sheet-and-view-for/m-p/11799893#M69706

      using (Transaction tx = new Transaction(doc))
      {
        // Start "create view" transation
        tx.Start("View Creation");

        // Create views of assembly
        ElementId titleBlockId = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks).Cast<FamilySymbol>().FirstOrDefault().Id;

        // Create sheet
        ViewSheet viewSheet = AssemblyViewUtils.CreateSheet(doc, assembly_instance_id, titleBlockId);

        // Create isometric view
        View3D view3d = AssemblyViewUtils.Create3DOrthographic(doc, assembly_instance_id);
        view3d.DetailLevel = ViewDetailLevel.Fine;
        view3d.Scale = 10;

        // Create orthographic views //
        // Top Elevation
        ViewSection elevationTop = AssemblyViewUtils.CreateDetailSection(doc, assembly_instance_id, AssemblyDetailViewOrientation.ElevationTop);
        elevationTop.DetailLevel = ViewDetailLevel.Fine;
        elevationTop.Scale = 10;

        // Left Elevation
        ViewSection elevationLeft = AssemblyViewUtils.CreateDetailSection(doc, assembly_instance_id, AssemblyDetailViewOrientation.ElevationLeft);
        elevationLeft.DetailLevel = ViewDetailLevel.Fine;
        elevationLeft.Scale = 10;

        // Front Elevation
        ViewSection elevationFront = AssemblyViewUtils.CreateDetailSection(doc, assembly_instance_id, AssemblyDetailViewOrientation.ElevationFront);
        elevationFront.DetailLevel = ViewDetailLevel.Fine;
        elevationFront.Scale = 10;

        // Right Elevation
        ViewSection elevationRight = AssemblyViewUtils.CreateDetailSection(doc, assembly_instance_id, AssemblyDetailViewOrientation.ElevationRight);
        elevationRight.DetailLevel = ViewDetailLevel.Fine;
        elevationRight.Scale = 10;

        // Locate all views on sheet at once
        Viewport.Create(doc, viewSheet.Id, view3d.Id, new XYZ(1.75, 1.75, 0));
        Viewport.Create(doc, viewSheet.Id, elevationTop.Id, new XYZ(1, 1.5, 0));
        Viewport.Create(doc, viewSheet.Id, elevationLeft.Id, new XYZ(0.5, 1, 0));
        Viewport.Create(doc, viewSheet.Id, elevationFront.Id, new XYZ(1, 1, 0));
        Viewport.Create(doc, viewSheet.Id, elevationRight.Id, new XYZ(1.5, 1, 0));

        // Create material takeoff
        ViewSchedule materialTakeoff = AssemblyViewUtils.CreateMaterialTakeoff(doc, assembly_instance_id);
        ScheduleSheetInstance.Create(doc, viewSheet.Id, materialTakeoff.Id, new XYZ(2.25, 1.5, 0));

        // Create parts list (schedule)
        ViewSchedule partList = AssemblyViewUtils.CreatePartList(doc, assembly_instance_id);
        ScheduleSheetInstance.Create(doc, viewSheet.Id, partList.Id, new XYZ(2.25, 2, 0));

        // Close "create view" transation
        tx.Commit();
        tx.Commit();
      }
    }
#endif // SAMPLE_CODE_SNIPPET_VIEW_CREATION

    /// <summary>
    /// The angle in the XY plane (azimuth),
    /// typically 0 to 360. 
    /// Left Front = 45
    /// Front Right = 135
    /// Right Back = 225
    /// Back Left = 310
    /// </summary>
    const double angleHorizD = 135;

    /// <summary>
    /// The vertical tilt (altitude),
    /// typically -90 to 90.
    /// -30 = Top
    /// 30 = Bottom
    /// </summary>
    const double angleVertD = -30;

    /// <summary>
    /// Return a unit vector in the specified direction.
    /// </summary>
    /// <param name="angleHorizD">Angle in XY plane 
    /// in degrees</param>
    /// <param name="angleVertD">Vertical tilt between 
    /// -90 and +90 degrees</param>
    /// <returns>Unit vector in the specified 
    /// direction.</returns>
    static XYZ VectorFromHorizVertAngles(
      double angleHorizD,
      double angleVertD)
    {
      // Convert degreess to radians.

      double degToRadian = Math.PI * 2 / 360;
      double angleHorizR = angleHorizD * degToRadian;
      double angleVertR = angleVertD * degToRadian;

      // Return unit vector in 3D

      double a = Math.Cos(angleVertR);
      double b = Math.Cos(angleHorizR);
      double c = Math.Sin(angleHorizR);
      double d = Math.Sin(angleVertR);

      return new XYZ(a * b, a * c, d);
    }

    /// <summary>
    /// Create and return a 3D view.
    /// </summary>
    static View3D CreateView3d(Document doc)
    {
      ViewFamilyType viewFamilyType
        = new FilteredElementCollector(doc)
          .OfClass(typeof(ViewFamilyType))
          .Cast<ViewFamilyType>()
          .Where(v => ViewFamily.ThreeDimensional == v.ViewFamily)
          .First();

      View3D view = View3D.CreateIsometric(doc, viewFamilyType.Id);

      XYZ eye = XYZ.Zero;

      XYZ forward = VectorFromHorizVertAngles(
        angleHorizD, angleVertD);

      XYZ up = VectorFromHorizVertAngles(
        angleHorizD, angleVertD + 90);

      ViewOrientation3D viewOrientation3D
        = new ViewOrientation3D(eye, up, forward);

      view.SetOrientation(viewOrientation3D);
      view.SaveOrientation();
      view.DetailLevel = ViewDetailLevel.Fine;
      view.Scale = _view_scale_3d;

      return view;
    }

    /// <summary>
    /// Create and return a section view focused 
    /// on the given target point and zoomed to 
    /// the given size with given right and up 
    /// directions.
    /// </summary>
    static ViewSection CreateViewSection(
      Document doc, 
      XYZ pOrigin, 
      double halfsize,
      XYZ vRight, 
      XYZ vUp,
      ref List<ElementId> idsToShow )
    {
      // Find a section view type

      ViewFamilyType viewFamilyType
        = new FilteredElementCollector(doc)
          .OfClass(typeof(ViewFamilyType))
          .Cast<ViewFamilyType>()
          .Where(v => ViewFamily.Section == v.ViewFamily)
          .First();

      // Create a BoundingBoxXYZ instance centered on wall
      //LocationCurve lc = wall.Location as LocationCurve;
      //Transform curveTransform = lc.Curve.ComputeDerivatives(0.5, true);
      // using 0.5 and "true" (to specify that the parameter is normalized) 
      // places the transform's origin at the center of the location curve)
      //XYZ origin = curveTransform.Origin; // mid-point of location curve
      //XYZ viewDirection = curveTransform.BasisX.Normalize(); // tangent vector along the location curve
      //XYZ normal = viewDirection.CrossProduct(XYZ.BasisZ).Normalize(); // location curve normal @ mid-point
      // can use this simplification because wall's "up" is vertical.
      // For a non-vertical situation (such as section through a sloped floor the surface normal would be needed)
      //transform.BasisZ = normal.CrossProduct(XYZ.BasisZ);
      //sectionBox.Min = new XYZ(-10, 0, 0);
      //sectionBox.Max = new XYZ(10, 12, 5);
      // Min & Max X values (-10 & 10) define the section line length on each side of the wall
      // Max Y (12) is the height of the section box// Max Z (5) is the far clip offset

      //XYZ p_sample_part = new XYZ(38, -47, 0);
      //double d_size = 4;

      XYZ vSize = new XYZ(halfsize, halfsize, halfsize);

      Transform transform = Transform.Identity;

      transform.Origin = pOrigin; // XYZ.Zero;
      transform.BasisX = vUp.CrossProduct(vRight);
      transform.BasisY = vUp;
      transform.BasisZ = vRight; 
      Debug.Assert(Util.IsEqual(1, transform.Determinant), "expected 1 determinant");

      BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
      sectionBox.Transform = transform;
      sectionBox.Min = -vSize;
      sectionBox.Max = vSize;

      ViewSection viewSection = ViewSection.CreateSection(
        doc, viewFamilyType.Id, sectionBox);

      viewSection.DetailLevel = ViewDetailLevel.Fine;
      viewSection.Scale = _view_scale;

      double f = 0.2 * halfsize;
      double rotation = vRight.AngleOnPlaneTo(vUp, transform.BasisZ); // this is not right
      idsToShow.Add(CreateTextNote(viewSection, "O", pOrigin).Id);
      idsToShow.Add(CreateTextNote(viewSection, "R", pOrigin + f * vRight).Id);
      idsToShow.Add(CreateTextNote(viewSection, "U", pOrigin + f * vUp).Id);
      
      return viewSection;
    }

    /// <summary>
    /// Return the primary connector 
    /// from the given connector set.
    /// </summary>
    static Connector GetPrimaryConnector(ConnectorSet conset)
    {
      Connector primary_connector = null;
      foreach (Connector c in conset)
      {
        MEPConnectorInfo info = c.GetMEPConnectorInfo();
        if (info.IsPrimary)
        {
          primary_connector = c;
          break;
        }
      }
      return primary_connector;
    }

    /// <summary>
    /// Return LCS of a duct, its local coordinate
    /// system, defined by the primary connector.
    /// </summary>
    static Transform GetDuctLcs(FabricationPart part)
    {
      ConnectorManager conmgr = part.ConnectorManager;
      ConnectorSet conset = conmgr.Connectors;
      Connector start = GetPrimaryConnector(conset);
      // Transform from local duct to world coordinate system
      Transform twcs = start.CoordinateSystem;
      Debug.Assert(Util.IsEqual(1, twcs.Determinant), "expected 1 twcs determinant");
      // Flip so that Z axis points into duct, not out of it
      twcs.BasisY = -(twcs.BasisY);
      twcs.BasisZ = -(twcs.BasisZ);
      Debug.Assert(Util.IsEqual(1, twcs.Determinant), "expected 1 flipped twcs determinant");
      return twcs;
    }

    /// <summary>
    /// Return midpoint between two points
    /// </summary>
    static XYZ MidPoint(XYZ a, XYZ b)
    {
      return a + 0.5 * (b - a);
    }

    /// <summary>
    /// Ensure that only the given element 
    /// is visible in view. 
    /// IsolateElementTemporary works fine,
    /// but saving to PDF displays the dialogue
    /// 'Export with Temporary Hide/Isolate'.
    /// Maybe use HideElements (all visible in 
    /// view) + UnhideElements (single) instead?
    /// Nope, that seems pretty tricky.
    /// Maybe this can be handled using the 
    /// DialogBoxShowing event or the Failure API?
    /// Yes, cf. https://thebuildingcoder.typepad.com/blog/2013/03/export-wall-parts-individually-to-dxf.html#3
    /// </summary>
    static void IsolateElementInView(IList<ElementId> idsToShow, View v)
    {

#if WORK_AROUND_Export_with_Temporary_Hide_Isolate_DIALOGUE
      Document doc = e.Document;
      Options geoopt = new Options();

      //Application app = doc.Application;
      // Zoom to using ShowElements

      List<ElementId> idsThisElement 
        = new List<ElementId>(1) { e.Id };

      v.IsolateElementTemporary(e.Id); // zoom to relevant area

      List<ElementId> idsAllElements // ICollection
        = new List<ElementId>( new FilteredElementCollector(doc, v.Id)
          //.WhereElementIsNotElementType()
          //.WhereElementIsViewIndependent()
          //.IsViewValidForElementIteration(doc,v.Id)
          //.OfClass(typeof(FamilyInstance))
          //.OfCategory(BuiltInCategory.OST_DuctFitting)
          //.ToElementIds();
          .Where<Element>(x
            => ((null != x.Category)
              && (null != x.LevelId)
              && (null != x.get_Geometry(geoopt))
              && (x.Id.IntegerValue > v.Id.IntegerValue)))
          .Select<Element,ElementId>(x=>x.Id));

      // Autodesk.Revit.Exceptions.ArgumentException:
      // One of the elements cannot be hidden.
      // Parameter name: elementIdSet'

      v.HideElements(idsAllElements);
      v.UnhideElements(idsThisElement);
#endif // WORK_AROUND_Export_with_Temporary_Hide_Isolate_DIALOGUE

      v.IsolateElementsTemporary(idsToShow);
    }

    /// <summary>
    /// Create a sheet and four views for the given 
    /// element: right, front, top and 3d, isolated
    /// and zoomed. For fabrication parts, orient 
    /// the views according to the duct LCS.
    /// </summary>
    public static void CreateSheetAndViewsFor(Element e)
    {
      Document doc = e.Document;
      BoundingBoxXYZ bb = e.get_BoundingBox(null);
      XYZ p = MidPoint(bb.Min, bb.Max);
      double halfsize = 0.5 * (bb.Max - bb.Min).GetLength();

      // Orient views according to part LCS

      FabricationPart part = e as FabricationPart;
      Transform twcs = (null != part)
        ? GetDuctLcs(part)
        : Transform.Identity;

      List<ElementId> idsToShowFront = new List<ElementId>() { e.Id };
      List<ElementId> idsToShowRight = new List<ElementId>() { e.Id };
      List<ElementId> idsToShowTop = new List<ElementId>() { e.Id };

      View view3d = CreateView3d(doc); // todo: adapt orientation to duct LCS 
      View viewFront = CreateViewSection(doc, p, halfsize, twcs.BasisY, twcs.BasisZ, ref idsToShowFront);
      View viewRight = CreateViewSection(doc, p, halfsize, twcs.BasisX, twcs.BasisZ, ref idsToShowRight);
      View viewTop = CreateViewSection(doc, p, halfsize, twcs.BasisX, twcs.BasisY, ref idsToShowTop);

      view3d.Name = "3D"; // Autodesk.Revit.Exceptions.ArgumentException: 'Name must be unique. Parameter name: name'
      viewFront.Name = "Front"; // PartCamDimensionerView
      viewRight.Name = "Right";
      viewTop.Name = "Top";

      IsolateElementInView(new List<ElementId>(1) { e.Id }, view3d);
      IsolateElementInView(idsToShowFront, viewFront);
      IsolateElementInView(idsToShowRight, viewRight);
      IsolateElementInView(idsToShowTop, viewTop);

      // Get title block

      FamilySymbol titleBlock
        = new FilteredElementCollector(doc)
          .OfClass(typeof(FamilySymbol))
          .OfCategory(BuiltInCategory.OST_TitleBlocks)
          .Where<Element>(x => x.Name.Equals(_title_block_name))
          .First() as FamilySymbol;

      if (null != titleBlock)
      {
        ViewSheet viewSheet = ViewSheet.Create(doc, titleBlock.Id);
        if (null != viewSheet)
        {
          string s = Util.GetProductCode(e);
          if (null == s)
          {
            s = string.Empty;
          }
          else
          {
            s += "-";
          }

          viewSheet.SheetNumber = "00";
          viewSheet.Name = $"Part CAM Dimension Sheet for {s}{e.Id.IntegerValue}";

          // Locate views on sheet: right front top 3d

          UV pmin = viewSheet.Outline.Min;
          UV pmax = viewSheet.Outline.Max;
          UV v = pmax - pmin;
          double w = v.U;
          double h = v.V;
          // adjust size to leave space for title sheet fields
          double left = pmin.U + 0.1 * w;
          double bottom = pmin.V + 0.15 * h;
          w *= 0.9;
          h *= 0.8;
          XYZ pul = new XYZ(left + 0.25 * w, bottom + 0.75 * h, 0);
          XYZ pur = new XYZ(left + 0.75 * w, bottom + 0.75 * h, 0);
          XYZ pll = new XYZ(left + 0.25 * w, bottom + 0.3 * h, 0);
          XYZ plr = new XYZ(left + 0.7 * w, bottom + 0.3 * h, 0);

          Viewport vpr = Viewport.Create(doc, viewSheet.Id, viewRight.Id, pul);
          Viewport vpf = Viewport.Create(doc, viewSheet.Id, viewFront.Id, pur);
          Viewport vpt = Viewport.Create(doc, viewSheet.Id, viewTop.Id, pll);
          Viewport vp3 = Viewport.Create(doc, viewSheet.Id, view3d.Id, plr);

          // Autodesk.Revit.Exceptions.InvalidOperationException:
          // This element does not support assignment of a user-specified name.
          //vpr.Name = "Right"; 
          //vpf.Name = "Front";
          //vpt.Name = "Top";
          //vp3.Name = "3D";

          // Create part type specific dimensioning in the right and front views

          // doc.Create.NewDimension( doc.ActiveView, line, refArray );

          // Print the sheet

          if (viewSheet.CanBePrinted
            /*&& Util.AskYesNoQuestion("Print the sheet?")*/ )
          {
            IList<ElementId> viewIds = new List<ElementId>(1) { 
              viewSheet.Id };

            string dir = "C:/tmp";
            string project_name = doc.Title;
            string path = dir + "/" + project_name; // + filename;

            bool _export_pdf = true;
            if (_export_pdf)
            {
              // Export PDF

              PDFExportOptions opt = new PDFExportOptions();
              opt.FileName = viewSheet.Name;
              doc.Export(dir, viewIds, opt);
            }

            // Export image

            ImageExportOptions imgopt = new ImageExportOptions();
            imgopt.ExportRange = ExportRange.SetOfViews;
            imgopt.SetViewsAndSheets(viewIds);
            imgopt.FilePath = path;

            doc.ExportImage(imgopt);
          }
        }
      }
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      uiapp.DialogBoxShowing
        += new EventHandler<DialogBoxShowingEventArgs>(
          Command.OnDialogBoxShowing);

      using (Transaction t = new Transaction(doc))
      {
        t.Start("Create sheet and four views");
        ElementId id_sample_element = new ElementId(14974273); // 70TG '1100 mmx600 mm-600 mmx600 mm-700 mmx600 mm' at (38.0286, -47.409, 0)
        Element e = doc.GetElement(id_sample_element);
        CreateSheetAndViewsFor(e);
        bool save = true; // Util.AskYesNoQuestion("Save the sheet?");
        if (save) 
        { 
          t.Commit();
        }
        else
        {
          t.RollBack();
        }
      }
      return Result.Succeeded;
    }
  }
}
