using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace CreateAndPrintSheetsAndViews
{
  /// <summary>
  /// Restrict user to select only fabrication part elements
  /// </summary>
  class FabricationPartSelectionFilter : ISelectionFilter
  {
    public bool AllowElement(Element e)
    {
      return e is FabricationPart;
    }

    public bool AllowReference(Reference r, XYZ p)
    {
      return true;
    }
  }
}
