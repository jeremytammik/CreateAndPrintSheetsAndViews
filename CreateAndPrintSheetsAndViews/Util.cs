using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace CreateAndPrintSheetsAndViews
{
  class Util
  {
    #region Geometrical Comparison
    public const double _eps = 1.0e-9;

    public static double Eps
    {
      get
      {
        return _eps;
      }
    }

    public static double MinLineLength
    {
      get
      {
        return _eps;
      }
    }

    public static double TolPointOnPlane
    {
      get
      {
        return _eps;
      }
    }

    public static bool IsZero(
      double a,
      double tolerance = _eps)
    {
      return tolerance > Math.Abs(a);
    }

    public static bool IsEqual(
      double a,
      double b,
      double tolerance = _eps)
    {
      return IsZero(b - a, tolerance);
    }

    /// <summary>
    ///     Comparison method for two real numbers
    ///     returning 0 if they are to be considered equal,
    ///     -1 if the first is smaller and +1 otherwise
    /// </summary>
    public static int Compare(
        double a,
        double b,
        double tolerance = _eps)
    {
      return IsEqual(a, b, tolerance)
          ? 0
          : a < b
              ? -1
              : 1;
    }

    /// <summary>
    ///     Comparison method for two XYZ objects
    ///     returning 0 if they are to be considered equal,
    ///     -1 if the first is smaller and +1 otherwise
    /// </summary>
    public static int Compare(
        XYZ p,
        XYZ q,
        double tolerance = _eps)
    {
      var d = Compare(p.X, q.X, tolerance);

      if (0 == d)
      {
        d = Compare(p.Y, q.Y, tolerance);

        if (0 == d) d = Compare(p.Z, q.Z, tolerance);
      }

      return d;
    }

    /// <summary>
    ///     Predicate to test whether two points or
    ///     vectors can be considered equal with the
    ///     given tolerance.
    /// </summary>
    public static bool IsEqual(
        XYZ p,
        XYZ q,
        double tolerance = _eps)
    {
      return 0 == Compare(p, q, tolerance);
    }


    /// <summary>
    ///     Return true if the vectors v and w
    ///     are non-zero and perpendicular.
    /// </summary>
    private bool IsPerpendicular(XYZ v, XYZ w)
    {
      var a = v.GetLength();
      var b = v.GetLength();
      var c = Math.Abs(v.DotProduct(w));
      return _eps < a
             && _eps < b
             && _eps > c;
      // c * c < _eps * a * b
    }

    /// <summary>
    ///     Return true if the vectors p and Q are parallel, 
    ///     or at least one of them is zero length.
    /// </summary>
    public static bool IsParallel(XYZ p, XYZ q)
    {
      return p.CrossProduct(q).IsZeroLength();
    }

    /// <summary>
    ///     Predicate returning true if three given points are collinear
    /// </summary>
    public static bool AreCollinear(XYZ p, XYZ q, XYZ r)
    {
      var v = q - p;
      var w = r - p;
      return IsParallel(v, w);
    }
    #endregion // Geometrical Comparison

    #region Unit Handling

    private const double _inchToMm = 25.4;
    private const double _footToMm = 12 * _inchToMm;
    private const double _footToMeter = _footToMm * 0.001;
    private const double _sqfToSqm = _footToMeter * _footToMeter;
    private const double _cubicFootToCubicMeter = _footToMeter * _sqfToSqm;

    /// <summary>
    ///     Convert a given length in feet to millimetres.
    /// </summary>
    public static double FootToMm(double length)
    {
      return length * _footToMm;
    }

    /// <summary>
    ///     Convert a given UV vector in feet to millimetres.
    /// </summary>
    public static UV FootToMm(UV v)
    {
      return v * _footToMm;
    }

    /// <summary>
    ///     Convert a given XYZ vector in feet to millimetres.
    /// </summary>
    public static XYZ FootToMm(XYZ v)
    {
      return v * _footToMm;
    }

    /// <summary>
    ///     Round a real number to the closest int.
    /// </summary>
    public static int ToInt(double x)
    {
      return (int)Math.Round(x,
          MidpointRounding.AwayFromZero);
    }

    /// <summary>
    ///     Convert a given length in feet to millimetres,
    ///     rounded to the closest millimetre.
    /// </summary>
    public static int FootToMmInt(double length)
    {
      return ToInt(_footToMm * length);
    }
    #endregion // Unit Handling

    #region Formatting

    /// <summary>
    ///     Return an English plural suffix for the given
    ///     number of items, i.e. 's' for zero or more
    ///     than one, and nothing for exactly one.
    /// </summary>
    public static string PluralSuffix(int n)
    {
      return 1 == n ? "" : "s";
    }

    /// <summary>
    ///     Return an English plural suffix 'ies' or
    ///     'y' for the given number of items.
    /// </summary>
    public static string PluralSuffixY(int n)
    {
      return 1 == n ? "y" : "ies";
    }

    /// <summary>
    ///     Return a dot (full stop) for zero
    ///     or a colon for more than zero.
    /// </summary>
    public static string DotOrColon(int n)
    {
      return 0 < n ? ":" : ".";
    }

    /// <summary>
    ///     Return a string for a real number
    ///     formatted to two decimal places.
    /// </summary>
    public static string RealString(double a)
    {
      return a.ToString("0.##");
    }

    /// <summary>
    ///     Return a string representation in degrees
    ///     for an angle given in radians.
    /// </summary>
    public static string AngleString(double angle)
    {
      return $"{RealString(angle * 180 / Math.PI)} degrees";
    }

    /// <summary>
    ///     Return a string for a length in millimetres
    ///     formatted as an integer value.
    /// </summary>
    public static string MmString(double length)
    {
      //return RealString( FootToMm( length ) ) + " mm";
      return $"{Math.Round(FootToMm(length))} mm";
    }

    /// <summary>
    ///     Return a string for a UV point
    ///     or vector with its coordinates
    ///     formatted to two decimal places.
    /// </summary>
    public static string PointString(
        UV p,
        bool onlySpaceSeparator = false)
    {
      var format_string = onlySpaceSeparator
          ? "{0} {1}"
          : "({0},{1})";

      return string.Format(format_string,
          RealString(p.U),
          RealString(p.V));
    }

    /// <summary>
    ///     Return a string for an XYZ point
    ///     or vector with its coordinates
    ///     formatted to two decimal places.
    /// </summary>
    public static string PointString(
        XYZ p,
        bool onlySpaceSeparator = false)
    {
      var format_string = onlySpaceSeparator
          ? "{0} {1} {2}"
          : "({0},{1},{2})";

      return string.Format(format_string,
          RealString(p.X),
          RealString(p.Y),
          RealString(p.Z));
    }

    /// <summary>
    ///     Return a string in millimetres
    ///     for an XYZ point or vector in feet.
    /// </summary>
    public static string PointStringMm(
        XYZ p,
        bool onlySpaceSeparator = false)
    {
      var format_string = onlySpaceSeparator
          ? "{0} {1} {2}"
          : "({0},{1},{2})";

      return string.Format(format_string,
          FootToMmInt(p.X),
          FootToMmInt(p.Y),
          FootToMmInt(p.Z));
    }

    /// <summary>
    ///     Return a string for an XYZ point
    ///     or vector with its coordinates
    ///     formatted to zero decimal places.
    /// </summary>
    public static string PointStringInt(
        XYZ p,
        bool onlySpaceSeparator = false)
    {
      var format_string = onlySpaceSeparator
          ? "{0} {1} {2}"
          : "({0},{1},{2})";

      return string.Format(format_string,
          ToInt(p.X), ToInt(p.Y), ToInt(p.Z));
    }

    /// <summary>
    ///     Return a string describing the given element:
    ///     .NET type name,
    ///     category name,
    ///     family and symbol name for a family instance,
    ///     element id and element name.
    /// </summary>
    public static string ElementDescription(
        Element e)
    {
      if (null == e) return "<null>";

      // For a wall, the element name equals the
      // wall type name, which is equivalent to the
      // family name ...

      var fi = e as FamilyInstance;

      var typeName = e.GetType().Name;

      var categoryName = null == e.Category
          ? string.Empty
          : $"{e.Category.Name} ";

      var familyName = null == fi
          ? string.Empty
          : $"{fi.Symbol.Family.Name} ";

      var symbolName = null == fi
                       || e.Name.Equals(fi.Symbol.Name)
          ? string.Empty
          : $"{fi.Symbol.Name} ";

      //string alias = e is FabricationPart
      //  ? ((FabricationPart)e).Alias + " "
      //  : string.Empty;

      return $"{typeName} {categoryName}{familyName}{symbolName}<{e.Id.IntegerValue} {e.Name}>";
    }

    private const string _caption = "PartCamDimensioner";

    public static void InfoMsg2(
        string instruction,
        string content)
    {
      Debug.WriteLine($"{instruction}\r\n{content}");
      var d = new TaskDialog(_caption);
      d.MainInstruction = instruction;
      d.MainContent = content;
      d.Show();
    }

    public static void InfoMsg3(
        string instruction,
        IList<string> content)
    {
      string s = string.Join("\r\n", content);
      Debug.WriteLine($"{instruction}\r\n{s}");
      var d = new TaskDialog(_caption);
      d.MainInstruction = instruction;
      d.MainContent = s;
      d.Show();
    }

    public static bool AskYesNoQuestion(
        string question)
    {
      TaskDialog taskDialog = new TaskDialog("Please answer Yes or No");
      taskDialog.MainContent = question;
      TaskDialogCommonButtons buttons
        = TaskDialogCommonButtons.Yes
          | TaskDialogCommonButtons.No;
      taskDialog.CommonButtons = buttons;
      TaskDialogResult result = taskDialog.Show();
      return (result == TaskDialogResult.Yes);
    }
    #endregion // Formatting

    public static string GetProductCode(Element e)
    {
      const BuiltInParameter _bip_product_code
        = BuiltInParameter.FABRICATION_PRODUCT_CODE;
      Parameter p = e.get_Parameter(_bip_product_code);
      return (null != p)
        ? p.AsString()
        : null;
    }
  }
}
