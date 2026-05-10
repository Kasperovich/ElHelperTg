using System.Drawing;
using ClosedXML.Excel;

/// <summary>Определение признака «выходной» по заливке шапки дня в графике (например, красная заливка).</summary>
internal static class ExcelDutyDayFillAnalyzer
{
    /// <returns>true, если заливка похожа на «выходной» акцент (красноватый фон).</returns>
    internal static bool CellLooksLikeRestDayFill(IXLCell cell)
    {
        try
        {
            var fill = cell.Style.Fill;

            return fill.PatternType switch
            {
                XLFillPatternValues.None or XLFillPatternValues.LightGray =>
                    EvalPatternLineColor(fill),
                XLFillPatternValues.Solid => EvalSolidBackground(fill.BackgroundColor),
                _ => EvalSolidBackground(fill.BackgroundColor),
            };
        }
        catch
        {
            return false;
        }
    }

    private static bool EvalPatternLineColor(IXLFill fill)
    {
        var xl = fill.PatternColor;
        if (xl.Equals(XLColor.NoColor))
            return EvalSolidBackground(fill.BackgroundColor);

        return EvalSolidBackground(xl);
    }

    private static bool EvalSolidBackground(XLColor xl)
    {
        if (xl.Equals(XLColor.NoColor))
            return false;

        try
        {
            if (xl.ColorType != XLColorType.Color)
                return false;

            return IsAccentRedRgb(xl.Color);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Эвристика «выделение красным»: доминирует красный канал над зелёным и синим.</summary>
    private static bool IsAccentRedRgb(Color c)
    {
        if (c.A < 64)
            return false;

        var r = c.R;
        var g = c.G;
        var b = c.B;

        return r >= 160 && r >= g + 55 && r >= b + 40;
    }
}
