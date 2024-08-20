using System.Text.RegularExpressions;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace TextReplaceAPI.DataTypes
{
    public partial class Styling
    {
        public bool Bold { get; set; }

        public bool Italics { get; set; }

        public bool Underline { get; set; }

        public bool Strikethrough { get; set; }

        public bool IsHighlighted { get; set; }

        public bool IsTextColored { get; set; }

        private string _highlightColor;
        public string HighlightColor
        {
            get { return _highlightColor; }
            set { _highlightColor = FormatColorString(value); }
        }

        private string _textColor;
        public string TextColor
        {
            get { return _textColor; }
            set { _textColor = FormatColorString(value); }
        }

        public Styling(
            bool bold = false,
            bool italics = false,
            bool underline = false,
            bool strikethrough = false,
            string highlightColor = "",
            string textColor = "")
        {
            Bold = bold;
            Italics = italics;
            Underline = underline;
            Strikethrough = strikethrough;
            IsHighlighted = (highlightColor != string.Empty);
            IsTextColored = (textColor != string.Empty);
            _highlightColor = FormatColorString(highlightColor);
            _textColor = FormatColorString(textColor);
        }

        /// <summary>
        /// For internal use only. Styles a .docx run.
        /// </summary>
        /// <param name="runProps"></param>
        /// <param name="style"></param>
        /// <returns>OpenXML RunProperties with styling applied</returns>
        public static Wordprocessing.RunProperties StyleRunProperties(
            Wordprocessing.RunProperties runProps, Styling style)
        {
            if (style.Bold)
            {
                runProps.Bold = new Wordprocessing.Bold();
            }

            if (style.Italics)
            {
                runProps.Italic = new Wordprocessing.Italic();
            }

            if (style.Underline)
            {
                runProps.Underline = new Wordprocessing.Underline()
                {
                    Val = Wordprocessing.UnderlineValues.Single
                };
            }

            if (style.Strikethrough)
            {
                runProps.Strike = new Wordprocessing.Strike();
            }

            if (style.IsHighlighted)
            {
                runProps.Shading = new Wordprocessing.Shading()
                {
                    Fill = style.HighlightColor,
                    Val = Wordprocessing.ShadingPatternValues.Clear,
                    Color = "auto"
                };
            }

            if (style.IsTextColored)
            {
                runProps.Color = new Wordprocessing.Color()
                {
                    Val = style.TextColor,
                    ThemeColor = Wordprocessing.ThemeColorValues.Accent1,
                    ThemeShade = "BF"
                };
            }

            return runProps;
        }

        /// <summary>
        /// For internal use only. Styles an .xlsx run.
        /// </summary>
        /// <param name="runProps"></param>
        /// <param name="style"></param>
        /// <returns>OpenXML RunProperties with styling applied</returns>
        public static Spreadsheet.RunProperties StyleRunProperties(
            Spreadsheet.RunProperties runProps, Styling style)
        {
            if (style.Bold)
            {
                runProps.AppendChild(new Spreadsheet.Bold());
            }

            if (style.Italics)
            {
                runProps.AppendChild(new Spreadsheet.Italic());
            }

            if (style.Underline)
            {
                runProps.AppendChild(new Spreadsheet.Underline()
                {
                    Val = Spreadsheet.UnderlineValues.Single
                });
            }

            if (style.Strikethrough)
            {
                runProps.AppendChild(new Spreadsheet.Strike());
            }

            // Excel spreadsheet highlighting has to be done on a cell rather than on a run,
            // so the method that constructs or edits the cells should set the highlighting value

            if (style.IsTextColored)
            {
                runProps.AppendChild(new Spreadsheet.Color()
                {
                    Rgb = "FF" + style.TextColor
                });
            }

            return runProps;
        }

        /// <summary>
        /// Format a 6-digit hex color code. Accepts strings with or without a '#', but removes the '#'.
        /// </summary>
        /// <param name="color"></param>
        /// <returns>The formatted 6-digit hex color code</returns>
        private static string FormatColorString(string color)
        {
            if (color == string.Empty)
            {
                return "000000";
            }

            // check if the string starts off with a pound (optional) and contains 6 hex digits
            if (HexStringRegex().Match(color).Success == false)
            {
                return "000000";
            }

            if (color[0] == '#')
            {
                color = color.Substring(1);
            }

            return color.ToUpper();
        }

        [GeneratedRegex("^#?[a-fA-F0-9]{6}$")]
        private static partial Regex HexStringRegex();
    }
}