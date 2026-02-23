namespace X21.Utils
{
    public static class RibbonXmlUtils
    {
        public static string EscapeLabel(string text)
        {
            // Escape '&' which is interpreted as the shortcut key.
            // http://stackoverflow.com/questions/21333786/developing-a-ribbon-tab-in-word-2010-using-ampersand-symbol-in-group-label-name
            if (text.Contains(" & "))
            {
                text = text.Replace(" & ", " && ");
            }

            return text;
        }
    }
}
