using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using BibleCommon.Consts;
using System.Xml.XPath;

namespace BibleCommon.Services
{
    public static class QuickStyleManager
    {
        public static readonly string StyleForStrongName = "strongsDictionary";
        public static readonly string StyleNameH2 = "h2";
        public static readonly string StyleNameH3 = "h3";        

        public enum PredefinedStyles
        {
            GrayHyperlink,
            H2,
            H3
        }        

        public static int AddQuickStyleDef(XDocument pageDoc, string styleName, PredefinedStyles styleType, XmlNamespaceManager xnm)
        {
            //"  <one:QuickStyleDef index="0" name="p" fontColor="#7F7F7F" highlightColor="automatic" font="Calibri" fontSize="11.0" spaceBefore="0.0" spaceAfter="0.0" />";
            //"  <one:QuickStyleDef index="1" name="h2" fontColor="#366092" highlightColor="automatic" font="Calibri" fontSize="13.0" bold="true" spaceBefore="0.0" spaceAfter="0.0" />";
            //"  <one:QuickStyleDef index="1" name="h3" fontColor="#366092" highlightColor="automatic" font="Calibri" fontSize="11.0" bold="true" spaceBefore="0.0" spaceAfter="0.0" />";
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            int maxIndex = -1;
            XElement latestStyleEl = null;
            foreach (var styleEl in pageDoc.Root.XPathSelectElements("one:QuickStyleDef", xnm))
            {
                var index = int.Parse((string)styleEl.Attribute("index"));

                if ((string)styleEl.Attribute("name") == styleName)
                    return index;

                if (index > maxIndex)
                    maxIndex = index;

                latestStyleEl = styleEl;
            }

            maxIndex++;
            var newStyleEl = new XElement(nms + "QuickStyleDef",
                                new XAttribute("index", maxIndex),
                                new XAttribute("name", styleName));

            ApplyStyleAttributes(newStyleEl, styleType);

            if (latestStyleEl == null)
                pageDoc.Root.AddFirst(newStyleEl);
            else
                latestStyleEl.AddAfterSelf(newStyleEl);

            return maxIndex;

        }

        public static void SetQuickStyleDefForCell(XElement cell, int styleIndex, XmlNamespaceManager xnm)
        {
            cell.XPathSelectElement("one:OEChildren/one:OE", xnm).SetAttributeValue("quickStyleIndex", styleIndex);
        }

        private static void ApplyStyleAttributes(XElement styleEl, PredefinedStyles styleType)
        {
            styleEl.Add(
                new XAttribute("highlightColor", "automatic"),
                new XAttribute("font", "Calibri"),                                
                new XAttribute("spaceBefore", "0.0"),
                new XAttribute("spaceAfter", "0.0"));

            switch (styleType)
            {
                case PredefinedStyles.GrayHyperlink:
                    styleEl.Add(
                        new XAttribute("fontColor", "#7F7F7F"),
                        new XAttribute("fontSize", "11.0"));
                    break;
                case PredefinedStyles.H2:
                    styleEl.Add(
                        new XAttribute("fontColor", "#366092"),
                        new XAttribute("fontSize", "13.0"),
                        new XAttribute("bold", "true"));
                    break;
                case PredefinedStyles.H3:
                    styleEl.Add(
                        new XAttribute("fontColor", "#366092"),
                        new XAttribute("fontSize", "11.0"),
                        new XAttribute("bold", "true"));
                    break;
            }
        }
    }
}
