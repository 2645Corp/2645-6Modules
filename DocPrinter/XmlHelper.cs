using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace DocPrinter
{
    class XmlHelper
    {
        private Dictionary<string,string> bookmark = new Dictionary<string,string>();

        public Dictionary<string,string> Bookmark
        {
            get
            {
                return bookmark;
            }
        }

        public XmlHelper(string name)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("tpls\\" + name + ".xml");
            XmlNode root = xmlDoc.SelectSingleNode(xmlDoc.DocumentElement.Name);
            foreach (XmlNode bmk in root.SelectNodes("bookmark"))
            {
                bookmark[bmk.InnerText] = bmk.Attributes["data"].Value;
            }
        }

    }
}
