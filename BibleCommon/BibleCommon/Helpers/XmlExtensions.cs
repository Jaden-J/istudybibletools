using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Xml;

namespace BibleCommon.Helpers
{
    public class CustomXPathFunctions : XsltContext
    {
        public CustomXPathFunctions()
        {
        }

        public CustomXPathFunctions(XmlNamespaceManager manager)            
        {
            foreach (string prefix in manager)
            {
                if (prefix != "xmlns")
                    this.AddNamespace(prefix, manager.LookupNamespace(prefix));
            } 
        }

        public override int CompareDocument(string baseUri, string nextbaseUri)
        {
            return 0;
        }

        public override bool PreserveWhitespace(XPathNavigator node)
        {
            return true;
        }

        public override IXsltContextFunction ResolveFunction(
            string prefix,
            string name,
            XPathResultType[] ArgTypes)
        {
            IXsltContextFunction resolvedFunction = null;
            if (name == EqualsFunction.FunctionName)
            {
                resolvedFunction = new EqualsFunction();
            }
            else if (name == ContainsFunction.FunctionName)
            {
                resolvedFunction = new ContainsFunction();
            }
            return resolvedFunction;
        }

        public override IXsltContextVariable ResolveVariable(string prefix, string name)
        {
            throw new NotImplementedException();
        }

        public override bool Whitespace
        {
            get { throw new NotImplementedException(); }
        }
    }

    public class ContainsFunction : IXsltContextFunction
    {
        public const string FunctionName = "contains";

        private XPathResultType[] m_argTypes;

        private string ExtractArgument(object arg)
        {
            string value = String.Empty;

            if (arg is string)
            {
                value = (string)(arg);
            }

            else if ((arg is XPathNodeIterator) && (((XPathNodeIterator)arg).MoveNext() == true))
            {
                value = ((XPathNodeIterator)arg).Current.ToString();
            }
            return value;
        }

        public object Invoke(XsltContext xsltContext, object[] args, System.Xml.XPath.XPathNavigator docContext)
        {
            if (args.Length != 2)
            {
                throw new ArgumentException("contains() takes two arguments");
            }

            string arg1 = ExtractArgument(args[0]);
            string arg2 = ExtractArgument(args[1]);

            return (0 <= arg1.IndexOf(arg2, StringComparison.OrdinalIgnoreCase));

        }

        #region IXsltContextFunction Members

        public XPathResultType[] ArgTypes
        {
            get
            {
                return m_argTypes;
            }
        }

        public int Maxargs
        {
            get { return 2; }
        }

        public int Minargs
        {
            get { return 2; }
        }

        public XPathResultType ReturnType
        {
            get { return XPathResultType.Boolean; }
        }

        #endregion
    }

    public class EqualsFunction : IXsltContextFunction
    {
        public const string FunctionName = "equals";

        private XPathResultType[] m_argTypes;
        //you can create array of return value like this
        //= new XPathResultType[] {XPathResultType.Any,XPathResultType.Any};

        private string ExtractArgument(object arg)
        {
            string value = String.Empty;

            if (arg is string)
            {
                value = (string)(arg);
            }

            else if ((arg is XPathNodeIterator) && (((XPathNodeIterator)arg).MoveNext() == true))
            {
                value = ((XPathNodeIterator)arg).Current.ToString();
            }
            return value;
        }

        public object Invoke(XsltContext xsltContext,
            object[] args,
            XPathNavigator docContext)
        {
            if (args.Length != 2)
            {
                throw new ArgumentException("equals() takes two arguments");
            }
            string arg1 = ExtractArgument(args[0]);
            string arg2 = ExtractArgument(args[1]);

            return String.Equals(arg1, arg2, StringComparison.OrdinalIgnoreCase);
        }

        #region IXsltContextFunction Members

        public XPathResultType[] ArgTypes
        {
            get
            {
                return m_argTypes;
            }
        }

        public int Maxargs
        {
            get { return 2; }
        }

        public int Minargs
        {
            get { return 2; }
        }

        public XPathResultType ReturnType
        {
            get { return XPathResultType.Boolean; }
        }

        #endregion
    }
}
