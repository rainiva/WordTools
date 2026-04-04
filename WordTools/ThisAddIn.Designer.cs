using Word = Microsoft.Office.Interop.Word;

namespace WordTools
{
    public partial class ThisAddIn
    {
        private Word.Application application;

        internal Word.Application Application
        {
            get
            {
                return application;
            }
            set
            {
                application = value;
            }
        }
    }

    internal sealed partial class Globals
    {
        private static ThisAddIn _ThisAddIn;
        private static Word.Application _Application;

        internal static ThisAddIn ThisAddIn
        {
            get
            {
                return _ThisAddIn;
            }
            set
            {
                if ((_ThisAddIn == null))
                {
                    _ThisAddIn = value;
                }
                else
                {
                    throw new System.NotSupportedException();
                }
            }
        }

        internal static Word.Application Application
        {
            get
            {
                return _Application;
            }
            set
            {
                if ((_Application == null))
                {
                    _Application = value;
                }
                else
                {
                    throw new System.NotSupportedException();
                }
            }
        }
    }
}
