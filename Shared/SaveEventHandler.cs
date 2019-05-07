using EnvDTE;
using EnvDTE80;
using System;
using System.Globalization;

namespace WhitespaceCleaner.Shared
{
    public sealed class SaveEventHandler
    {
        private DTE2 _applicationObject;
        private DocumentEvents _documentEvents;
        private string _whitespaceRegex;
        private bool _saved;

        public void OnConnection(DTE2 application)
        {
            _applicationObject = application;
            _documentEvents = _applicationObject.Events.DocumentEvents;
            _whitespaceRegex = ":Zs+$";
            if (double.TryParse(_applicationObject.Version, NumberStyles.Number, CultureInfo.InvariantCulture, out double version))
                if (version >= 11.0)
                    _whitespaceRegex = "[^\\S\\r\\n]+(?=\\r?$)";

            _saved = false;

            _documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }

        private void DocumentEvents_DocumentSaved(Document document)
        {
            if (!_saved)
            {
                try
                {
                    var props = _applicationObject.get_Properties("TextEditor", document.Language);
                    if (props == null)
                        props = _applicationObject.get_Properties("TextEditor", "AllLanguages");

                    var keepTabs = ((bool)props.Item("InsertTabs").Value);
                    if (!keepTabs)
                    {
                        var tabSize = (short)props.Item("TabSize").Value;

                        _applicationObject.Find.FindReplace(vsFindAction.vsFindActionReplaceAll, "\t",
                                             (int)vsFindOptions.vsFindOptionsRegularExpression,
                                             new string(' ', tabSize),
                                             vsFindTarget.vsFindTargetCurrentDocument, string.Empty, string.Empty,
                                             vsFindResultsLocation.vsFindResultsNone);
                    }

                    // Remove all the trailing whitespaces.
                    _applicationObject.Find.FindReplace(vsFindAction.vsFindActionReplaceAll,
                                         _whitespaceRegex,
                                         (int)vsFindOptions.vsFindOptionsRegularExpression,
                                         string.Empty,
                                         vsFindTarget.vsFindTargetCurrentDocument, string.Empty, string.Empty,
                                         vsFindResultsLocation.vsFindResultsNone);

                    _saved = true;
                    document.Save();
                }
                catch (Exception)
                {
                }
            }
            else
                _saved = false;
        }
    }
}
