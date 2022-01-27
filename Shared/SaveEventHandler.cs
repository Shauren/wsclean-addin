using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Shared;
using System;
using System.Globalization;

namespace WhitespaceCleaner.Shared
{
    public sealed class SaveEventHandler
    {
        private DTE2 _applicationObject;
        private DocumentEvents _documentEvents;
        private string _whitespaceRegex;
        private ITextReplacer _textReplacer;
        private bool _saved;

        public void OnConnection(DTE2 application, ITextReplacer textReplacer)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _applicationObject = application;
            _documentEvents = _applicationObject.Events.DocumentEvents;
            _whitespaceRegex = ":Zs+$";
            if (double.TryParse(_applicationObject.Version, NumberStyles.Number, CultureInfo.InvariantCulture, out double version))
                if (version >= 11.0)
                    _whitespaceRegex = "[^\\S\\r\\n]+(?=\\r?$)";

            _textReplacer = textReplacer;
            _saved = false;

            _documentEvents.DocumentSaved += DocumentEvents_DocumentSaved;
        }

        private void DocumentEvents_DocumentSaved(Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

                        _textReplacer.Replace("\t", new string(' ', tabSize), document);
                    }

                    // Remove all the trailing whitespaces.
                    _textReplacer.Replace(_whitespaceRegex, string.Empty, document);

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
