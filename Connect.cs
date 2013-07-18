using System;
using EnvDTE;
using EnvDTE80;
using Extensibility;

namespace WhitespaceCleaner
{
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    public class Connect : IDTExtensibility2
    {
        /// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
        public Connect()
        {
        }

        /// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param term='application'>Root object of the host application.</param>
        /// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param term='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _applicationObject = (DTE2)application;
            _addInInstance = (AddIn)addInInst;
            _documentEvents = _applicationObject.Events.DocumentEvents;
            _whitespaceRegex = ":Zs+$";
            double version;
            if (double.TryParse(_applicationObject.Version, out version))
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
                                             (int)EnvDTE.vsFindOptions.vsFindOptionsRegularExpression,
                                             new String(' ', tabSize),
                                             vsFindTarget.vsFindTargetCurrentDocument, String.Empty, String.Empty,
                                             vsFindResultsLocation.vsFindResultsNone);
                    }

                    // Remove all the trailing whitespaces.
                    _applicationObject.Find.FindReplace(vsFindAction.vsFindActionReplaceAll,
                                         _whitespaceRegex,
                                         (int)EnvDTE.vsFindOptions.vsFindOptionsRegularExpression,
                                         String.Empty,
                                         vsFindTarget.vsFindTargetCurrentDocument, String.Empty, String.Empty,
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

        /// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
        {
        }

        /// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref Array custom)
        {
        }

        /// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom)
        {
        }

        /// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
        /// <param term='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom)
        {
        }

        private DTE2 _applicationObject;
        private AddIn _addInInstance;
        private DocumentEvents _documentEvents;
        private string _whitespaceRegex;
        private bool _saved;
    }
}
