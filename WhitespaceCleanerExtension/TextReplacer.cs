using EnvDTE;
using EnvDTE80;
using Shared;

namespace WhitespaceCleanerExtension
{
    public sealed class TextReplacer : ITextReplacer
    {
        private DTE2 _applicationObject;

        public TextReplacer(DTE2 applicationObject)
        {
            _applicationObject = applicationObject;
        }

        public void Replace(string whatRegex, string with, Document document)
        {
            _applicationObject.Find.FindReplace(vsFindAction.vsFindActionReplaceAll, whatRegex,
                                 (int)vsFindOptions.vsFindOptionsRegularExpression,
                                 with,
                                 vsFindTarget.vsFindTargetCurrentDocument, string.Empty, string.Empty,
                                 vsFindResultsLocation.vsFindResultsNone);
        }
    }
}
