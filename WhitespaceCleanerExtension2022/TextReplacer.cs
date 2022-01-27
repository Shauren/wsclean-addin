using EnvDTE;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
using Shared;
using System;
using WhitespaceCleanerExtension;

namespace WhitespaceCleanerExtension2022
{
    public sealed class TextReplacer : ITextReplacer
    {
        public void Replace(string whatRegex, string with, Document document)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (TryGetTextBufferAt(document.FullName, out var textBuffer))
            {
                var findService = SaveCommandPackage.Instance.ComponentModel.GetService<IFindService>();

                var finderFactory = findService.CreateFinderFactory(whatRegex, with, FindOptions.UseRegularExpressions);
                var finder = finderFactory.Create(textBuffer.CurrentSnapshot);

                using (var edit = textBuffer.CreateEdit())
                {
                    foreach (var match in finder.FindForReplaceAll())
                        edit.Replace(match.Match, match.Replace);

                    edit.Apply();
                }
            }
        }

        private static bool TryGetTextBufferAt(string filePath, out ITextBuffer textBuffer)
        {
            if (VsShellUtilities.IsDocumentOpen(SaveCommandPackage.Instance, filePath, Guid.Empty, out var _, out var _, out var windowFrame))
            {
                IVsTextView view = VsShellUtilities.GetTextView(windowFrame);
                if (view.GetBuffer(out var lines) == 0)
                {
                    if (lines is IVsTextBuffer buffer)
                    {
                        var editorAdapterFactoryService = SaveCommandPackage.Instance.ComponentModel.GetService<IVsEditorAdaptersFactoryService>();
                        textBuffer = editorAdapterFactoryService.GetDataBuffer(buffer);
                        return true;
                    }
                }
            }

            textBuffer = null;
            return false;
        }
    }
}
