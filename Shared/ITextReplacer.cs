using EnvDTE;

namespace Shared
{
    public interface ITextReplacer
    {
        void Replace(string whatRegex, string with, Document document);
    }
}
