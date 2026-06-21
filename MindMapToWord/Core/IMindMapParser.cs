using System.IO;

namespace MindMapToWord.Core
{
    public interface IMindMapParser
    {
        string FormatName { get; }
        string[] SupportedExtensions { get; }
        MindMapDocument Parse(string filePath);
        bool CanParse(string filePath);
    }
}
