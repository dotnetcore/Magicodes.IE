
using System.Collections.Generic;

namespace Magicodes.IE.IO
{

    internal sealed class SheetState
    {
        public List<ImageAnchor> Images { get; } = new();
        public List<string> MergeCells { get; } = new();
        public string? AutoFilter { get; set; }
        public List<(string Ref, string Uri)> Hyperlinks { get; } = new();
        public List<DataValidation> DataValidations { get; } = new();
        public List<TableDefinition> Tables { get; } = new();
        public SheetProtection? Protection { get; set; }
        public PageSetup? PageSetup { get; set; }
        public OutlineSettings? Outline { get; set; }
        public List<Comment> Comments { get; } = new();
        public List<ConditionalFormatting> ConditionalFormattings { get; } = new();
        public int[]? StyleIds { get; set; }
        public ColumnMeta[]? Columns { get; set; }
    }
}
