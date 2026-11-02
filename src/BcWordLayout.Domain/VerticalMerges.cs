using DocumentFormat.OpenXml.Wordprocessing;

namespace BcWordLayout.Domain;

/// <summary>
/// The vertically merged cells (<c>w:vMerge</c>) of one table (GitHub issue #9): which physical cells, row after
/// row, Word draws as one.
/// </summary>
/// <remarks>
/// <para>
/// THE SHAPE. A merge starts at a cell carrying <c>w:vMerge w:val="restart"</c> and takes in each cell straight
/// below it carrying <c>w:vMerge</c> with no value (or <c>continue</c>), each at the same grid column with the same
/// <c>w:gridSpan</c>. Word shows the first cell's content and keeps the others empty. Microsoft's own layouts use
/// it: the BC 29 header/footer designs merge their first column down two or three rows (<c>ExternalDefault.docx</c>
/// in the corpus, and seven siblings), <c>SpendRequestDocument.docx</c> merges a <c>gridSpan</c> 2 cell over two
/// rows of a <c>gridAfter</c> grid, and the NA demo invoices merge both side columns of their address block down
/// nine rows.
/// </para>
/// <para>
/// A <c>continue</c> cell with no merged cell straight above it (another grid column, another span, or nothing
/// merged there) is <see cref="Broken"/>: there is nothing for it to continue. It counts as a merge of its own, as
/// its first cell, so a tool treats it as an ordinary cell, and <see cref="TableGridConsistencyGuard"/> refuses an
/// edit that creates one.
/// </para>
/// </remarks>
internal sealed class VerticalMerges
{
    /// <summary>What a cell's <c>w:vMerge</c> says.</summary>
    internal enum State
    {
        /// <summary>No <c>w:vMerge</c>.</summary>
        None,

        /// <summary><c>w:vMerge w:val="restart"</c>: the first cell of a merge.</summary>
        Restart,

        /// <summary><c>w:vMerge</c> with no value or <c>continue</c>: merged into the cell above.</summary>
        Continue,
    }

    /// <summary>
    /// One cell of a merge: its row (<c>get_layout_info</c>'s numbering), its physical index in that row, and the
    /// grid columns it covers.
    /// </summary>
    internal sealed record Member(int Row, TableRow InnerRow, int Col, int GridColumn, int GridSpan, TableGridNavigator.CellSlot Cell);

    /// <summary>One merge: its cells top to bottom, the first one the cell whose content Word shows.</summary>
    internal sealed class Group
    {
        private readonly List<Member> members = [];

        /// <summary>The merged cells, top to bottom.</summary>
        public IReadOnlyList<Member> Members => members;

        /// <summary>The cell Word shows the merge's content from.</summary>
        public Member First => members[0];

        /// <summary>The last row the merge reaches.</summary>
        public int LastRow => members[^1].Row;

        /// <summary>True when the merge covers more than one row: one cell alone merges nothing.</summary>
        public bool SpansRows => members.Count > 1;

        internal void Add(Member member) => members.Add(member);

        /// <summary>"rows 0..2", for messages.</summary>
        public string DescribeRows() => $"rows {First.Row}..{LastRow}";
    }

    private readonly Dictionary<TableCell, Group> groupByCell;

    private VerticalMerges(IReadOnlyList<Group> groups, IReadOnlyList<Member> broken, Dictionary<TableCell, Group> groupByCell)
    {
        Groups = groups;
        Broken = broken;
        this.groupByCell = groupByCell;
    }

    /// <summary>Every merge, in the order their first cells appear.</summary>
    public IReadOnlyList<Group> Groups { get; }

    /// <summary>The <c>continue</c> cells with no merged cell straight above them (see the type remarks).</summary>
    public IReadOnlyList<Member> Broken { get; }

    /// <summary>True when the table merges no cell vertically, so a tool has nothing to keep whole.</summary>
    public bool IsEmpty => Groups.Count == 0;

    /// <summary>Reads <paramref name="table"/>'s merges, walking its rows as <see cref="TableGridNavigator.Rows"/> does.</summary>
    public static VerticalMerges Of(Table table)
    {
        ArgumentNullException.ThrowIfNull(table);

        var groups = new List<Group>();
        var broken = new List<Member>();
        var groupByCell = new Dictionary<TableCell, Group>(ReferenceEqualityComparer.Instance);

        // The merges the previous row's cells belong to, by the grid columns each cell covers: a continue cell
        // joins the merge whose cell above covers exactly its own columns.
        var above = new Dictionary<(int GridColumn, int GridSpan), Group>();
        var rows = TableGridNavigator.Rows(table);
        for (var r = 0; r < rows.Count; r++)
        {
            var current = new Dictionary<(int GridColumn, int GridSpan), Group>();
            if (rows[r].InnerRow is { } innerRow)
            {
                var cells = TableGridNavigator.Cells(innerRow);
                var grid = TableGridNavigator.GridEdges(innerRow).Before;
                for (var c = 0; c < cells.Count; c++)
                {
                    var member = new Member(r, innerRow, c, grid, cells[c].GridSpan, cells[c]);
                    grid += cells[c].GridSpan;

                    var state = StateOf(cells[c].InnerCell);
                    if (state == State.None)
                    {
                        continue;
                    }

                    var key = (member.GridColumn, member.GridSpan);
                    if (state != State.Continue || !above.TryGetValue(key, out var group))
                    {
                        if (state == State.Continue)
                        {
                            broken.Add(member);
                        }

                        group = new Group();
                        groups.Add(group);
                    }

                    group.Add(member);
                    current[key] = group;
                    groupByCell[cells[c].InnerCell!] = group;
                }
            }

            above = current;
        }

        return new VerticalMerges(groups, broken, groupByCell);
    }

    /// <summary>What <paramref name="cell"/>'s <c>w:vMerge</c> says; <see cref="State.None"/> for a null cell.</summary>
    public static State StateOf(TableCell? cell)
    {
        var merge = cell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>();
        if (merge is null)
        {
            return State.None;
        }

        return merge.Val?.Value == MergedCellValues.Restart ? State.Restart : State.Continue;
    }

    /// <summary>The merge <paramref name="cell"/> belongs to, or null when it carries no <c>w:vMerge</c>.</summary>
    public Group? GroupOf(TableCell? cell) =>
        cell is not null && groupByCell.TryGetValue(cell, out var group) ? group : null;

    /// <summary>
    /// The merge <paramref name="cell"/> continues when it is not that merge's first cell, so Word shows another
    /// row's content in its place; null otherwise.
    /// </summary>
    public Group? ContinuedBy(TableCell? cell) =>
        GroupOf(cell) is { } group && !ReferenceEquals(group.First.Cell.InnerCell, cell) ? group : null;

    /// <summary>
    /// The cell whose content Word shows where <paramref name="slot"/> is: the merge's first cell for a cell that
    /// continues one, else the slot itself.
    /// </summary>
    public TableGridNavigator.CellSlot Shown(TableGridNavigator.CellSlot slot) =>
        ContinuedBy(slot.InnerCell) is { } group ? group.First.Cell : slot;

    /// <summary>Gives <paramref name="cell"/> the <c>w:vMerge</c> of <paramref name="state"/>, in its schema place.</summary>
    internal static void SetState(TableCell cell, State state)
    {
        ArgumentNullException.ThrowIfNull(cell);

        var tcPr = cell.GetFirstChild<TableCellProperties>();
        tcPr?.RemoveAllChildren<VerticalMerge>();
        if (state == State.None)
        {
            return;
        }

        if (tcPr is null)
        {
            tcPr = new TableCellProperties();
            cell.PrependChild(tcPr);
        }

        // Word writes a continue cell as a bare <w:vMerge/>, as every corpus merge does.
        var merge = state == State.Restart ? new VerticalMerge { Val = MergedCellValues.Restart } : new VerticalMerge();
        tcPr.AddChild(merge, throwOnError: false);
    }
}
