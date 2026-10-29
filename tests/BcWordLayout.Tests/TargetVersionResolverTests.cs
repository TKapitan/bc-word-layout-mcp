using BcWordLayout.Domain;

namespace BcWordLayout.Tests;

/// <summary>
/// <see cref="TargetVersionResolver"/> over real folder trees (GitHub issue #152, ADR-0007): the nearest
/// <c>app.json</c> above the layout, then <c>BCWL_BC_VERSION</c>, else unknown with the reason.
/// </summary>
public class TargetVersionResolverTests
{
    [Fact]
    public void The_nearest_app_json_above_the_layout_wins()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "28.0.0.0");
        var nearest = repo.AppJson("app", "29.0.0.0");

        var target = TargetVersionResolver.Resolve(repo.Path("app/src/layouts/Invoice.docx"), environmentValue: null);

        Assert.Equal(new BcVersion(29, 0), target.Version);
        Assert.Equal(TargetVersionSource.AppJson, target.Source);
        Assert.Equal("app.json", target.SourceName);
        Assert.Equal($"{nearest}: application 29.0.0.0", target.Detail);
    }

    [Fact]
    public void The_walk_stops_at_the_repository_root()
    {
        // An app.json outside the repository belongs to some other project, not to this layout.
        using var outer = new TargetFolder();
        Directory.Delete(Path.Combine(outer.Root, ".git"));
        outer.AppJson(".", "29.0.0.0");
        Directory.CreateDirectory(outer.Path("repo/.git/HEAD"));

        var target = TargetVersionResolver.Resolve(outer.Path("repo/layouts/Invoice.docx"), environmentValue: null);

        Assert.Null(target.Version);
        Assert.Equal(TargetVersionSource.Unknown, target.Source);
        Assert.Contains("(the repository root)", target.Detail, StringComparison.Ordinal);
        Assert.Contains("BCWL_BC_VERSION is not set", target.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void A_git_file_marks_a_worktree_root_too()
    {
        using var outer = new TargetFolder();
        Directory.Delete(Path.Combine(outer.Root, ".git"));
        outer.AppJson(".", "29.0.0.0");
        outer.Write("worktree/.git", "gitdir: C:/elsewhere/.git/worktrees/w");

        var target = TargetVersionResolver.Resolve(outer.Path("worktree/Invoice.docx"), environmentValue: null);

        Assert.Null(target.Version);
    }

    [Fact]
    public void App_json_beats_the_environment_variable()
    {
        using var repo = new TargetFolder();
        repo.AppJson(".", "28.0.0.0");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: "29");

        Assert.Equal(new BcVersion(28, 0), target.Version);
        Assert.Equal(TargetVersionSource.AppJson, target.Source);
    }

    [Fact]
    public void An_app_json_without_application_falls_through_to_the_environment_variable()
    {
        using var repo = new TargetFolder();
        var appJson = repo.Write("app.json", """{ "name": "Layouts", "platform": "29.0.0.0" }""");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: "29");

        Assert.Equal(new BcVersion(29, 0), target.Version);
        Assert.Equal(TargetVersionSource.Environment, target.Source);
        Assert.Equal($"BCWL_BC_VERSION=29 ({appJson} has no application property)", target.Detail);
    }

    [Fact]
    public void An_app_json_without_application_and_no_variable_is_unknown()
    {
        using var repo = new TargetFolder();
        var appJson = repo.Write("app.json", """{ "name": "Layouts" }""");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: null);

        Assert.Null(target.Version);
        Assert.Equal($"{appJson} has no application property, and BCWL_BC_VERSION is not set", target.Detail);
    }

    [Fact]
    public void App_json_may_carry_comments_and_trailing_commas()
    {
        using var repo = new TargetFolder();
        repo.Write("app.json", """
            {
              // The AL tooling accepts both.
              "name": "Layouts",
              "application": "29.1.0.0", /* minimum */
            }
            """);

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: null);

        Assert.Equal(new BcVersion(29, 1), target.Version);
    }

    [Theory]
    [InlineData("{ \"application\": ")] // malformed JSON
    [InlineData("[ \"29.0.0.0\" ]")] // not an object
    public void An_unreadable_app_json_makes_the_target_unknown_instead_of_falling_through(string content)
    {
        using var repo = new TargetFolder();
        var appJson = repo.Write("app.json", content);

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: "29");

        Assert.Null(target.Version);
        Assert.StartsWith($"{appJson} could not be read:", target.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void An_oversized_app_json_is_not_parsed()
    {
        using var repo = new TargetFolder();
        repo.Write("app.json", "{ \"application\": \"29.0.0.0\", \"x\": \"" + new string('x', 1024 * 1024) + "\" }");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: null);

        Assert.Null(target.Version);
        Assert.Contains("more than the 1048576 an app.json may have", target.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void An_application_that_is_not_a_version_makes_the_target_unknown()
    {
        using var repo = new TargetFolder();
        var appJson = repo.AppJson(".", "latest");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: "29");

        Assert.Null(target.Version);
        Assert.Equal($"{appJson}: application 'latest' is not a Business Central version", target.Detail);
    }

    [Fact]
    public void The_environment_variable_names_the_target_outside_an_AL_project()
    {
        using var repo = new TargetFolder();

        var target = TargetVersionResolver.Resolve(repo.Path("layouts/Invoice.docx"), environmentValue: " 29.1 ");

        Assert.Equal(new BcVersion(29, 1), target.Version);
        Assert.Equal(TargetVersionSource.Environment, target.Source);
        Assert.Equal("environment", target.SourceName);
        Assert.Equal("BCWL_BC_VERSION=29.1", target.Detail);
    }

    [Fact]
    public void An_invalid_environment_value_makes_the_target_unknown_with_the_reason()
    {
        using var repo = new TargetFolder();

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: "next");

        Assert.Null(target.Version);
        Assert.StartsWith("BCWL_BC_VERSION='next' is not a Business Central version", target.Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void With_neither_the_target_is_unknown_and_says_where_it_looked()
    {
        using var repo = new TargetFolder();

        var target = TargetVersionResolver.Resolve(repo.Path("layouts/Invoice.docx"), environmentValue: null);

        Assert.Null(target.Version);
        Assert.Equal("unknown", target.SourceName);
        Assert.Equal(
            $"no app.json above the layout, searched up to {repo.Root} (the repository root), and BCWL_BC_VERSION is not set",
            target.Detail);
        Assert.Equal($"unknown ({target.Detail})", target.Describe());
    }

    [Fact]
    public void A_layout_that_does_not_exist_yet_resolves_from_the_folders_above_it()
    {
        // create_layout resolves the target for its output path, whose folder it may still have to create.
        using var repo = new TargetFolder();
        repo.AppJson(".", "29.0.0.0");

        var target = TargetVersionResolver.Resolve(Path.Combine(repo.Root, "new", "folder", "Invoice.docx"), environmentValue: null);

        Assert.Equal(new BcVersion(29, 0), target.Version);
    }

    [Fact]
    public void A_known_target_describes_its_version_and_source()
    {
        using var repo = new TargetFolder();
        var appJson = repo.AppJson(".", "28.0.0.0");

        var target = TargetVersionResolver.Resolve(repo.Path("Invoice.docx"), environmentValue: null);

        Assert.Equal(
            $"BC 28.0 (2026 release wave 1), from app.json ({appJson}: application 28.0.0.0)",
            target.Describe());
    }
}
