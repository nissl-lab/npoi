using Nuke.Common;
using Nuke.Common.IO;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tools.DotNet;
using Nuke.Common.Utilities.Collections;
using static Nuke.Common.Tools.DotNet.DotNetTasks;

partial class Build : NukeBuild
{
    /// Support plugins are available for:
    ///   - JetBrains ReSharper        https://nuke.build/resharper
    ///   - JetBrains Rider            https://nuke.build/rider
    ///   - Microsoft VisualStudio     https://nuke.build/visualstudio
    ///   - Microsoft VSCode           https://nuke.build/vscode

    public static int Main () => Execute<Build>(x => x.Compile);

    [Parameter("Configuration to build - Default is 'Debug' (local) or 'Release' (server)")]
    readonly Configuration Configuration = IsLocalBuild ? Configuration.Debug : Configuration.Release;

    [Solution] Solution Solution;

    static AbsolutePath SourceDirectory => RootDirectory / "src";

    static AbsolutePath ArtifactsDirectory => RootDirectory / "artifacts";

    [Secret]
    [Parameter("GitHub API token")]
    readonly string GitHubToken;

    protected override void OnBuildInitialized()
    {
        Serilog.Log.Information("BUILD SETUP");
        Serilog.Log.Information("\tSolution: {Solution}", Solution);
        Serilog.Log.Information("\tConfiguration: {Configuration}", Configuration);

        Serilog.Log.Information("Build environment:");
        Serilog.Log.Information("\tHost: {Host}", Host.GetType());
    }

    Target Clean => _ => _
        .Before(Restore)
        .Executes(() =>
        {
            SourceDirectory.GlobDirectories("**/bin", "**/obj").ForEach(x => x.DeleteDirectory());
        });

    Target Restore => _ => _
        .Executes(() =>
        {
            DotNetRestore(_ => _
                .SetProjectFile(Solution)
            );
        });

    Target Compile => _ => _
        .DependsOn(Restore)
        .Executes(() =>
        {
            DotNetBuild(_ =>_
                .EnableNoRestore()
                .SetConfiguration(Configuration)
                .SetDeterministic(IsServerBuild)
                .SetContinuousIntegrationBuild(IsServerBuild)
                .SetVerbosity(DotNetVerbosity.Minimal)
                .AddNoWarns(1591) // missing XML documentation comment
                .SetProjectFile(Solution)
            );
        });

    Target Test => _ => _
        .DependsOn(Compile)
        .Executes(() =>
        {
            DotNetTest(_ =>_
                .EnableNoBuild()
                .EnableNoRestore()
                .SetConfiguration(Configuration)
                .SetProjectFile(Solution)
            );
        });

    Target Pack => _ => _
        .After(Test)
        .Produces(ArtifactsDirectory / "**")
        .Executes(() =>
        {
            DotNetPack(_ =>_
                .SetConfiguration(Configuration)
                .SetOutputDirectory(ArtifactsDirectory)
                .SetDeterministic(IsServerBuild)
                .SetContinuousIntegrationBuild(IsServerBuild)
                .SetProject(Solution.GetProject("NPOI.Core"))
            );
        });
}
