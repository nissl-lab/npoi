using System.Runtime.InteropServices;
using Nuke.Common;
using Nuke.Common.CI.GitHubActions;
using Nuke.Common.IO;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tooling;
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
            DeleteCompilationArtifacts();
            ArtifactsDirectory.CreateOrCleanDirectory();
        });

    static void DeleteCompilationArtifacts()
    {
        var solutionDirectory = RootDirectory / "solution";
        solutionDirectory.GlobDirectories("**/bin", "**/obj").ForEach(x => x.DeleteDirectory());
        (solutionDirectory / "Debug").DeleteDirectory();
        (solutionDirectory / "Release").DeleteDirectory();
    }

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
                // obsolete missing XML documentation comment, XML comment on not valid language element, XML comment has badly formed XML, no matching tag in XML comment
                // need to use escaped separator in order for this to work
                .AddProperty("NoWarn", string.Join("%3B", new [] { 169, 612, 618, 1591, 1587, 1570, 1572, 1573, 1574 }))
                .SetProjectFile(Solution)
            );

            // copy files from projects in order to get them to be part of pack

        });

    Target Test => _ => _
        .DependsOn(Compile, InstallFonts)
        .Executes(() =>
        {
            DotNetTest(_ =>_
                .EnableNoBuild()
                .EnableNoRestore()
                .SetConfiguration(Configuration)
                .SetProjectFile(Solution)
                .When(Host is GitHubActions, settings => settings.SetLoggers("GitHubActions"))
                .When(!RuntimeInformation.IsOSPlatform(OSPlatform.Windows), settings => settings.SetFramework("net6.0"))
            );
        });

    Target InstallFonts => _ => _
        .OnlyWhenDynamic(() => RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && Host is GitHubActions)
        .Executes(() =>
        {
            ProcessTasks.StartProcess("sudo", "apt install -y fonts-noto-color-emoji");
            ProcessTasks.StartProcess("mkdir", "-p /usr/local/share/fonts");
            ProcessTasks.StartProcess("cp", "/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf /usr/local/share/fonts/");
            ProcessTasks.StartProcess("chmod", "644 /usr/local/share/fonts/NotoColorEmoji.ttf");
            ProcessTasks.StartProcess("fc-cache", "-fv");
        });

    Target Pack => _ => _
        .After(Test)
        .Produces(ArtifactsDirectory / "**")
        .Executes(() =>
        {
            // make sure we make fresh build
            DeleteCompilationArtifacts();

            var packTarget = Solution.GetProject("NPOI.Pack");

            DotNetPack(_ =>_
                .SetConfiguration(Configuration)
                .SetOutputDirectory(ArtifactsDirectory)
                .SetDeterministic(IsServerBuild)
                .SetContinuousIntegrationBuild(IsServerBuild)
                // obsolete missing XML documentation comment, XML comment on not valid language element, XML comment has badly formed XML, no matching tag in XML comment
                // need to use escaped separator in order for this to work
                .AddProperty("NoWarn", string.Join("%3B", new [] { 169, 612, 618, 1591, 1587, 1570, 1572, 1573, 1574 }))
                .SetProperty("EnablePackageValidation", "false")
                .SetProject(packTarget)
            );
        });
}
