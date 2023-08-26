using System;
using System.Linq;
using System.Runtime.InteropServices;
using Nuke.Common;
using Nuke.Common.CI.GitHubActions;
using Nuke.Common.Git;
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
    [GitRepository] readonly GitRepository GitRepository;

    static AbsolutePath SourceDirectory => RootDirectory / "src";

    static AbsolutePath ArtifactsDirectory => RootDirectory / "publish";

    string TagVersion => GitRepository.Tags.SingleOrDefault(x => x.StartsWith("v"))?[1..];

    string BranchVersion => GitRepository.Branch?.StartsWith("release") == true
        ? GitRepository.Branch[7..]
        : null;

    // either from tag or branch
    string PublishVersion => TagVersion ?? BranchVersion;

    bool IsPublishBuild => !string.IsNullOrWhiteSpace(PublishVersion);

    string VersionSuffix;

    static bool IsRunningOnWindows => RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

    [Secret]
    [Parameter("GitHub API token")]
    readonly string GitHubToken;

    protected override void OnBuildInitialized()
    {
        VersionSuffix = !IsPublishBuild
            ? $"preview-{DateTime.UtcNow:yyyyMMdd-HHmm}"
            : "";

        if (IsLocalBuild)
        {
            VersionSuffix = $"dev-{DateTime.UtcNow:yyyyMMdd-HHmm}";
        }

        Serilog.Log.Information("BUILD SETUP");
        Serilog.Log.Information("\tSolution: {Solution}", Solution);
        Serilog.Log.Information("\tConfiguration: {Configuration}", Configuration);
        Serilog.Log.Information("\tVersion suffix: {VersionSuffix}", VersionSuffix);
        Serilog.Log.Information("\tPublish version: {PublishVersion}", PublishVersion);

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
            static void StartSudoProcess(string arguments) => ProcessTasks.StartProcess("sudo", arguments).WaitForExit();

            // replace broken font - the one coming from APT doesn't contain all expected tables
            StartSudoProcess("rm /usr/share/fonts/truetype/noto/NotoColorEmoji.ttf");
            StartSudoProcess("curl -sS -L -o /usr/share/fonts/truetype/noto/NotoColorEmoji-Regular.ttf https://fonts.gstatic.com/s/notocoloremoji/v25/Yq6P-KqIXTD0t4D9z1ESnKM3-HpFab5s79iz64w.ttf");
        });

    Target Pack => _ => _
        .After(Test)
        .Produces(ArtifactsDirectory / "**")
        .Executes(() =>
        {
            // make sure we make fresh build
            DeleteCompilationArtifacts();

            var packTarget = Solution.GetProject("NPOI.Pack");

            DotNetPack(_ =>
            {
                var packSettings = _
                    .SetProject(packTarget)
                    .SetConfiguration(Configuration)
                    .SetOutputDirectory(ArtifactsDirectory)
                    .SetDeterministic(IsServerBuild)
                    .SetContinuousIntegrationBuild(IsServerBuild)
                    // obsolete missing XML documentation comment, XML comment on not valid language element, XML comment has badly formed XML, no matching tag in XML comment
                    // need to use escaped separator in order for this to work
                    .AddProperty("NoWarn", string.Join("%3B", new[] { 169, 612, 618, 1591, 1587, 1570, 1572, 1573, 1574 }))
                    .SetProperty("EnablePackageValidation", "false");

                if (IsPublishBuild)
                {
                    // force version from tag/branch
                    packSettings = packSettings
                        .SetAssemblyVersion(PublishVersion)
                        .SetFileVersion(PublishVersion)
                        .SetInformationalVersion(PublishVersion)
                        .SetVersionSuffix(VersionSuffix)
                        .SetVersionPrefix(PublishVersion);
                }

                return packSettings;
            });
        });
}
