using Nuke.Common.CI.GitHubActions;
using Nuke.Common.CI.GitHubActions.Configuration;
using Nuke.Common.Execution;
using Nuke.Common.Utilities;
using System.Collections.Generic;

[CustomGitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    OnPushBranches = ["main", "master", "release*", "poi/*"],
    InvokedTargets = [nameof(Clean), nameof(Test), nameof(Pack)],
    TimeoutMinutes = 20,
    CacheKeyFiles = [],
    PublishCondition = "runner.os == 'Windows'"
)]
[CustomGitHubActions("PR",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    On = [GitHubActionsTrigger.PullRequest],
    InvokedTargets = [nameof(Clean), nameof(Test), nameof(Pack)],
    TimeoutMinutes = 20,
    CacheKeyFiles = [],
    ConcurrencyCancelInProgress = true,
    PublishCondition = "runner.os == 'Windows'"
)]
partial class Build;

class CustomGitHubActionsAttribute : GitHubActionsAttribute
{
    public CustomGitHubActionsAttribute(string name, GitHubActionsImage image, params GitHubActionsImage[] images) : base(name, image, images)
    {
    }

    protected override GitHubActionsJob GetJobs(GitHubActionsImage image, IReadOnlyCollection<ExecutableTarget> relevantTargets)
    {
        var job = base.GetJobs(image, relevantTargets);

        var newSteps = new List<GitHubActionsStep>(job.Steps);

        newSteps.Insert(0, new GitHubActionsSetupDotNetStep(["8.0", "9.0"]));

        job.Steps = newSteps.ToArray();
        return job;
    }
}

class GitHubActionsSetupDotNetStep : GitHubActionsStep
{
    public GitHubActionsSetupDotNetStep(string[] versions)
    {
        Versions = versions;
    }

    string[] Versions { get; }

    public override void Write(CustomFileWriter writer)
    {
        writer.WriteLine("- uses: actions/setup-dotnet@v4");

        using (writer.Indent())
        {
            writer.WriteLine("with:");
            using (writer.Indent())
            {
                writer.WriteLine("dotnet-version: |");
                using (writer.Indent())
                {
                    foreach (var version in Versions)
                    {
                        writer.WriteLine(version);
                    }
                }
            }
        }
    }
}