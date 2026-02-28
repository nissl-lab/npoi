using Nuke.Common.CI.GitHubActions;
using Nuke.Common.CI.GitHubActions.Configuration;
using Nuke.Common.Execution;
using Nuke.Common.Tooling;
using Nuke.Common.Utilities;
using Nuke.Common.Utilities.Collections;
using System.Collections.Generic;

[CustomGitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    OnPushBranches = ["main", "master", "release*", "poi/*"],
    InvokedTargets = [nameof(Clean), nameof(Test), nameof(Pack)],
    TimeoutMinutes = 20,
    CacheKeyFiles = [],
    PublishCondition = "runner.os == 'Linux'"
)]
[CustomGitHubActions("PR",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    On = [GitHubActionsTrigger.PullRequest],
    InvokedTargets = [nameof(Clean), nameof(Test), nameof(Pack)],
    TimeoutMinutes = 20,
    CacheKeyFiles = [],
    ConcurrencyCancelInProgress = true,
    PublishCondition = "runner.os == 'Linux'"
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

        newSteps.Insert(0, new GitHubActionsSetupDotNetStep(["10.0"]));

        job.Steps = newSteps.ToArray();

        return new GitHubActionsJobContinueOnError(job)
        {
            ContinueOnError = image == GitHubActionsImage.WindowsLatest
        };
    }
}

class GitHubActionsJobContinueOnError : GitHubActionsJob
{
    public GitHubActionsJobContinueOnError(GitHubActionsJob job)
    {
        Name = job.Name;
        Image = job.Image;
        TimeoutMinutes = job.TimeoutMinutes;
        ConcurrencyGroup = job.ConcurrencyGroup;
        ConcurrencyCancelInProgress = job.ConcurrencyCancelInProgress;
        Steps = job.Steps;
    }

    public bool ContinueOnError { get; set; }

    public override void Write(CustomFileWriter writer)
    {
        writer.WriteLine($"{Name}:");

        using (writer.Indent())
        {
            writer.WriteLine($"name: {Name}");
            writer.WriteLine($"runs-on: {Image.GetValue()}");

            if (ContinueOnError)
                writer.WriteLine("continue-on-error: true");

            if (TimeoutMinutes > 0)
                writer.WriteLine($"timeout-minutes: {TimeoutMinutes}");

            if (!ConcurrencyGroup.IsNullOrWhiteSpace() || ConcurrencyCancelInProgress)
            {
                writer.WriteLine("concurrency:");
                using (writer.Indent())
                {
                    var group = ConcurrencyGroup;
                    if (group.IsNullOrWhiteSpace())
                        group = "${{ github.workflow }} @ ${{ github.event.pull_request.head.label || github.head_ref || github.run_id }}";

                    writer.WriteLine($"group: {group}");
                    if (ConcurrencyCancelInProgress)
                        writer.WriteLine("cancel-in-progress: true");
                }
            }

            writer.WriteLine("steps:");
            using (writer.Indent())
            {
                Steps.ForEach(x => x.Write(writer));
            }
        }
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
        writer.WriteLine("- uses: actions/setup-dotnet@v5");

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