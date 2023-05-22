using Nuke.Common.CI.GitHubActions;

[GitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    OnPushBranches = new [] { "master", "main" },
    OnPullRequestBranches = new [] { "master", "main" },
    InvokedTargets = new[] { nameof(Clean), nameof(Compile), nameof(Test), nameof(Pack) }
)]
partial class Build
{
}

