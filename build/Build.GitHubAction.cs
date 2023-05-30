using Nuke.Common.CI.GitHubActions;

[GitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    OnPushBranches = new[] { "main", "master" },
    InvokedTargets = new[] { nameof(Clean), nameof(Compile), nameof(Pack) }
)]
[GitHubActions("PR",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    On = new [] { GitHubActionsTrigger.PullRequest },
    InvokedTargets = new[] { nameof(Clean), nameof(Compile), nameof(Pack) }
)]
partial class Build
{
}

