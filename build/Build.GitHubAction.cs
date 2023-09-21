using Nuke.Common.CI.GitHubActions;

[GitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    OnPushBranches = new[] { "main", "master", "release*" },
    InvokedTargets = new[] { nameof(Clean), nameof(Test), nameof(Pack) },
    TimeoutMinutes = 20,
    CacheKeyFiles = new string[0]
)]
[GitHubActions("PR",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    On = new [] { GitHubActionsTrigger.PullRequest },
    InvokedTargets = new[] { nameof(Clean), nameof(Test), nameof(Pack) },
    TimeoutMinutes = 20,
    CacheKeyFiles = new string[0]
)]
partial class Build
{
}

