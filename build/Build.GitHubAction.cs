using Nuke.Common.CI.GitHubActions;

[GitHubActions("CI",
    GitHubActionsImage.WindowsLatest,
    GitHubActionsImage.UbuntuLatest,
    OnPushBranches = ["main", "master", "release*", "poi/*"],
    InvokedTargets = [nameof(Clean), nameof(Test), nameof(Pack)],
    TimeoutMinutes = 20,
    CacheKeyFiles = [],
    PublishCondition = "runner.os == 'Windows'"
)]
[GitHubActions("PR",
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

