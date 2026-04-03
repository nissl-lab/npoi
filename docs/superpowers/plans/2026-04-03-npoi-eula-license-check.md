# NPOI OSMF EULA License Acceptance MSBuild Target Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** When any project referencing NPOI NuGet package is built without `<AcceptNPOIOSMFLicense>true</AcceptNPOIOSMFLicense>`, show a warning message in the IDE.

**Architecture:** Create an MSBuild targets file that gets included in the NPOI NuGet package. The target checks for the property during build and emits a warning if not set.

**Tech Stack:** MSBuild, NuGet packaging

---

### Task 1: Create MSBuild targets file for EULA license check

**Files:**
- Create: `build\NPOI.EULA.targets`

- [ ] **Step 1: Create the targets file with license check logic**

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="4.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <AcceptNPOIOSMFLicense Condition="'$(AcceptNPOIOSMFLicense)' == ''">false</AcceptNPOIOSMFLicense>
  </PropertyGroup>

  <Target Name="CheckNPOIOSMFLicense" BeforeTargets="BeforeBuild">
    <Warning Text="NPOI: You must accept the OSMF EULA license to use NPOI. Add &lt;AcceptNPOIOSMFLicense&gt;true&lt;/AcceptNPOIOSMFLicense&gt; to your project file." Condition="'$(AcceptNPOIOSMFLicense)' != 'true'" Importance="high" />
  </Target>
</Project>
```

- [ ] **Step 2: Commit**

```bash
git add build/NPOI.EULA.targets
git commit -m "feat: add MSBuild target to check OSMF EULA acceptance"
```

---

### Task 2: Include targets file in NuGet package

**Files:**
- Modify: `Directory.Build.props:113-119`

- [ ] **Step 1: Add the targets file to the NuGet package ItemGroup**

In `Directory.Build.props`, add the targets file to the ItemGroup:

```xml
<ItemGroup>
  <None Include="..\LICENSE" Pack="true" Visible="false" PackagePath="" />
  <None Include="..\OSMFEULA.txt" Pack="true" Visible="false" PackagePath="" />
  <None Include="..\build\README.md" Pack="true" Visible="false" PackagePath="" />
  <None Include="..\logo\*.png" Pack="true" Visible="false" PackagePath="logo\" />
  <None Include="..\logo\*.jpg" Pack="true" Visible="false" PackagePath="logo\" />
  <None Include="..\build\NPOI.EULA.targets" Pack="true" Visible="false" PackagePath="build\" />
</ItemGroup>
```

- [ ] **Step 2: Build the NuGet package to verify**

```bash
dotnet build solution/NPOI.Core.sln -c Release
dotnet pack main/NPOI.Core/NPOI.Core.csproj -c Release
```

- [ ] **Step 3: Commit**

```bash
git add Directory.Build.props
git commit -m "feat: include EULA license check targets in NuGet package"
```

---

### Task 3: Verify the implementation

**Files:**
- Test: Create a test project that references NPOI

- [ ] **Step 1: Create a test project to verify warning appears**

Create a minimal console app that references NPOI.Core without setting AcceptNPOIOSMFLicense. Build it and verify the warning appears in output.

- [ ] **Step 2: Add AcceptNPOIOSMFLicense to test project**

Add `<AcceptNPOIOSMFLicense>true</AcceptNPOIOSMFLicense>` to the test project and verify the warning disappears.

- [ ] **Step 3: Commit**

```bash
git add test-verification/
git commit -m "test: add verification for EULA license check target"
```
