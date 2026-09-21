# Releasing ConfigMgr Health Dashboard

## 1. Test

This repository has no Pester suite. Parse the entry script and import the module under Windows PowerShell 5.1. The release stops on a parse error or an import error.

```powershell
powershell -NoProfile -Command "$e = $null; [void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path .\start-mecmhealthdashboard.ps1), [ref]$null, [ref]$e); 'parse errors {0}' -f @($e).Count; Import-Module .\Module\MECMHealthDashCommon.psd1 -Force -ErrorAction Stop; 'import ok'"
```

## 2. Version

The version is `YYYY.MM.DD.BBBB`: the release date, then a four-digit, zero-padded build number. The build number increases by 1 for each release and never resets. Build 0008 is the first release with this scheme; the 7 releases before it used `1.x.y` numbers. Keep the zero-padded text everywhere; `[version]` drops the leading zeros.

Set the same version in three places:

- `Module/MECMHealthDashCommon.psd1`, `ModuleVersion`
- `start-mecmhealthdashboard.ps1`, header line `Version    : <ver>`
- `CHANGELOG.md`, the top heading `## [<ver>] - <date>`

Add the new `CHANGELOG.md` entry at the top. Do not edit the entries of earlier releases.

## 3. Shared module

Check the vendored SuiteCommon copy for drift. Sync it if the check reports drift.

```powershell
C:\projects\app-packager-suite\sync-suitecommon.ps1 -Consumer C:\projects\mecm-health-dashboard -Check
```

## 4. Commit, tag, and package

Commit to `master`. Every commit to `master` is part of a release: the suite installer build refuses a component whose `master` is ahead of its latest tag. Tag the commit `v<ver>` with an annotated tag. Build the zip from the tag. The zip excludes the tests and this file.

```bash
git tag -a v<ver> -m v<ver>
git archive --format=zip -o ConfigMgrHealthDashboard-<ver>.zip v<ver> -- . ':(exclude)Tests' ':(exclude)*.Tests.ps1' ':(exclude)RELEASING.md'
sha256sum ConfigMgrHealthDashboard-<ver>.zip | sed 's/ \*/  /' > checksums.txt
```

Extract the zip to a temporary folder. Import `Module/MECMHealthDashCommon.psd1` under Windows PowerShell 5.1. The import must succeed.

## 5. Publish

Push `master` and the tag. Create the GitHub release with the title `v<ver>` and two assets: the zip and `checksums.txt`. The release must not be a draft. The release notes have a `##` headline with one concrete outcome metric, the `###` sections of the changelog entry, and the footer `Full changelog: CHANGELOG.md`.
