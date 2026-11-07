# Releasing Collection Manager

## 1. Test

Run the checks under Windows PowerShell 5.1 and under PowerShell 7. The release stops if the lint, a test or the coverage floor fails. Window tests open windows off screen and need a desktop session.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Invoke-Checks.ps1
pwsh -NoProfile -File .\Invoke-Checks.ps1
```

When the release changes a write action, also run the lab test on CM01. It creates ZZTEST- collections and one ZZTEST- saved query, runs the actions, removes everything, and checks that no other collection changed. See `Tests/TESTPLAN.md`.

```powershell
pwsh -NoProfile -File .\Tests\Lab\Invoke-CollectionLab.ps1
```

## 2. Version

The version is `YYYY.MM.DD.BBBB`: the release date, then a four-digit, zero-padded build number. The build number increases by 1 for each release and never resets. Build 0008 is the first release with this scheme; the 7 releases before it used `1.x.y` numbers. Keep the zero-padded text everywhere; `[version]` drops the leading zeros.

Set the same version in three places:

- `Module/CollectionManagerCommon.psd1`, `ModuleVersion`
- `start-collectionmanager.ps1`, header line `Version    : <ver>`
- `CHANGELOG.md`, the top heading `## [<ver>] - <date>`

Add the new `CHANGELOG.md` entry at the top. Do not edit the entries of earlier releases.

## 3. Shared module

Check the vendored SuiteCommon copy for drift. Sync it if the check reports drift.

```powershell
C:\projects\app-packager-suite\sync-suitecommon.ps1 -Consumer C:\projects\collection-manager -Check
```

## 4. Commit, tag, and package

Commit to `main`. Every commit to `main` is part of a release: the suite installer build refuses a component whose `main` is ahead of its latest tag. Tag the commit `v<ver>` with an annotated tag. Build the zip from the tag. The zip excludes the tests, the check script, the workflows and this file.

```bash
git tag -a v<ver> -m v<ver>
git archive --format=zip -o CollectionManager-<ver>.zip v<ver> -- . ':(exclude)Tests' ':(exclude)*.Tests.ps1' ':(exclude)RELEASING.md' ':(exclude)Invoke-Checks.ps1' ':(exclude).github'
sha256sum CollectionManager-<ver>.zip | sed 's/ \*/  /' > checksums.txt
```

Extract the zip to a temporary folder. Import `Module/CollectionManagerCommon.psd1` under Windows PowerShell 5.1. The import must succeed.

## 5. Publish

Push `main` and the tag. Create the GitHub release with the title `v<ver>` and two assets: the zip and `checksums.txt`. The release must not be a draft. The release notes have a `##` headline with one concrete outcome metric, the `###` sections of the changelog entry, and the footer `Full changelog: CHANGELOG.md`.
