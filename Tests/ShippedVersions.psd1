# Reference values for every version list that the shipped templates carry.
# TestsShippedLists.Tests.ps1 compares Templates*.json with this file.
# Refresh steps: docs\VERSION-REFRESH.md. Each section names its source page.
# Every value below was read from the source page on VerifiedOn.
@{
    VerifiedOn = '2026-11-14'

    # Newest entry that each list must carry. A test fails when the template list has no entry for it.
    Newest = @{
        ConfigMgrClient   = '2609'
        M365AppsBuild     = '2610'
        Windows11Build    = 'Windows 11 v26H2'
        Windows10Build    = 'Windows 10 v22H2'
        ServerBuild       = 'Windows 2025'
    }

    # ConfigMgr current branch version -> client and site build (5.00.<build>.<revision>).
    # The template matches the build with ClientVersion like '5.00.<build>.10%'.
    ConfigMgrClient = @{
        Source   = 'https://learn.microsoft.com/intune/configmgr/core/servers/manage/updates'
        Versions = @{
            '1511' = '8325'; '1602' = '8355'; '1606' = '8412'; '1610' = '8458'; '1702' = '8498'
            '1706' = '8540'; '1710' = '8577'; '1802' = '8634'; '1806' = '8692'; '1810' = '8740'
            '1902' = '8790'; '1906' = '8853'; '1910' = '8913'; '2002' = '8968'; '2006' = '9012'
            '2010' = '9040'; '2103' = '9049'; '2107' = '9058'; '2111' = '9068'; '2203' = '9078'
            '2207' = '9088'; '2211' = '9096'; '2303' = '9106'; '2309' = '9122'; '2403' = '9128'
            '2409' = '9132'; '2503' = '9135'; '2509' = '9141'; '2603' = '9146'; '2609' = '9152'
        }
    }

    # Microsoft 365 Apps version -> build (16.0.<build>.<revision>).
    M365AppsBuild = @{
        Source   = 'https://learn.microsoft.com/officeupdates/update-history-microsoft365-apps-by-date'
        Versions = @{
            '2001' = '12430'; '2002' = '12527'; '2003' = '12624'; '2004' = '12730'; '2005' = '12827'; '2006' = '13001'
            '2007' = '13029'; '2008' = '13127'; '2009' = '13231'; '2010' = '13328'; '2011' = '13426'; '2012' = '13530'
            '2101' = '13628'; '2102' = '13801'; '2103' = '13901'; '2104' = '13929'; '2105' = '14026'; '2106' = '14131'
            '2107' = '14228'; '2108' = '14326'; '2109' = '14430'; '2110' = '14527'; '2111' = '14701'; '2112' = '14729'
            '2201' = '14827'; '2202' = '14931'; '2203' = '15028'; '2204' = '15128'; '2205' = '15225'; '2206' = '15330'
            '2207' = '15427'; '2208' = '15601'; '2209' = '15629'; '2210' = '15726'; '2211' = '15831'; '2212' = '15928'
            '2301' = '16026'; '2302' = '16130'; '2303' = '16227'; '2304' = '16327'; '2305' = '16501'; '2306' = '16529'
            '2307' = '16626'; '2308' = '16731'; '2309' = '16827'; '2310' = '16924'; '2311' = '17029'; '2312' = '17126'
            '2401' = '17231'; '2402' = '17328'; '2403' = '17425'; '2404' = '17531'; '2405' = '17628'; '2406' = '17726'
            '2407' = '17830'; '2408' = '17928'; '2409' = '18025'; '2410' = '18129'; '2411' = '18227'; '2412' = '18324'
            '2501' = '18429'; '2502' = '18526'; '2503' = '18623'; '2504' = '18730'; '2505' = '18827'; '2506' = '18925'
            '2507' = '19029'; '2508' = '19127'; '2509' = '19231'; '2510' = '19328'; '2511' = '19426'; '2512' = '19530'
            '2601' = '19628'; '2602' = '19725'; '2603' = '19822'; '2604' = '19929'; '2605' = '20026'; '2606' = '20131'
            '2607' = '20228'; '2608' = '20326'; '2609' = '20430'; '2610' = '20527'
        }
        # Versions 1705 to 1912 are old Office 365 ProPlus builds that the current page no longer lists.
        # They are kept as shipped. Status: Unverified.
        Unverified = @{
            '1705' = '8201'; '1708' = '8431'; '1802' = '9029'; '1803' = '9126'; '1808' = '10730'
            '1902' = '11328'; '1908' = '11929'; '1912' = '12325'
        }
    }

    # Microsoft 365 Apps update channel -> CDN GUID (SMS_G_System_OFFICE365PROPLUSCONFIGURATIONS.cfgUpdateChannel).
    M365AppsChannel = @{
        Source   = 'https://learn.microsoft.com/intune/device-configuration/settings-catalog/update-office'
        Channels = @{
            'Current Channel'                              = '492350f6-3a01-4f97-b9c0-c7c6ddf67d60'
            'Current Channel (Preview)'                    = '64256afe-f5d9-4f86-8936-8840a6a4f5be'
            'Monthly Enterprise Channel'                   = '55336b82-a18d-4dd6-b5f6-9e5095c314a6'
            'Semi-Annual Enterprise Channel'               = '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114'
            'Semi-Annual Enterprise Channel (Preview)'     = 'b8f9b850-328d-4355-9145-c59439a0c4cf'
            'Beta Channel'                                 = '5440fd1f-7ecb-4221-8110-145efaa6372f'
        }
    }

    # Windows client name -> OS build (SMS_R_System.Build = '10.0.<build>').
    WindowsClientBuild = @{
        Source = 'https://learn.microsoft.com/windows/release-health/release-information and https://learn.microsoft.com/windows/release-health/windows11-release-information'
        Builds = @{
            'Windows 10 v1507' = '10240'; 'Windows 10 v1511' = '10586'; 'Windows 10 v1607' = '14393'
            'Windows 10 v1703' = '15063'; 'Windows 10 v1709' = '16299'; 'Windows 10 v1803' = '17134'
            'Windows 10 v1809' = '17763'; 'Windows 10 v1903' = '18362'; 'Windows 10 v1909' = '18363'
            'Windows 10 v2004' = '19041'; 'Windows 10 v20H2' = '19042'; 'Windows 10 v21H1' = '19043'
            'Windows 10 v21H2' = '19044'; 'Windows 10 v22H2' = '19045'
            'Windows 11 v21H2' = '22000'; 'Windows 11 v22H2' = '22621'; 'Windows 11 v23H2' = '22631'
            'Windows 11 v24H2' = '26100'; 'Windows 11 v25H2' = '26200'; 'Windows 11 v26H1' = '28000'
            'Windows 11 v26H2' = '26300'
        }
    }

    # Windows Server name -> OS build (SMS_G_System_OPERATING_SYSTEM.BuildNumber).
    WindowsServerBuild = @{
        Source = 'https://learn.microsoft.com/windows/release-health/windows-server-release-info'
        Builds = @{
            'Windows 2016' = '14393'; 'Windows 2019' = '17763'; 'Windows 2022' = '20348'; 'Windows 2025' = '26100'
        }
    }
}
