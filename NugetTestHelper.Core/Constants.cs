using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuGetTestHelper
{
    public sealed class Constants
    {
        public const string InstallAction = "Install";
        public const string UninstallAction = "Uninstall";
        public const string UpgradeAction = "Upgrade";
        public const string PackageSourcesSectionName = "packageSources";
        public const string ActivePackageSourcesSectionName = "activePackageSource";
        public const string NuGetOfficialSourceFeedName = "NuGet official package source";
        public const string NuGetOfficialSourceFeedValue = "https://nuget.org/api/v2/";
    }

    public static class ProjectTemplates
    {
        public static string MVC4RazorWebAppCsharp = "MvcWebApplicationProjectTemplate.11.cshtml.vstemplate";
        public static string MVC4CshtmlWebAppTemplateName = "MvcFacebookApplicationProjectTemplate.11.cshtml.vstemplate";
        public static string ConsoleAppCSharp = "csConsoleApplication.vstemplate";
        public static string WindowsFormAppCSharp = "csWindowsApplication.vstemplate";
        public static string PortableClassLibrary = "csPortableClassLibrary.vstemplate";
        public static string Windows8CSharp = "Microsoft.CS.WinRT.ClassLibrary";
        public static string SilverLight = "SilverlightClassLibrary.vstemplate";
        public static string WindowsPhone = "Discover.vstemplate";
        public static string CloudService = "CloudService_cs.zip";
    }

    public static class ProjectTargetFrameworks
    {
        public static string Net45 = "|$TargetFrameworkVersion$=4.5";
        public static string Net40 = "|$TargetFrameworkVersion$=4.0";
        public static string Net35 = "|$TargetFrameworkVersion$=3.5";
        public static string Net20 = "|$TargetFrameworkVersion$=2.0";
    }

    public static class ProjectTargetWindowsAzureTools
    {
        public static string Version24 = "|$TargetWindowsAzureToolsVersion$=2.4";
        public static string Version23 = "|$TargetWindowsAzureToolsVersion$=2.3";
        public static string Version22 = "|$TargetWindowsAzureToolsVersion$=2.2";
    }
}
