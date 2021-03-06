﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ComponentModel.Composition;
using NuGet;
using NuGetTestHelper;
using System.ComponentModel.Composition.Hosting;
using System.Runtime.Versioning;

namespace NuGetTestHelper
{
    public class NuGetPackageManager
    {     
        public bool InstallPackage(string packageId, string packageVersion,out string installationMessage,bool updateAll=false)
        {           
            PackageId = packageId;
            PackageVersion = packageVersion;
            bool installed = vsProjectManager.InstallPackage(PackageId, PackageVersion,10 * 60 * 1000,updateAll);
            installationMessage = GetPackageInstallationOutput();
            return installed;
        }

        public bool InstallPackage(string packageFullPath, out string installationMessage, bool updateAll=false)
        {
            this.PackageFullPath = packageFullPath;
            ZipPackage zipPackage = new ZipPackage(PackageFullPath);
            PackageId = zipPackage.Id;
            PackageVersion = zipPackage.Version.ToString();
            bool installed = vsProjectManager.InstallPackage(PackageId, PackageVersion, 10 * 60 * 1000, updateAll);
            installationMessage = GetPackageInstallationOutput();
            return installed;
        }

        public string AnalyzePackage(string packageFullPath)
        {
            return NugetProcessUtility.AnalyzeNugetPackage(packageFullPath);
        }
                
        internal static NuGetPackageManager GetNuGetPackageManager(VsProjectManager projectManager)
        {
            return new NuGetPackageManager(projectManager);
        }

        private NuGetPackageManager(VsProjectManager projectManager)
        {
            vsProjectManager = projectManager;
        } 

        private static bool CheckForIndexRange(int index, string output)
        {
            if (index == -1 || index >= output.Length)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private string GetPackageInstallationOutput()
        {
            string output = vsProjectManager.GetPackageInstallationOutput();
            //return output;
            string installCommand = "Install.ps1" + @"""" + " " + @"""" + PackageId;
            int indexOfInstallCommand = output.IndexOf(installCommand,StringComparison.OrdinalIgnoreCase);

            if (!CheckForIndexRange(indexOfInstallCommand, output))
            {
                return null;
            }

            int endindexOfInstallCommand = output.IndexOf("PM>", indexOfInstallCommand,StringComparison.OrdinalIgnoreCase);
            if (!CheckForIndexRange(endindexOfInstallCommand, output))
            {
                return null;
            }
            string installCommandOutput = output.Substring(indexOfInstallCommand, endindexOfInstallCommand - indexOfInstallCommand);
            return installCommandOutput;
        }

        private string PackageFullPath = string.Empty;
        private VsProjectManager vsProjectManager;
        private string PackageId;
        private string PackageVersion;
    }
}
