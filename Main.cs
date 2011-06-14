//-----------------------------------------------------------------------
// <copyright company="CoApp Project">
//     Copyright (c) 2011 Eric Schultz. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace CoApp.VCMsiCreate
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using Toolkit.Exceptions;
    using Toolkit.Extensions;
    using Toolkit.Scripting.Languages.PropertySheet;
    using Toolkit.Utility;
    using CoApp.Toolkit.Crypto;
    using System.Xml.Linq;
    using VCMsiCreate.Properties;
    using CoApp.Toolkit.Package;
   

    internal class VCMsiCreateMain
    {
        private const string help =
            @"
Usage:
-------

VCMsiCreate [options] 
    
    Options:
    --------
    --help                      this help
    --nologo                    don't display the logo
    --load-config=<file>        loads configuration from <file>
    --verbose                   prints verbose messages
    --cert-file=<file>          pfx certificate file
    --cert-pass=<password>      password to open pfx file
    --msm-dir=<dir>             path to the directory holding the merge modules for Visual C++
                                (defaults to C:\Program Files\Common Files\Merge Modules or
                                C:\Program Files (x86)\etc...)
    --version=<version>         The version of C++ you want to create MSI's for, such as 8.0, 9.0, 10.0, etc.
";
        private bool verbose = false;
        private ProcessUtility _candle;
        private ProcessUtility _light;
        private ProcessUtility _signTool;
        private string certFile;
        private string certPass;
        private string msmDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Common Files","Merge Modules");
        private string version;


        /// <summary>
        /// Entry Point
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static int Main(string[] args)
        {
            return new VCMsiCreateMain().main(args);
        }



        private int main(string[] args)
        {
            var options = args.Switches();
            var parameters = args.Parameters();



            #region Parse Options

            try
            {
                foreach (string arg in options.Keys)
                {
                    IEnumerable<string> argumentParameters = options[arg];

                    switch (arg)
                    {
                        /* global switches */
                        case "load-config":
                            // all ready done, but don't get too picky.
                            break;
                        case "nologo":
                            this.Assembly().SetLogo("");
                            break;

                        case "verbose":
                            verbose = true;
                            break;

                        case "help":
                            return Help();

                        case "cert-file":
                            certFile = argumentParameters.First();
                            break;
                        case "cert-pass":
                            certPass = argumentParameters.First();
                            break;
                        case "msm-dir":
                            msmDir = argumentParameters.First();
                            break;
                        case "version":
                            version = argumentParameters.First();
                            break;
                    }
                }



                Logo();



            #endregion

                
                _candle = new ProcessUtility(ProgramFinder.ProgramFilesAndDotNet.ScanForFile("candle.exe"));
                _light = new ProcessUtility(ProgramFinder.ProgramFilesAndDotNet.ScanForFile("light.exe"));

                _signTool = new ProcessUtility(ProgramFinder.ProgramFilesAndDotNet.ScanForFile("signTool.exe"));
                
                var pfxStore = PfxStoreLoader.Load(certFile, certPass);
                if (pfxStore != null)
                {
                    var codeSigningCert = pfxStore.FindCodeSigningCert();
                    if (codeSigningCert == null)
                        return Fail("No code signing certificate found.");
                }
                else
                {
                    return Fail("Couldn't open the certificate file.");
                }

                var files = Directory.EnumerateFiles(msmDir, "*.msm", SearchOption.TopDirectoryOnly);
                var groups = from f in files
                             let fi = new FileInfo(f)
                             where fi.Name.StartsWith("Microsoft") && !fi.Name.Contains("Debug")
                             let version = new String(fi.Name.Split('_')[1].Substring(2).TakeAllBut(1).ToArray())
                             group f by version into versions
                             select new
                             {
                                 Version = versions.Key,
                                 ArchGroups = (from a in versions

                                               let possArch = new String(a.TakeAllBut(4).TakeFromEnd(3).ToArray())

                                               group a by possArch into archs
                                               select new
                                               {
                                                   Arch = archs.Key,
                                                   Files = archs
                                               })

                             };

                foreach (var g in groups)
                {
                    var vers = g.Version;

                    if (version == null || version.ExtendVersion() == vers.ExtendVersion())
                    {
                        
                        foreach (var a in g.ArchGroups)
                        {
                            Console.WriteLine("Building {0} for {1}".format(vers, a.Arch));
                            var outputFile = "{0}-{1}-{2}.msi".format("VC", vers.ExtendVersion(), a.Arch);
                            XDocument wix;
                            var huid = new Huid("Visual C/C++ Runtime Library", vers.ExtendVersion(), a.Arch, "");

                            using (var reader = new StringReader(Resources.WixTemplate))
                            {
                                wix = XDocument.Load(reader);
                            }

                            var productTag = wix.Descendants("Product").First();

                            var targetDir = (from d in productTag.Descendants()
                                             where d.Name == "Directory" &&
                                                 d.Attribute("Id").Value == "TARGETDIR"
                                             select d).First();

                            productTag.SetAttributeValue("Name", "Visual C/C++ Runtime Library");
                            productTag.SetAttributeValue("Version", vers.ExtendVersion());
                            productTag.SetAttributeValue("Manufacturer", "Microsoft");
                            productTag.SetAttributeValue("Id", huid.ToString());


                            int counter = 0;
                            var ids = new List<int>();
                            foreach (var f in a.Files)
                            {
                                targetDir.Add(
                                    new XElement("Merge",
                                        new XAttribute("Id", "VC" + counter),
                                        new XAttribute("SourceFile", f),
                                        new XAttribute("DiskId", "1"),
                                        new XAttribute("Language", "0")));
                                ids.Add(counter);
                                counter++;
                            }

                            var feature = productTag.Descendants("Feature").First();
                            foreach (var i in ids)
                            {
                                feature.Add(
                                    new XElement("MergeRef",
                                        new XAttribute("Id", "VC" + i)));
                            }

                            var packageTag = productTag.Descendants("Package").First();
                            packageTag.SetAttributeValue("Platform", a.Arch);

                            //add the Wix namespace
                            XNamespace wixNS = "http://schemas.microsoft.com/wix/2006/wi";

                            foreach (var n in wix.Descendants())
                            {
                                n.Name = wixNS + n.Name.LocalName;
                            }

                            var tempPrefix = Path.GetTempFileName();
                            File.WriteAllText(tempPrefix + ".wxs", wix.ToString());

                            if (_candle.Exec("-nologo -sw1075 -out {0}.wixobj {0}.wxs", tempPrefix) != 0)
                                return Fail(_candle.StandardOut);


                            //we suppress the lack of UpgradeCode warning since we don't use them 
                            if (_light.Exec("-nologo -sw1076 -out {1} {0}.wixobj", tempPrefix, outputFile) != 0)
                                return Fail(_light.StandardOut);

                            if (!signFile(outputFile))
                                return Fail("Couldn't sign file");

                        }
                    }
                   
                }
                             
                        
                            

            }
            catch (ConsoleException e)
            {
                return Fail("   {0}", e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return Fail("   {0}", e.Message);
            }

            return 0;
        }

        

        #region fail/help/logo

        public static int Fail(string text, params object[] par)
        {
            Logo();
            using (new ConsoleColors(ConsoleColor.Red, ConsoleColor.Black))
            {
                Console.WriteLine("Error:{0}", text.format(par));
            }
            return 1;
        }

        private static int Help()
        {
            Logo();
            using (new ConsoleColors(ConsoleColor.White, ConsoleColor.Black))
            {
                help.Print();
            }
            return 0;
        }

        private static void Logo()
        {
            using (new ConsoleColors(ConsoleColor.Cyan, ConsoleColor.Black))
            {
                Assembly.GetEntryAssembly().Logo().Print();
            }
            Assembly.GetEntryAssembly().SetLogo("");
        }

        #endregion


        bool signFile(string filename)
        {
            bool signed = false;
            foreach (var s in PfxStoreLoader.TimestampServers)
            {
                for (int i = 0; i < 5; i++)
                {
                    signed = (_signTool.
                        Exec(@"sign /v /t {3} /f ""{0}"" /p ""{1}"" ""{2}""",
                            certFile, certPass, filename, s) == 0);
                    if (signed)
                        break;
                }
                if (signed)
                    break;
            }
            if (!signed)
            {

                return false;
            }
            return true;
        }
    }
}