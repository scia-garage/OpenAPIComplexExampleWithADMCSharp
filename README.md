 # Prerequisities
- installed latest version of SCIA Engineer (19.1 patch 2)
- installed MS Visual Studio https://visualstudio.microsoft.com/cs/vs/ Community edition is sufficient. You can also use Visual Studio Code https://code.visualstudio.com/

# To start new C# project in MS Visual Studio to produce app that will use the Scia OpenAPI:
- create empty C# console application project with .NET 4.6.1
- Add reference to following dlls and edit properties of references and set Copy Local = False
   * SCIA.OpenAPI.dll located in SCIA Engineer install folder
   * ModelExchanger.AnalysisDataModel.dll located in subfolder OpenAPI_dll in SCIA Engineer install folder
   * ModelExchanger.AnalysisDataModel.Contracts.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * ModelExchanger.AnalysisDataModel.Implementation.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * ModelExchanger.Shared.Models.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.AdmToAdm.AdmSignalR.Models.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.Common.Contracts.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.Common.Implementation.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.Common.Models.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.ModelExchangerExtension.Contracts.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.ModelExchangerExtension.Integration.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
   * SciaTools.Kernel.ModelExchangerExtension.Models.dll located in sub folder OpenAPI_dll in SCIA Engineer install folder
- Create new / use configuration for x86 / x64 as needed according to SCIA Engineer Architecture
- write method for resolving of assemblies - see sample code
   ```C#
        private static void SciaOpenApiAssemblyResolve()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string dllName = args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll";
                string dllFullPath = Path.Combine(SciaEngineerFullPath, dllName);
                if (!File.Exists(dllFullPath))
                {
                    //return null;
                    dllFullPath = Path.Combine(SciaEngineerFullPath, "OpenAPI_dll", dllName);
                }
                if (!File.Exists(dllFullPath))
                {
                    return null;
                }
                return Assembly.LoadFrom(dllFullPath);
            };
        }
   ```
- write method for deliting temp folder
   ``` C#
        private static void DeleteTemp()
        {

            if (Directory.Exists(SciaEngineerTempPath)){
                Directory.Delete(SciaEngineerTempPath, true);
            }

        }
   ```
- Write your application that uses the SCIA.OpenAPI functions documentation of OpenAPI is [here](https://help.scia.net/api/19.1.1023/index.html). In this case you need to create Environment, RunSCIAEngineer and CreateAdmObject for ADM objects
- For methods and objects of so called ADM - documentation is [here](http://docs.calatrava.scia.net/html/9900284a-b4b6-405e-b157-630e24129caf.htm#!)
- Methods using OpenAPI have to run in single thread appartment use STAThread 
- Don't forget to use "using" statement for environment object creation OR call the Environment's Dispose() method when you finish your work with SCIA OpenAPI


# New project with Visual Studio Code
- read incrustions how to use this IDE here https://docs.microsoft.com/en-us/dotnet/core/tutorials/with-visual-studio-code

# To use this example application:
- clone this GIT repository
  - start command line
  - go to desired directory
  - write command "git clone <url_to_this_exmaple_app_repo_that_can_be_found_on_github_project_page>"
- open project in MS Visual Studio
- check correctness of path to referenced DLL located in your Scia Engineer install folder: SCIA.OpenAPI.dll (e.g. edit the .csproj)
- check correctness of paths in code poinitng to Scia Engineer install directory and template file
- select appropriate configuration (x86/x64) as needed according to SCIA Engineer Architecture
- build
- run

# Troubleshooting:
