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
- Write your application that use the SCIA.OpenAPI functions
- Methods using OpenAPI have to run in single thread appartment use STAThread 
- Don't forget to use "using" statement for environment object creation OR call the Environment's Dispose() method when you finish your work with SCIA OpenAPI

