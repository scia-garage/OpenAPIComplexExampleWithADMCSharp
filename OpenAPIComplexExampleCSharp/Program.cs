using ModelExchanger.AnalysisDataModel;
using ModelExchanger.AnalysisDataModel.Enums;
using ModelExchanger.AnalysisDataModel.Libraries;
using ModelExchanger.AnalysisDataModel.Loads;
using ModelExchanger.AnalysisDataModel.Subtypes;
using ModelExchanger.AnalysisDataModel.StructuralElements;
using ModelExchanger.AnalysisDataModel.StructuralReferences.Curves;
using ModelExchanger.AnalysisDataModel.StructuralReferences.Points;
using SciaTools.Kernel.ModelExchangerExtension.Models.Exchange;
using System;
using SCIA.OpenAPI.Results;
using Results64Enums;
using System.IO;
using System.Reflection;
using SCIA.OpenAPI.Utils;
using SciaTools.AdmToAdm.AdmSignalR.Models.ModelModification;
using System.Collections.Generic;
using SciaTools.Kernel.Common.Implementation.App;
using SciaTools.Kernel.ModelExchangerExtension.Integration.Modules;
using SciaTools.Kernel.ModelExchangerExtension.Contracts.AnalysisModelModifications;
using SciaTools.Kernel.ModelExchangerExtension.Models.AnalysisModelModifications;
using SciaTools.Kernel.ModelExchangerExtension.Contracts.Services;

namespace OpenAPIAndADMDemo
{
    class Program
    {
        private static string GetAppPath()
        {
            //var directory = new DirectoryInfo(Environment.CurrentDirectory);
            //return directory.Parent.FullName;
            return @"c:\WORK\SCIA-ENGINEER\TESTING-VERSIONS\Full_19.1.2010.64_rel_19.1_patch_2_x64\"; // SEn application installation folder, don't forget run "EP_regsvr32 esa.exe" from commandline with Admin rights
        }

        /// <summary>
        /// Path to Scia engineer
        /// </summary>
        static private string SciaEngineerFullPath => GetAppPath();


        /// <sumamary>
        /// Path to SCIA Engineer temp
        /// </sumamary>
        static private string SciaEngineerTempPath => GetTempPath();

        private static string GetTempPath()
        {
            return @"c:\WORK\SCIA-ENGINEER\TESTING-VERSIONS\Full_19.1.2010.64_rel_19.1_patch_2_x64\Temp\"; // Must be SEn application temp path, run SEn and go to menu: Setup -> Options -> Directories -> Temporary files
        }

        static private string SciaEngineerProjecTemplate => GetTemplatePath();

        private static string GetTemplatePath()
        {
            //Open project in SCIA Engineer on specified path
            string MyAppPath = AppDomain.CurrentDomain.BaseDirectory;
            //return Path.Combine(MyAppPath, @"..\..\..\..\res\OpenAPIEmptyProject.esa");//path to teh empty SCIA Engineer project
            return @"C:\WORK\SourceCodes\OpenAPIComplexExampleWithADMCSharp\res\OpenAPIEmptyProject.esa";
        }

        static private string AppLogPath => GetThisAppLogPath();

        static private string GetThisAppLogPath()
        {
            return @"c:\TEMP\OpenAPI\MyLogsTemp"; // Folder for storing of log files for this console application
        }

        private static void DeleteTemp()
        {

            if (Directory.Exists(SciaEngineerTempPath))
            {
                Directory.Delete(SciaEngineerTempPath, true);
            }

        }

        /// <summary>
        /// Assembly resolve method has to be call here
        /// </summary>
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
        static void RunSCIAOpenAPI()
        {
            using (SCIA.OpenAPI.Environment env = new SCIA.OpenAPI.Environment(SciaEngineerFullPath, AppLogPath, "1.0.0.0"))//path to the location of your installation and temp path for logs)
            {
                //Run SCIA Engineer application
                bool openedSE = env.RunSCIAEngineer(SCIA.OpenAPI.Environment.GuiMode.ShowWindowShow);
                if (!openedSE)
                {
                    return;
                }
                Console.WriteLine($"SEn opened");
                SciaFileGetter fileGetter = new SciaFileGetter();
                var EsaFile = fileGetter.PrepareBasicEmptyFile(@"C:/TEMP/");//path where the template file "template.esa" is created
                if (!File.Exists(EsaFile))
                {
                    throw new InvalidOperationException($"File from manifest resource is not created ! Temp: {env.AppTempPath}");
                }

                SCIA.OpenAPI.EsaProject proj = env.OpenProject(EsaFile);
                //SCIA.OpenAPI.EsaProject proj = env.OpenProject(SciaEngineerProjecTemplate);
                if (proj == null)
                {
                    return;
                }
                Console.WriteLine($"Proj opened");


                // info about Project 
                ProjectInformation projectInformation = new ProjectInformation(Guid.NewGuid(), "ProjectX")
                {
                    BuildingType = "SimpleFrame",
                    Location = "39XG+P7 Praha",
                    LastUpdate = DateTime.Today,
                    Status = "Draft",
                    ProjectType = "New construction"
                };

                // info about Model ModelExchanger.AnalysisDataModel.ModelInformation
                ModelInformation modelInformation = new ModelInformation(Guid.NewGuid(), "ModelOne")
                {
                    Discipline = "Static",
                    Owner = "JB",
                    LevelOfDetail = "200",
                    LastUpdate = DateTime.Today,
                    SourceApplication = "OpenAPI",
                    RevisionNumber = "1",
                    SourceCompany = "SCIA",
                    SystemOfUnits = SystemOfUnits.Metric

                };

                Console.WriteLine($"Set grade for concrete material: ");
                string conMatGrade = Console.ReadLine();

                StructuralMaterial concrete = new StructuralMaterial(Guid.NewGuid(), "Concrete", MaterialType.Concrete, conMatGrade);


                Console.WriteLine($"Set grade for steel material: ");
                string steelMatGrade = Console.ReadLine();
                StructuralMaterial steel = new StructuralMaterial(Guid.NewGuid(), "Steel", MaterialType.Steel, steelMatGrade);
                ResultOfPartialAddToAnalysisModel addResult = proj.Model.CreateAdmObject(concrete, steel);
                bool isMessageSent = addResult.IsMessageSendResult;
                AdmChangeStatus status = addResult.PartialAddResult.Status;

                Console.WriteLine($"Materials created in ADM");

                //Create cross-sections in local ADM
                Console.WriteLine($"Set steel profile: ");
                string steelProfile = Console.ReadLine();

                StructuralCrossSection steelprofile = new StructuralManufacturedCrossSection(Guid.NewGuid(), steelProfile, steel, steelProfile, FormCode.ISection, DescriptionId.EuropeanIBeam);

                addResult = proj.Model.CreateAdmObject(steelprofile);

                Console.WriteLine($"Set height of concrete rectangle in mm: ");
                double heigth = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine($"Set width of concrete rectangle in mm: ");
                double width = Convert.ToDouble(Console.ReadLine());
                StructuralCrossSection concreteRectangle = new StructuralParametricCrossSection(Guid.NewGuid(), "Concrete", concrete, ProfileLibraryId.Rectangle, new UnitsNet.Length[2] { UnitsNet.Length.FromMillimeters(heigth), UnitsNet.Length.FromMillimeters(width) });
                addResult = proj.Model.CreateAdmObject(concreteRectangle);
                Console.WriteLine($"CSSs created in ADM");

                Console.WriteLine($"Set parameter a: ");
                double a = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine($"Set parameter b: ");
                double b = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine($"Set parameter c: ");
                double c = Convert.ToDouble(Console.ReadLine());


                StructuralPointConnection N1 = new StructuralPointConnection(Guid.NewGuid(), "N1", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N1);
                StructuralPointConnection N2 = new StructuralPointConnection(Guid.NewGuid(), "N2", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N2);
                StructuralPointConnection N3 = new StructuralPointConnection(Guid.NewGuid(), "N3", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N3);
                StructuralPointConnection N4 = new StructuralPointConnection(Guid.NewGuid(), "N4", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N4);
                StructuralPointConnection N5 = new StructuralPointConnection(Guid.NewGuid(), "N5", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N5);
                StructuralPointConnection N6 = new StructuralPointConnection(Guid.NewGuid(), "N6", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N6);
                StructuralPointConnection N7 = new StructuralPointConnection(Guid.NewGuid(), "N7", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N7);
                StructuralPointConnection N8 = new StructuralPointConnection(Guid.NewGuid(), "N8", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N8);

                var beamB1lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N5 }) };
                StructuralCurveMember B1 = new StructuralCurveMember(Guid.NewGuid(), "B1", beamB1lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                    Layer = "Columns",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B1);
                var beamB2lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N2, N6 }) };
                StructuralCurveMember B2 = new StructuralCurveMember(Guid.NewGuid(), "B2", beamB2lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                    Layer = "Columns",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B2);
                var beamB3lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N7 }) };
                StructuralCurveMember B3 = new StructuralCurveMember(Guid.NewGuid(), "B3", beamB3lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                    Layer = "Columns",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B3);
                var beamB4lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N8 }) };
                StructuralCurveMember B4 = new StructuralCurveMember(Guid.NewGuid(), "B4", beamB4lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                    Layer = "Columns",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B4);
                var beamB5lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N5, N6 }) };
                StructuralCurveMember B5 = new StructuralCurveMember(Guid.NewGuid(), "B5", beamB5lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                    Layer = "Beams",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B5);
                var beamB6lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N6, N7 }) };
                StructuralCurveMember B6 = new StructuralCurveMember(Guid.NewGuid(), "B6", beamB6lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                    Layer = "Beams",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B6);
                var beamB7lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N8 }) };
                StructuralCurveMember B7 = new StructuralCurveMember(Guid.NewGuid(), "B7", beamB7lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                    Layer = "Beams",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B7);
                var beamB8lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N5 }) };
                StructuralCurveMember B8 = new StructuralCurveMember(Guid.NewGuid(), "B8", beamB8lines, steelprofile)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                    Layer = "Beams",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B8);

                Constraint<UnitsNet.RotationalStiffness?> FreeRotation = new Constraint<UnitsNet.RotationalStiffness?>(ConstraintType.Free, UnitsNet.RotationalStiffness.FromKilonewtonMetersPerRadian(0));
                Constraint<UnitsNet.RotationalStiffness?> FixedRotation = new Constraint<UnitsNet.RotationalStiffness?>(ConstraintType.Rigid, UnitsNet.RotationalStiffness.FromKilonewtonMetersPerRadian(1e+10));
                Constraint<UnitsNet.ForcePerLength?> FixedTranslation = new Constraint<UnitsNet.ForcePerLength?>(ConstraintType.Rigid, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(1e+10));
                Constraint<UnitsNet.ForcePerLength?> FreeTranslation = new Constraint<UnitsNet.ForcePerLength?>(ConstraintType.Free, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(0));
                Constraint<UnitsNet.Pressure?> FixedTranslationLine = new Constraint<UnitsNet.Pressure?>(ConstraintType.Rigid, UnitsNet.Pressure.FromKilopascals(1e+10));
                Constraint<UnitsNet.RotationalStiffnessPerLength?> FreeRotationLine = new Constraint<UnitsNet.RotationalStiffnessPerLength?>(ConstraintType.Free, UnitsNet.RotationalStiffnessPerLength.FromKilonewtonMetersPerRadianPerMeter(0));

                StructuralPointSupport PS1 = new StructuralPointSupport(Guid.NewGuid(), "SPS1", N1)
                {
                    RotationX = FreeRotation,
                    RotationY = FixedRotation,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation
                };

                StructuralPointSupport PS2 = new StructuralPointSupport(Guid.NewGuid(), "SPS2", N2)
                {
                    RotationX = FreeRotation,
                    RotationY = FixedRotation,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation
                };
                StructuralPointSupport PS3 = new StructuralPointSupport(Guid.NewGuid(), "SPS3", N3)
                {
                    RotationX = FreeRotation,
                    RotationY = FixedRotation,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation
                };
                StructuralPointSupport PS4 = new StructuralPointSupport(Guid.NewGuid(), "SPS4", N4)
                {
                    RotationX = FreeRotation,
                    RotationY = FixedRotation,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation
                };
                addResult = proj.Model.CreateAdmObject(PS1, PS2, PS3, PS4);

                Console.WriteLine($"Set thickness of the slab: ");
                double thickness = Convert.ToDouble(Console.ReadLine());
                var edgecurves = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N5, N6 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N6, N7 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N8 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N5 })
                };
                StructuralSurfaceMember S1 = new StructuralSurfaceMember(Guid.NewGuid(), "S1", edgecurves, concrete, UnitsNet.Length.FromMeters(thickness))
                {
                    Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Plate),
                    Behaviour = Member2DBehaviour.Isotropic,
                    Alignment = Member2DAlignment.Centre,
                    EccentricityEz = UnitsNet.Length.FromMeters(0),
                    Shape = Member2DShape.Flat
                };
                addResult = proj.Model.CreateAdmObject(S1);

                Console.WriteLine($"Set length of opening in slab  in m: ");
                double lengthOpening = Convert.ToDouble(Console.ReadLine());
                Console.WriteLine($"Set width of opening in slab  in m: ");
                double withOpening = Convert.ToDouble(Console.ReadLine());

                StructuralPointConnection N9 = new StructuralPointConnection(Guid.NewGuid(), "N9", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N9);
                StructuralPointConnection N10 = new StructuralPointConnection(Guid.NewGuid(), "N10", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N10);
                StructuralPointConnection N11 = new StructuralPointConnection(Guid.NewGuid(), "N11", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N11);
                StructuralPointConnection N12 = new StructuralPointConnection(Guid.NewGuid(), "N12", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
                addResult = proj.Model.CreateAdmObject(N12);

                var openingEdges = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N9, N10 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N10, N11 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N11, N12 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N12, N9 })
                };
                StructuralSurfaceMemberOpening O1S1 = new StructuralSurfaceMemberOpening(Guid.NewGuid(), "O1", S1, openingEdges);
                addResult = proj.Model.CreateAdmObject(O1S1);


                var edgecurvesS2 = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N2 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N2, N3 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N4 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N1 })
                };
                StructuralSurfaceMember S2 = new StructuralSurfaceMember(Guid.NewGuid(), "S2", edgecurvesS2, concrete, UnitsNet.Length.FromMeters(thickness))
                {
                    Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Plate),
                    Behaviour = Member2DBehaviour.Isotropic,
                    Alignment = Member2DAlignment.Centre,
                    EccentricityEz = UnitsNet.Length.FromMeters(0),
                    Shape = Member2DShape.Flat
                };
                addResult = proj.Model.CreateAdmObject(S2);

                StructuralPointConnection N13 = new StructuralPointConnection(Guid.NewGuid(), "N13", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N13);
                StructuralPointConnection N14 = new StructuralPointConnection(Guid.NewGuid(), "N14", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N14);
                StructuralPointConnection N15 = new StructuralPointConnection(Guid.NewGuid(), "N15", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N15);
                StructuralPointConnection N16 = new StructuralPointConnection(Guid.NewGuid(), "N16", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N16);

                var regionEdges = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N13, N14 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N14, N15 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N15, N16 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N16, N13 })
                };
                StructuralSurfaceMemberRegion SMR = new StructuralSurfaceMemberRegion(Guid.NewGuid(), "Region", S2, regionEdges, concrete)
                {
                    Thickness = UnitsNet.Length.FromMeters(2 * thickness),
                    EccentricityEz = UnitsNet.Length.FromMeters(0),
                    Alignment = Member2DAlignment.Centre
                };
                addResult = proj.Model.CreateAdmObject(SMR);

                Subsoil subsoil = new Subsoil("Subsoil", UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(80.5), UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(35.5), UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(50), UnitsNet.ForcePerLength.FromMeganewtonsPerMeter(15.5), UnitsNet.ForcePerLength.FromMeganewtonsPerMeter(10.2));
                StructuralSurfaceConnection SS1 = new StructuralSurfaceConnection(Guid.NewGuid(), "SS1", S2, subsoil);
                addResult = proj.Model.CreateAdmObject(SS1);

                var edgecurvesS3 = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N4 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N8 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N7 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N3 })
                };
                StructuralSurfaceMember S3 = new StructuralSurfaceMember(Guid.NewGuid(), "S3", edgecurvesS3, concrete, UnitsNet.Length.FromMeters(thickness))
                {
                    Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Wall),
                    Behaviour = Member2DBehaviour.Isotropic,
                    Alignment = Member2DAlignment.Centre,
                    EccentricityEz = UnitsNet.Length.FromMeters(0),
                    Shape = Member2DShape.Flat
                };
                addResult = proj.Model.CreateAdmObject(S3);


                RelConnectsSurfaceEdge LH1 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH1", S3, 0)
                {
                    StartPointRelative = 0,
                    EndPointRelative = 1,
                    TranslationX = FixedTranslationLine,
                    TranslationY = FixedTranslationLine,
                    TranslationZ = FixedTranslationLine,
                    RotationX = FreeRotationLine,
                };
                addResult = proj.Model.CreateAdmObject(LH1);
                RelConnectsSurfaceEdge LH2 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH2", S3, 1)
                {
                    StartPointRelative = 0,
                    EndPointRelative = 1,
                    TranslationX = FixedTranslationLine,
                    TranslationY = FixedTranslationLine,
                    TranslationZ = FixedTranslationLine,
                    RotationX = FreeRotationLine,
                };
                addResult = proj.Model.CreateAdmObject(LH2);
                RelConnectsSurfaceEdge LH3 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH3", S3, 2)
                {
                    StartPointRelative = 0,
                    EndPointRelative = 1,
                    TranslationX = FixedTranslationLine,
                    TranslationY = FixedTranslationLine,
                    TranslationZ = FixedTranslationLine,
                    RotationX = FreeRotationLine,
                };
                addResult = proj.Model.CreateAdmObject(LH3);
                RelConnectsSurfaceEdge LH4 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH4", S3, 3)
                {
                    StartPointRelative = 0,
                    EndPointRelative = 1,
                    TranslationX = FixedTranslationLine,
                    TranslationY = FixedTranslationLine,
                    TranslationZ = FixedTranslationLine,
                    RotationX = FreeRotationLine,
                };
                addResult = proj.Model.CreateAdmObject(LH4);


                RelConnectsStructuralMember H1 = new RelConnectsStructuralMember(Guid.NewGuid(), "H1", B1)
                {
                    Position = Position.End,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation,
                    RotationX = FreeRotation,
                };
                addResult = proj.Model.CreateAdmObject(H1);
                RelConnectsStructuralMember H2 = new RelConnectsStructuralMember(Guid.NewGuid(), "H2", B2)
                {
                    Position = Position.End,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation,
                    RotationX = FreeRotation,
                    RotationY = FreeRotation,
                    RotationZ = FreeRotation
                };
                addResult = proj.Model.CreateAdmObject(H2);
                RelConnectsStructuralMember H3 = new RelConnectsStructuralMember(Guid.NewGuid(), "H3", B3)
                {
                    Position = Position.End,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation,
                    RotationX = FreeRotation,
                    RotationY = FreeRotation,
                    RotationZ = FreeRotation
                };
                addResult = proj.Model.CreateAdmObject(H3);
                RelConnectsStructuralMember H4 = new RelConnectsStructuralMember(Guid.NewGuid(), "H4", B4)
                {
                    Position = Position.End,
                    TranslationX = FixedTranslation,
                    TranslationY = FixedTranslation,
                    TranslationZ = FixedTranslation,
                    RotationX = FreeRotation,
                    RotationY = FreeRotation,
                    RotationZ = FreeRotation
                };
                addResult = proj.Model.CreateAdmObject(H4);

                StructuralPointConnection N17 = new StructuralPointConnection(Guid.NewGuid(), "N17", UnitsNet.Length.FromMeters(-1 * b), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
                addResult = proj.Model.CreateAdmObject(N17);

                var beamB9lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N17 }) };
                StructuralCurveMember B9 = new StructuralCurveMember(Guid.NewGuid(), "B9", beamB9lines, concreteRectangle)
                {
                    Behaviour = CurveBehaviour.Standard,
                    SystemLine = CurveAlignment.Centre,
                    Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                    Layer = "Beams",
                    EccentricityEy = UnitsNet.Length.FromMeters(0),
                    EccentricityEz = UnitsNet.Length.FromMeters(0)
                };
                addResult = proj.Model.CreateAdmObject(B9);

                StructuralCurveConnection LSB = new StructuralCurveConnection(Guid.NewGuid(), "LSB", B9)
                {
                    Origin = Origin.FromStart,
                    CoordinateDefinition = CoordinateDefinition.Relative,
                    StartPointRelative = 0.25,
                    EndPointRelative = 0.75
                };
                addResult = proj.Model.CreateAdmObject(LSB);


                StructuralLoadGroup LG1 = new StructuralLoadGroup(Guid.NewGuid(), "LG1", LoadGroupType.Variable)
                {
                    Load = new CSInfrastructure.FlexibleEnum<Load>(Load.Domestic)
                };
                addResult = proj.Model.CreateAdmObject(LG1);


                StructuralLoadCase LC1 = new StructuralLoadCase(Guid.NewGuid(), "LC1", ActionType.Variable, LG1, LoadCaseType.Static)
                {
                    Duration = Duration.Long,
                    Specification = Specification.Standard
                };
                addResult = proj.Model.CreateAdmObject(LC1);

                Console.WriteLine($"Set value of line load on  kN/m: ");
                double lineloadValue = Convert.ToDouble(Console.ReadLine());


                StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB1 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB1", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B1))
                {
                    Direction = ActionDirection.Y,
                    Distribution = CurveDistribution.Uniform
                };
                StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB2 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB2", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B2))
                {
                    Direction = ActionDirection.Y,
                    Distribution = CurveDistribution.Uniform
                };
                StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB3 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB3", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B3))
                {
                    Direction = ActionDirection.Y,
                    Distribution = CurveDistribution.Uniform
                };
                StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB4 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB4", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B4))
                {
                    Direction = ActionDirection.Y,
                    Distribution = CurveDistribution.Uniform
                };

                addResult = proj.Model.CreateAdmObject(lineloadB1, lineloadB2, lineloadB3, lineloadB4);

                StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E1 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E1", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 0))
                {
                    Direction = ActionDirection.Z
                };
                StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E2 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E2", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 1))
                {
                    Direction = ActionDirection.Z
                };
                StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E3 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E3", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 2))
                {
                    Direction = ActionDirection.Z
                };
                StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E4 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E4", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 3))
                {
                    Direction = ActionDirection.Z
                };

                addResult = proj.Model.CreateAdmObject(edgeloadS1E1, edgeloadS1E2, edgeloadS1E3, edgeloadS1E4);


                StructuralLoadCase LC2 = new StructuralLoadCase(Guid.NewGuid(), "LC2", ActionType.Variable, LG1, LoadCaseType.Static)
                {
                    Duration = Duration.Long,
                    Specification = Specification.Standard
                };
                addResult = proj.Model.CreateAdmObject(LC2);

                Console.WriteLine($"Set value of surface load on the slab in kN/m^2: ");
                double surfaceloadValue = Convert.ToDouble(Console.ReadLine());


                StructuralSurfaceAction sls1 = new StructuralSurfaceAction(Guid.NewGuid(), "sls1", UnitsNet.Pressure.FromKilonewtonsPerSquareMeter(surfaceloadValue), S1, LC2)
                {
                    Direction = ActionDirection.Z,
                    Location = Location.Length
                };

                addResult = proj.Model.CreateAdmObject(sls1);

                StructuralPointAction<PointStructuralReferenceOnPoint> FP = new StructuralPointAction<PointStructuralReferenceOnPoint>(Guid.NewGuid(), "FP", UnitsNet.Force.FromKilonewtons(150), LC2, PointForceAction.InNode, new PointStructuralReferenceOnPoint(N13))
                {
                    Direction = ActionDirection.Z
                };
                addResult = proj.Model.CreateAdmObject(FP);

                var Combinations = new StructuralLoadCombinationData[2] { new StructuralLoadCombinationData(LC1, 1.0, 1.5), new StructuralLoadCombinationData(LC2, 1.0, 1.35) };
                StructuralLoadCombination C1 = new StructuralLoadCombination(Guid.NewGuid(), "C1", LoadCaseCombinationCategory.AccordingNationalStandard, Combinations)
                {
                    NationalStandard = LoadCaseCombinationStandard.EnUlsSetC
                };
                addResult = proj.Model.CreateAdmObject(C1);

                proj.Model.RefreshModel_ToSCIAEngineer();

                Console.WriteLine($"My model sent to SEn");



                // Run calculation
                proj.RunCalculation();
                Console.WriteLine($"My model calculate");

                //Initialize Results API
                ResultsAPI rapi = proj.Model.InitializeResultsAPI();
                if (rapi != null)
                {
                    //Create container for 1D results
                    Result IntFor1Db1 = new Result();
                    //Results key for internal forces on beam 1
                    ResultKey keyIntFor1Db1 = new ResultKey
                    {
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        EntityType = eDsElementType.eDsElementType_Beam,
                        EntityName = B1.Name,
                        Dimension = eDimension.eDim_1D,
                        ResultType = eResultType.eFemBeamInnerForces,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    };
                    //Load 1D results based on results key
                    IntFor1Db1 = rapi.LoadResult(keyIntFor1Db1);
                    if (IntFor1Db1 != null)
                    {
                        Console.WriteLine(IntFor1Db1.GetTextOutput());
                        var N = IntFor1Db1.GetMagnitudeName(0);
                        var Nvalue = IntFor1Db1.GetValue(0, 0);
                        Console.WriteLine(N);
                        Console.WriteLine(Nvalue);
                    }
                    //combination
                    //Create container for 1D results
                    Result IntFor1Db1Combi = new Result();
                    //Results key for internal forces on beam 1
                    ResultKey keyIntFor1Db1Combi = new ResultKey
                    {
                        EntityType = eDsElementType.eDsElementType_Beam,
                        EntityName = B1.Name,
                        CaseType = eDsElementType.eDsElementType_Combination,
                        CaseId = C1.Id,
                        Dimension = eDimension.eDim_1D,
                        ResultType = eResultType.eFemBeamInnerForces,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    };
                    // Load 1D results based on results key
                    IntFor1Db1Combi = rapi.LoadResult(keyIntFor1Db1Combi);
                    if (IntFor1Db1Combi != null)
                    {
                        // Console.WriteLine(IntFor1Db1Combi.GetTextOutput());
                    }
                    //Results key for reaction on node 1
                    ResultKey keyReactionsSu1 = new ResultKey
                    {
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        EntityType = eDsElementType.eDsElementType_Node,
                        EntityName = N1.Name,
                        Dimension = eDimension.eDim_reactionsPoint,
                        ResultType = eResultType.eReactionsNodes,
                        CoordSystem = eCoordSystem.eCoordSys_Global
                    };
                    Result reactionsSu1 = new Result();
                    reactionsSu1 = rapi.LoadResult(keyReactionsSu1);
                    if (reactionsSu1 != null)
                    {
                        Console.WriteLine(reactionsSu1.GetTextOutput());
                    }

                    Result Def2Ds1 = new Result();
                    // Results key for internal forces on slab
                    ResultKey keyDef2Ds1 = new ResultKey
                    {
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        EntityType = eDsElementType.eDsElementType_Slab,
                        EntityName = S1.Name,
                        Dimension = eDimension.eDim_2D,
                        ResultType = eResultType.eFemDeformations,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    };

                    Def2Ds1 = rapi.LoadResult(keyDef2Ds1);
                    if (Def2Ds1 != null)
                    {
                        Console.WriteLine(Def2Ds1.GetTextOutput());

                        double maxvalue = 0;
                        double pivot;
                        for (int i = 0; i < Def2Ds1.GetMeshElementCount(); i++)
                        {
                            pivot = Def2Ds1.GetValue(2, i);
                            if (System.Math.Abs(pivot) > System.Math.Abs(maxvalue))
                            {
                                maxvalue = pivot;

                            }
                        }
                        Console.WriteLine("Maximum deformation on slab:");
                        Console.WriteLine(maxvalue);
                    }

                }
                Console.WriteLine($"Press key to exit");
                Console.ReadKey();
                proj.CloseProject(SCIA.OpenAPI.SaveMode.SaveChangesNo);
                env.Dispose();
            }
        }
        static private object SciaOpenApiWorker(SCIA.OpenAPI.Environment env)
        {
           

            //Run SCIA Engineer application
            bool openedSE = env.RunSCIAEngineer(SCIA.OpenAPI.Environment.GuiMode.ShowWindowShow);
            if (!openedSE)
            {
                throw new InvalidOperationException($"SCIA Engineer not started");
            }
            Console.WriteLine($"SEn opened");
            SciaFileGetter fileGetter = new SciaFileGetter();
            var EsaFile = fileGetter.PrepareBasicEmptyFile(@"C:/TEMP/");//path where the template file "template.esa" is created
            if (!File.Exists(EsaFile))
            {
                throw new InvalidOperationException($"File from manifest resource is not created ! Temp: {env.AppTempPath}");
            }

            SCIA.OpenAPI.EsaProject proj = env.OpenProject(EsaFile);
            //SCIA.OpenAPI.EsaProject proj = env.OpenProject(SciaEngineerProjecTemplate);
            if (proj == null)
            {
                throw new InvalidOperationException($"File from manifest resource is not opened ! Temp: {env.AppTempPath}");
            }
            Console.WriteLine($"Proj opened");

                  

            // info about Project 
            ProjectInformation projectInformation = new ProjectInformation(Guid.NewGuid(), "ProjectX")
            {
                BuildingType = "SimpleFrame",
                Location = "39XG+P7 Praha",
                LastUpdate = DateTime.Today,
                Status = "Draft",
                ProjectType = "New construction"
            };

            

            // info about Model ModelExchanger.AnalysisDataModel.ModelInformation
            ModelInformation modelInformation = new ModelInformation(Guid.NewGuid(), "ModelOne")
            {
                Discipline = "Static",
                Owner = "JB",
                LevelOfDetail = "200",
                LastUpdate = DateTime.Today,
                SourceApplication = "OpenAPI",
                RevisionNumber = "1",
                SourceCompany = "SCIA",
                SystemOfUnits = SystemOfUnits.Metric

            };

            Console.WriteLine($"Set grade for concrete material: ");
            string conMatGrade = Console.ReadLine();

            StructuralMaterial concrete = new StructuralMaterial(Guid.NewGuid(), "Concrete", MaterialType.Concrete, conMatGrade);


            Console.WriteLine($"Set grade for steel material: ");
            string steelMatGrade = Console.ReadLine();
            StructuralMaterial steel = new StructuralMaterial(Guid.NewGuid(), "Steel", MaterialType.Steel, steelMatGrade);
            ResultOfPartialAddToAnalysisModel addResult = proj.Model.CreateAdmObject(concrete, steel);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
           
            Console.WriteLine($"Materials created in ADM");

            //Create cross-sections in local ADM
            Console.WriteLine($"Set steel profile: ");
            string steelProfile = Console.ReadLine();

            StructuralCrossSection steelprofile = new StructuralManufacturedCrossSection(Guid.NewGuid(), steelProfile, steel, steelProfile, FormCode.ISection, DescriptionId.EuropeanIBeam);

            addResult = proj.Model.CreateAdmObject(steelprofile);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Console.WriteLine($"Set height of concrete rectangle in mm: ");
            double heigth = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine($"Set width of concrete rectangle in mm: ");
            double width = Convert.ToDouble(Console.ReadLine());
            StructuralCrossSection concreteRectangle = new StructuralParametricCrossSection(Guid.NewGuid(), "Concrete", concrete, ProfileLibraryId.Rectangle, new UnitsNet.Length[2] { UnitsNet.Length.FromMillimeters(heigth), UnitsNet.Length.FromMillimeters(width) });
            addResult = proj.Model.CreateAdmObject(concreteRectangle);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            Console.WriteLine($"CSSs created in ADM");

            Console.WriteLine($"Set parameter a: ");
            double a = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine($"Set parameter b: ");
            double b = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine($"Set parameter c: ");
            double c = Convert.ToDouble(Console.ReadLine());


            StructuralPointConnection N1 = new StructuralPointConnection(Guid.NewGuid(), "N1", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N2 = new StructuralPointConnection(Guid.NewGuid(), "N2", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N3 = new StructuralPointConnection(Guid.NewGuid(), "N3", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N3);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N4 = new StructuralPointConnection(Guid.NewGuid(), "N4", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N5 = new StructuralPointConnection(Guid.NewGuid(), "N5", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N5);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N6 = new StructuralPointConnection(Guid.NewGuid(), "N6", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N6);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N7 = new StructuralPointConnection(Guid.NewGuid(), "N7", UnitsNet.Length.FromMeters(a), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N7);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N8 = new StructuralPointConnection(Guid.NewGuid(), "N8", UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(b), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N8);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB1lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N5 }) };
            StructuralCurveMember B1 = new StructuralCurveMember(Guid.NewGuid(), "B1", beamB1lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                Layer = "Columns",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB2lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N2, N6 }) };
            StructuralCurveMember B2 = new StructuralCurveMember(Guid.NewGuid(), "B2", beamB2lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                Layer = "Columns",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB3lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N7 }) };
            StructuralCurveMember B3 = new StructuralCurveMember(Guid.NewGuid(), "B3", beamB3lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                Layer = "Columns",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B3);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB4lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N8 }) };
            StructuralCurveMember B4 = new StructuralCurveMember(Guid.NewGuid(), "B4", beamB4lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Column),
                Layer = "Columns",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB5lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N5, N6 }) };
            StructuralCurveMember B5 = new StructuralCurveMember(Guid.NewGuid(), "B5", beamB5lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                Layer = "Beams",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B5);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB6lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N6, N7 }) };
            StructuralCurveMember B6 = new StructuralCurveMember(Guid.NewGuid(), "B6", beamB6lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                Layer = "Beams",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B6);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB7lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N8 }) };
            StructuralCurveMember B7 = new StructuralCurveMember(Guid.NewGuid(), "B7", beamB7lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                Layer = "Beams",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B7);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var beamB8lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N5 }) };
            StructuralCurveMember B8 = new StructuralCurveMember(Guid.NewGuid(), "B8", beamB8lines, steelprofile)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                Layer = "Beams",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B8);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            Constraint<UnitsNet.RotationalStiffness?> FreeRotation = new Constraint<UnitsNet.RotationalStiffness?>(ConstraintType.Free, UnitsNet.RotationalStiffness.FromKilonewtonMetersPerRadian(0));
            Constraint<UnitsNet.RotationalStiffness?> FixedRotation = new Constraint<UnitsNet.RotationalStiffness?>(ConstraintType.Rigid, UnitsNet.RotationalStiffness.FromKilonewtonMetersPerRadian(1e+10));
            Constraint<UnitsNet.ForcePerLength?> FixedTranslation = new Constraint<UnitsNet.ForcePerLength?>(ConstraintType.Rigid, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(1e+10));
            Constraint<UnitsNet.ForcePerLength?> FreeTranslation = new Constraint<UnitsNet.ForcePerLength?>(ConstraintType.Free, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(0));
            Constraint<UnitsNet.Pressure?> FixedTranslationLine = new Constraint<UnitsNet.Pressure?>(ConstraintType.Rigid, UnitsNet.Pressure.FromKilopascals(1e+10));
            Constraint<UnitsNet.RotationalStiffnessPerLength?> FreeRotationLine = new Constraint<UnitsNet.RotationalStiffnessPerLength?>(ConstraintType.Free, UnitsNet.RotationalStiffnessPerLength.FromKilonewtonMetersPerRadianPerMeter(0));

            StructuralPointSupport PS1 = new StructuralPointSupport(Guid.NewGuid(), "SPS1", N1)
            {
                RotationX = FreeRotation,
                RotationY = FixedRotation,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation
            };

            StructuralPointSupport PS2 = new StructuralPointSupport(Guid.NewGuid(), "SPS2", N2)
            {
                RotationX = FreeRotation,
                RotationY = FixedRotation,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation
            };
            StructuralPointSupport PS3 = new StructuralPointSupport(Guid.NewGuid(), "SPS3", N3)
            {
                RotationX = FreeRotation,
                RotationY = FixedRotation,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation
            };
            StructuralPointSupport PS4 = new StructuralPointSupport(Guid.NewGuid(), "SPS4", N4)
            {
                RotationX = FreeRotation,
                RotationY = FixedRotation,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation
            };
            addResult = proj.Model.CreateAdmObject(PS1, PS2, PS3, PS4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Console.WriteLine($"Set thickness of the slab: ");
            double thickness = Convert.ToDouble(Console.ReadLine());
            var edgecurves = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N5, N6 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N6, N7 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N8 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N5 })
                };
            StructuralSurfaceMember S1 = new StructuralSurfaceMember(Guid.NewGuid(), "S1", edgecurves, concrete, UnitsNet.Length.FromMeters(thickness))
            {
                Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Plate),
                Behaviour = Member2DBehaviour.Isotropic,
                Alignment = Member2DAlignment.Centre,
                EccentricityEz = UnitsNet.Length.FromMeters(0),
                Shape = Member2DShape.Flat
            };
            addResult = proj.Model.CreateAdmObject(S1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Console.WriteLine($"Set length of opening in slab  in m: ");
            double lengthOpening = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine($"Set width of opening in slab  in m: ");
            double withOpening = Convert.ToDouble(Console.ReadLine());

            StructuralPointConnection N9 = new StructuralPointConnection(Guid.NewGuid(), "N9", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N9);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N10 = new StructuralPointConnection(Guid.NewGuid(), "N10", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N10);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N11 = new StructuralPointConnection(Guid.NewGuid(), "N11", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N11);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N12 = new StructuralPointConnection(Guid.NewGuid(), "N12", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(c));
            addResult = proj.Model.CreateAdmObject(N12);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            var openingEdges = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N9, N10 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N10, N11 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N11, N12 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N12, N9 })
                };
            StructuralSurfaceMemberOpening O1S1 = new StructuralSurfaceMemberOpening(Guid.NewGuid(), "O1", S1, openingEdges);
            addResult = proj.Model.CreateAdmObject(O1S1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }


            var edgecurvesS2 = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N2 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N2, N3 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N4 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N1 })
                };
            StructuralSurfaceMember S2 = new StructuralSurfaceMember(Guid.NewGuid(), "S2", edgecurvesS2, concrete, UnitsNet.Length.FromMeters(thickness))
            {
                Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Plate),
                Behaviour = Member2DBehaviour.Isotropic,
                Alignment = Member2DAlignment.Centre,
                EccentricityEz = UnitsNet.Length.FromMeters(0),
                Shape = Member2DShape.Flat
            };
            addResult = proj.Model.CreateAdmObject(S2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralPointConnection N13 = new StructuralPointConnection(Guid.NewGuid(), "N13", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N13);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N14 = new StructuralPointConnection(Guid.NewGuid(), "N14", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b - 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N14);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N15 = new StructuralPointConnection(Guid.NewGuid(), "N15", UnitsNet.Length.FromMeters(0.5 * a + 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N15);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointConnection N16 = new StructuralPointConnection(Guid.NewGuid(), "N16", UnitsNet.Length.FromMeters(0.5 * a - 0.5 * lengthOpening), UnitsNet.Length.FromMeters(0.5 * b + 0.5 * withOpening), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N16);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            var regionEdges = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N13, N14 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N14, N15 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N15, N16 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N16, N13 })
                };
            StructuralSurfaceMemberRegion SMR = new StructuralSurfaceMemberRegion(Guid.NewGuid(), "Region", S2, regionEdges, concrete)
            {
                Thickness = UnitsNet.Length.FromMeters(2 * thickness),
                EccentricityEz = UnitsNet.Length.FromMeters(0),
                Alignment = Member2DAlignment.Centre
            };
            addResult = proj.Model.CreateAdmObject(SMR);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Subsoil subsoil = new Subsoil("Subsoil", UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(80.5), UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(35.5), UnitsNet.SpecificWeight.FromMeganewtonsPerCubicMeter(50), UnitsNet.ForcePerLength.FromMeganewtonsPerMeter(15.5), UnitsNet.ForcePerLength.FromMeganewtonsPerMeter(10.2));
            StructuralSurfaceConnection SS1 = new StructuralSurfaceConnection(Guid.NewGuid(), "SS1", S2, subsoil);
            addResult = proj.Model.CreateAdmObject(SS1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            var edgecurvesS3 = new Curve<StructuralPointConnection>[4] {
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N3, N4 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N4, N8 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N8, N7 }),
                    new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N7, N3 })
                };
            StructuralSurfaceMember S3 = new StructuralSurfaceMember(Guid.NewGuid(), "S3", edgecurvesS3, concrete, UnitsNet.Length.FromMeters(thickness))
            {
                Type = new CSInfrastructure.FlexibleEnum<Member2DType>(Member2DType.Wall),
                Behaviour = Member2DBehaviour.Isotropic,
                Alignment = Member2DAlignment.Centre,
                EccentricityEz = UnitsNet.Length.FromMeters(0),
                Shape = Member2DShape.Flat
            };
            addResult = proj.Model.CreateAdmObject(S3);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            RelConnectsSurfaceEdge LH1 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH1", S3, 0)
            {
                StartPointRelative = 0,
                EndPointRelative = 1,
                TranslationX = FixedTranslationLine,
                TranslationY = FixedTranslationLine,
                TranslationZ = FixedTranslationLine,
                RotationX = FreeRotationLine,
            };
            addResult = proj.Model.CreateAdmObject(LH1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsSurfaceEdge LH2 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH2", S3, 1)
            {
                StartPointRelative = 0,
                EndPointRelative = 1,
                TranslationX = FixedTranslationLine,
                TranslationY = FixedTranslationLine,
                TranslationZ = FixedTranslationLine,
                RotationX = FreeRotationLine,
            };
            addResult = proj.Model.CreateAdmObject(LH2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsSurfaceEdge LH3 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH3", S3, 2)
            {
                StartPointRelative = 0,
                EndPointRelative = 1,
                TranslationX = FixedTranslationLine,
                TranslationY = FixedTranslationLine,
                TranslationZ = FixedTranslationLine,
                RotationX = FreeRotationLine,
            };
            addResult = proj.Model.CreateAdmObject(LH3);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsSurfaceEdge LH4 = new RelConnectsSurfaceEdge(Guid.NewGuid(), "LH4", S3, 3)
            {
                StartPointRelative = 0,
                EndPointRelative = 1,
                TranslationX = FixedTranslationLine,
                TranslationY = FixedTranslationLine,
                TranslationZ = FixedTranslationLine,
                RotationX = FreeRotationLine,
            };
            addResult = proj.Model.CreateAdmObject(LH4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }


            RelConnectsStructuralMember H1 = new RelConnectsStructuralMember(Guid.NewGuid(), "H1", B1)
            {
                Position = Position.End,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation,
                RotationX = FreeRotation,
            };
            addResult = proj.Model.CreateAdmObject(H1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsStructuralMember H2 = new RelConnectsStructuralMember(Guid.NewGuid(), "H2", B2)
            {
                Position = Position.End,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation,
                RotationX = FreeRotation,
                RotationY = FreeRotation,
                RotationZ = FreeRotation
            };
            addResult = proj.Model.CreateAdmObject(H2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsStructuralMember H3 = new RelConnectsStructuralMember(Guid.NewGuid(), "H3", B3)
            {
                Position = Position.End,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation,
                RotationX = FreeRotation,
                RotationY = FreeRotation,
                RotationZ = FreeRotation
            };
            addResult = proj.Model.CreateAdmObject(H3);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            RelConnectsStructuralMember H4 = new RelConnectsStructuralMember(Guid.NewGuid(), "H4", B4)
            {
                Position = Position.End,
                TranslationX = FixedTranslation,
                TranslationY = FixedTranslation,
                TranslationZ = FixedTranslation,
                RotationX = FreeRotation,
                RotationY = FreeRotation,
                RotationZ = FreeRotation
            };
            addResult = proj.Model.CreateAdmObject(H4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralPointConnection N17 = new StructuralPointConnection(Guid.NewGuid(), "N17", UnitsNet.Length.FromMeters(-1 * b), UnitsNet.Length.FromMeters(0), UnitsNet.Length.FromMeters(0));
            addResult = proj.Model.CreateAdmObject(N17);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            var beamB9lines = new Curve<StructuralPointConnection>[1] { new Curve<StructuralPointConnection>(CurveGeometricalShape.Line, new StructuralPointConnection[2] { N1, N17 }) };
            StructuralCurveMember B9 = new StructuralCurveMember(Guid.NewGuid(), "B9", beamB9lines, concreteRectangle)
            {
                Behaviour = CurveBehaviour.Standard,
                SystemLine = CurveAlignment.Centre,
                Type = new CSInfrastructure.FlexibleEnum<Member1DType>(Member1DType.Beam),
                Layer = "Beams",
                EccentricityEy = UnitsNet.Length.FromMeters(0),
                EccentricityEz = UnitsNet.Length.FromMeters(0)
            };
            addResult = proj.Model.CreateAdmObject(B9);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralCurveConnection LSB = new StructuralCurveConnection(Guid.NewGuid(), "LSB", B9)
            {
                Origin = Origin.FromStart,
                CoordinateDefinition = CoordinateDefinition.Relative,
                StartPointRelative = 0.25,
                EndPointRelative = 0.75
            };
            addResult = proj.Model.CreateAdmObject(LSB);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralLoadGroup LG1 = new StructuralLoadGroup(Guid.NewGuid(), "LG1", LoadGroupType.Variable)
            {
                Load = new CSInfrastructure.FlexibleEnum<Load>(Load.Domestic)
            };
            addResult = proj.Model.CreateAdmObject(LG1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralLoadCase LC1 = new StructuralLoadCase(Guid.NewGuid(), "LC1", ActionType.Variable, LG1, LoadCaseType.Static)
            {
                Duration = Duration.Long,
                Specification = Specification.Standard
            };
            addResult = proj.Model.CreateAdmObject(LC1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Console.WriteLine($"Set value of line load on  kN/m: ");
            double lineloadValue = Convert.ToDouble(Console.ReadLine());


            StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB1 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB1", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B1))
            {
                Direction = ActionDirection.Y,
                Distribution = CurveDistribution.Uniform
            };
            StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB2 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB2", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B2))
            {
                Direction = ActionDirection.Y,
                Distribution = CurveDistribution.Uniform
            };
            StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB3 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB3", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B3))
            {
                Direction = ActionDirection.Y,
                Distribution = CurveDistribution.Uniform
            };
            StructuralCurveAction<CurveStructuralReferenceOnBeam> lineloadB4 = new StructuralCurveAction<CurveStructuralReferenceOnBeam>(Guid.NewGuid(), "lineLoadB4", CurveForceAction.OnBeam, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnBeam(B4))
            {
                Direction = ActionDirection.Y,
                Distribution = CurveDistribution.Uniform
            };

            addResult = proj.Model.CreateAdmObject(lineloadB1, lineloadB2, lineloadB3, lineloadB4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E1 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E1", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 0))
            {
                Direction = ActionDirection.Z
            };
            StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E2 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E2", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 1))
            {
                Direction = ActionDirection.Z
            };
            StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E3 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E3", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 2))
            {
                Direction = ActionDirection.Z
            };
            StructuralCurveAction<CurveStructuralReferenceOnEdge> edgeloadS1E4 = new StructuralCurveAction<CurveStructuralReferenceOnEdge>(Guid.NewGuid(), "edgeLoadS1E4", CurveForceAction.OnEdge, UnitsNet.ForcePerLength.FromKilonewtonsPerMeter(lineloadValue), LC1, new CurveStructuralReferenceOnEdge(S1, 3))
            {
                Direction = ActionDirection.Z
            };

            addResult = proj.Model.CreateAdmObject(edgeloadS1E1, edgeloadS1E2, edgeloadS1E3, edgeloadS1E4);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }


            StructuralLoadCase LC2 = new StructuralLoadCase(Guid.NewGuid(), "LC2", ActionType.Variable, LG1, LoadCaseType.Static)
            {
                Duration = Duration.Long,
                Specification = Specification.Standard
            };
            addResult = proj.Model.CreateAdmObject(LC2);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }

            Console.WriteLine($"Set value of surface load on the slab in kN/m^2: ");
            double surfaceloadValue = Convert.ToDouble(Console.ReadLine());


            StructuralSurfaceAction sls1 = new StructuralSurfaceAction(Guid.NewGuid(), "sls1", UnitsNet.Pressure.FromKilonewtonsPerSquareMeter(surfaceloadValue), S1, LC2)
            {
                Direction = ActionDirection.Z,
                Location = Location.Length
            };

            addResult = proj.Model.CreateAdmObject(sls1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            StructuralPointAction<PointStructuralReferenceOnPoint> FP = new StructuralPointAction<PointStructuralReferenceOnPoint>(Guid.NewGuid(), "FP", UnitsNet.Force.FromKilonewtons(150), LC2, PointForceAction.InNode, new PointStructuralReferenceOnPoint(N13))
            {
                Direction = ActionDirection.Z
            };
            addResult = proj.Model.CreateAdmObject(FP);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            var Combinations = new StructuralLoadCombinationData[2] { new StructuralLoadCombinationData(LC1, 1.0, 1.5), new StructuralLoadCombinationData(LC2, 1.0, 1.35) };
            StructuralLoadCombination C1 = new StructuralLoadCombination(Guid.NewGuid(), "C1", LoadCaseCombinationCategory.AccordingNationalStandard, Combinations)
            {
                NationalStandard = LoadCaseCombinationStandard.EnUlsSetC
            };
            addResult = proj.Model.CreateAdmObject(C1);
            if (addResult.PartialAddResult.Status != AdmChangeStatus.Ok) { throw HandleErrorResult(addResult); }
            proj.Model.RefreshModel_ToSCIAEngineer();

            Console.WriteLine($"My model sent to SEn");



            // Run calculation
            proj.RunCalculation();
            Console.WriteLine($"My model calculate");

            //storage for results
            OpenApiE2EResults storage = new OpenApiE2EResults();

            //Initialize Results API
            ResultsAPI resultsApi = proj.Model.InitializeResultsAPI();
            if (resultsApi == null)
            {
                return storage;
            }
            {
                OpenApiE2EResult beamB1InnerForLc = new OpenApiE2EResult("beamB1InnerForcesLC1")
                {
                    ResultKey = new ResultKey
                    {
                        EntityType = eDsElementType.eDsElementType_Beam,
                        EntityName = B1.Name,
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        Dimension = eDimension.eDim_1D,
                        ResultType = eResultType.eFemBeamInnerForces,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    }
                };
                beamB1InnerForLc.Result = resultsApi.LoadResult(beamB1InnerForLc.ResultKey);
                storage.SetResult(beamB1InnerForLc);
            }
            {
               OpenApiE2EResult beamInnerForcesCombi = new OpenApiE2EResult("beamInnerForcesCombi")
                {
                   ResultKey = new ResultKey
                   {
                        EntityType = eDsElementType.eDsElementType_Beam,
                        EntityName = B1.Name,
                        CaseType = eDsElementType.eDsElementType_Combination,
                        CaseId = C1.Id,
                        Dimension = eDimension.eDim_1D,
                        ResultType = eResultType.eFemBeamInnerForces,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                   }
                };
                beamInnerForcesCombi.Result = resultsApi.LoadResult(beamInnerForcesCombi.ResultKey);
                storage.SetResult(beamInnerForcesCombi);
            }
            {
                OpenApiE2EResult slabInnerForces = new OpenApiE2EResult("slabInnerForces")
                {
                    ResultKey = new ResultKey
                    {
                        EntityType = eDsElementType.eDsElementType_Slab,
                        EntityName = S1.Name,
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        Dimension = eDimension.eDim_2D,
                        ResultType = eResultType.eFemInnerForces,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    }
                };
                slabInnerForces.Result = resultsApi.LoadResult(slabInnerForces.ResultKey);
                storage.SetResult(slabInnerForces);
            }
            {
                OpenApiE2EResult slabDeformations = new OpenApiE2EResult("slabDeformations")
                {
                    ResultKey = new ResultKey
                    {
                        EntityType = eDsElementType.eDsElementType_Slab,
                        EntityName = S1.Name,
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        Dimension = eDimension.eDim_2D,
                        ResultType = eResultType.eFemDeformations,
                        CoordSystem = eCoordSystem.eCoordSys_Local
                    }
                };
                slabDeformations.Result = resultsApi.LoadResult(slabDeformations.ResultKey);
                storage.SetResult(slabDeformations);
            }
            {
                OpenApiE2EResult reactions = new OpenApiE2EResult("Reactions")
                {
                    ResultKey = new ResultKey
                    {
                        CaseType = eDsElementType.eDsElementType_LoadCase,
                        CaseId = LC1.Id,
                        EntityType = eDsElementType.eDsElementType_Node,
                        EntityName = N1.Name,
                        Dimension = eDimension.eDim_reactionsPoint,
                        ResultType = eResultType.eReactionsNodes,
                        CoordSystem = eCoordSystem.eCoordSys_Global
                    }
                };
                reactions.Result = resultsApi.LoadResult(reactions.ResultKey);
                storage.SetResult(reactions);
            }
           // proj.CloseProject(SCIA.OpenAPI.SaveMode.SaveChangesNo);
            return storage;
        }

        private static Exception HandleErrorResult(ResultOfPartialAddToAnalysisModel addResult)
        {
            switch (addResult.PartialAddResult.Status)
            {
                case AdmChangeStatus.InvalidInput:
                    throw new Exception(addResult.PartialAddResult.Warnings);
                case AdmChangeStatus.Error:
                    throw new Exception(addResult.PartialAddResult.Errors);
                case AdmChangeStatus.NotDone:
                    throw new Exception(addResult.PartialAddResult.Warnings);
            }
            if (addResult.PartialAddResult.Exception != null)
            {
                throw addResult.PartialAddResult.Exception;
            }
            throw new Exception("Unknown Error");
        }

        static void Main(string[] args)
        {
            ExcelTest();

            //SciaOpenApiAssemblyResolve();
            //DeleteTemp();

            ////RunSCIAOpenAPI();
            //SciaOpenApiContext Context = new SciaOpenApiContext(SciaEngineerFullPath, SciaOpenApiWorker);//to use this construct you need to have a program exe in SCIA ENG. exe folder
            //SciaOpenApiUtils.RunSciaOpenApi(Context);
            //if (Context.Exception != null)
            //{
            //    Console.WriteLine(Context.Exception.Message);
            //    return;
            //}
            //if (!(Context.Data is OpenApiE2EResults data))
            //{
            //    Console.WriteLine("SOMETHING IS WRONG NO Results DATA !");
            //    return;
            //}
            //foreach (var item in data.GetAll())
            //{
            //    Console.WriteLine(item.Value.Result.GetTextOutput());
            //}
            //var slabDef = data.GetResult("slabDeformations").Result;
            //if (slabDef != null)
            //{
            //    double maxvalue = 0;
            //    double pivot;
            //    for (int i = 0; i < slabDef.GetMeshElementCount(); i++)
            //    {
            //        pivot = slabDef.GetValue(2, i);
            //        if (System.Math.Abs(pivot) > System.Math.Abs(maxvalue))
            //        {
            //            maxvalue = pivot;

            //        }
            //    }
            //    Console.WriteLine("Maximum deformation on slab:");
            //    Console.WriteLine(maxvalue);
            //}
            //Console.WriteLine($"Press key to exit");
            //Console.ReadKey();
        }

        private static void ExcelTest()
        {
            

           
        // info about Project 
        ProjectInformation projectInformation = new ProjectInformation(Guid.NewGuid(), "ProjectX")
            {
                BuildingType = "SimpleFrame",
                Location = "39XG+P7 Praha",
                LastUpdate = DateTime.Today,
                Status = "Draft",
                ProjectType = "New construction"
            };



            // info about Model ModelExchanger.AnalysisDataModel.ModelInformation
            ModelInformation modelInformation = new ModelInformation(Guid.NewGuid(), "ModelOne")
            {
                Discipline = "Static",
                Owner = "JB",
                LevelOfDetail = "200",
                LastUpdate = DateTime.Today,
                SourceApplication = "OpenAPI",
                RevisionNumber = "1",
                SourceCompany = "SCIA",
                SystemOfUnits = SystemOfUnits.Metric

            };
            var AnalysisObjects = new List<IAnalysisObject>();
            AnalysisObjects.Add(projectInformation);
            AnalysisObjects.Add(modelInformation);


            BootstrapperBase bootstrapperADM;
            bootstrapperADM = new BootstrapperBase();
            bootstrapperADM.Boostrapp<ModelExchangerExtensionIntegrationModule>();
            var ModelHolder = bootstrapperADM.Container.Resolve<IAnalysisModelHolder>();
            PartialAddResult actualResult = ModelHolder.AddToModel(AnalysisObjects);
            
           

            BootstrapperBase bootstrapperExchanger;
            bootstrapperExchanger = new BootstrapperBase();
            bootstrapperExchanger.Boostrapp<ModelExchangerExtensionIntegrationModule>();
            ExchangeResult result = bootstrapperExchanger.Container.Resolve<ICoreToExcelFileService>().WriteExcel(ModelHolder.AnalysisModel, @"C:/TEMP/A.xls");

            ExchangeCoreResult readedExcelModel =  bootstrapperExchanger.Container.Resolve<IExcelToCoreFileService>().ReadExcel(@"C:/TEMP/A.xls");
            //readedExcelModel.Model.IsModelValid

            result =  bootstrapperExchanger.Container.Resolve<ICoreToJsonFileService>().WriteJson(ModelHolder.AnalysisModel, @"C:/TEMP/A.json");
            ExchangeCoreResult readedJsonModel =  bootstrapperExchanger.Container.Resolve<IJsonToCoreFileService>().ReadJson(@"C:/TEMP/A.json");

        
        }
    }
}
