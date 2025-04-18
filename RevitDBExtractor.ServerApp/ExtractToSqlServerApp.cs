using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.External;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using RevCore = RevitDBExtractor.Core;   // your namespace

public class ExtractToSqlServerApp : IExternalDBApplication
{
    public ExternalDBApplicationResult OnStartup(ControlledApplication app)
    {
        // Run after Revit is fully booted
        app.ApplicationInitialized += OnRevitReady;
        return ExternalDBApplicationResult.Succeeded;
    }

    private void OnRevitReady(object sender, ApplicationInitializedEventArgs e)
    {
        var uiapp = new UIApplication((sender as ControlledApplication)!);
        var app = uiapp.Application;

        // Model path supplied via env var, cmd?line arg, or JSON config
        string model = Environment.GetEnvironmentVariable("REVIT_MODEL")
                       ?? throw new InvalidOperationException("REVIT_MODEL env?var not set");

        using (var doc = app.OpenDocumentFile(model))
        {
            var extractor = new RevCore.ModelExtractor();
            extractor.ExportToSql(doc);          // <-- your Core logic
            doc.Close(false);                    // NO save / sync
        }

        // Gracefully quit Revit once idle
        app.DoActionOnIdle(UIApplication_Quit);
    }

    private static void UIApplication_Quit(object? s, IdlingEventArgs e) =>
        (s as UIApplication)?.Application.Quit();

    public ExternalDBApplicationResult OnShutdown(ControlledApplication app)
        => ExternalDBApplicationResult.Succeeded;
}
