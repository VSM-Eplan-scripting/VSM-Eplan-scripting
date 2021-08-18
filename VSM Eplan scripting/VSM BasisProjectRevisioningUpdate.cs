using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Threading.Tasks;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
public class BasisProjectRevisioning
{
    //Called from Voortman menu, only by certain people
    [DeclareAction("VSM_ReleaseBasicProject")]
    public void VSM_ReleaseBasicProject()
    {

        Form form = new Form();
        TextBox textBox = new TextBox();

        string BasicProjectName = @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten\Basis project Voortman IEC structuur.elk";
        string ProgressText = "VSM File Basic project versioning:" + Environment.NewLine;

        //Build form		
        form.Text = "VSM File deploy";
        textBox.ReadOnly = true;
        textBox.Multiline = true;

        textBox.Text = ProgressText;

        textBox.SetBounds(12, 12, 800, 200);

        textBox.Anchor = textBox.Anchor | AnchorStyles.Left;

        form.Controls.AddRange(new Control[] { textBox, textBox });
        form.ClientSize = new Size(Math.Max(300, textBox.Right + 12), Math.Max(200, textBox.Bottom + 12));
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;

        //only if basic project
        if (VSM_GetSelectedProject() == BasicProjectName)
        {
            int ProjectVersionString = Int32.Parse(GetProjectPropertyAction("EPLAN.Project.UserSupplementaryField5", null, BasicProjectName));
            int NewBasicProjectVersion = ProjectVersionString + 1;

            ProgressText = ProgressText + "Set Basic Project Version file to " + NewBasicProjectVersion + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Show();
            SetBasicProjectVersion(NewBasicProjectVersion);
            ProgressText = ProgressText + "Set Basic Project Version property to " + NewBasicProjectVersion + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            SetProjectPropertyVersion(BasicProjectName, NewBasicProjectVersion);
            ProgressText = ProgressText + "Archive XML settings" + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ArchiveBasicProjectXML(NewBasicProjectVersion);
            ProgressText = ProgressText + "Export XML settings"  + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ExportBasicProjectXML(BasicProjectName);
            ProgressText = ProgressText + "Archive old Basic project template "  + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ArchiveBasicProjectELC(NewBasicProjectVersion);
            ProgressText = ProgressText + "Export layer settings" + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ExportBasicProjectELC(BasicProjectName);
            ProgressText = ProgressText + "Archive old layer " + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ArchiveBasicProjectTemplate(NewBasicProjectVersion);
            ProgressText = ProgressText + "Export new Basic project template "  + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ExportBasicProjectTemplate();
            ProgressText = ProgressText + "Archive old Basic project backup " + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ArchiveBasicProjectBackup(NewBasicProjectVersion);
            ProgressText = ProgressText + "Export new Basic project backup " + Environment.NewLine;
            textBox.Text = ProgressText;
            form.Update();
            ExportBasicProjectBackup();

            form.Hide();
        }
        else
        {
            MessageBox.Show("Basic Project not selected");
        }
    }


    //[DeclareAction("VSM_UpdateProjectsInFolder")]
    //public void VSM_UpdateProjectsInFolder()
    //{
    //    FolderBrowserDialog folderDlg = new System.Windows.Forms.FolderBrowserDialog();
    //    folderDlg.ShowNewFolderButton = true;
    //    folderDlg.SelectedPath = @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Projects\VAR";

    //    ActionCallingContext acc = new ActionCallingContext(); //zak met parameters voor uitvoeren van actie
    //    CommandLineInterpreter cli = new CommandLineInterpreter(); //voert de actie uit

    //    DialogResult result = folderDlg.ShowDialog();
    //    if (result == DialogResult.OK)
    //    {
    //        foreach (string filePath in Directory.GetFiles(folderDlg.SelectedPath, "*.el*"))
    //        {
    //            if (IsEditableProject(filePath))
    //            {
    //                acc.AddParameter("Project", filePath);
    //                cli.Execute("ProjectOpen", acc);
    //                CompareProjectProperty();
    //            }
    //            else
    //            { 
                
    //            }
    //        }
    //    }
    //}

    //Check version when project is openened
    [DeclareEventHandler("Eplan.EplApi.OnPostOpenProject")]
    public void OnPostOpenProject(IEventParameter iEventParameter)
    {
        try
        {
            EventParameterString oEventParameterString = new EventParameterString(iEventParameter);
            //MessageBox.Show("Projekt öffnen:\n" + oEventParameterString.String, "OnPostOpenProject");
            CompareProjectProperty(oEventParameterString.String);
        }
        catch (System.InvalidCastException exc)
        {
            MessageBox.Show(exc.Message, "Failure");
        }
    }

    [DeclareEventHandler("onActionEnd.String.RMSetRemoveWriteProtectionAction")]
    public void OnActionEndRMSetRemoveWriteProtectionAction()
    {
        CompareProjectProperty(VSM_GetSelectedProject());
    }


    public void CompareProjectProperty(string SelectedProject)
    {
        int ProjectVersion;
        string ProjectVersionString;
        string ProjectShortName;
        int BasicProjectVersion;
        DialogResult result;

        //Get versions, from property and from basic project
        ProjectVersionString = GetProjectPropertyAction("EPLAN.Project.UserSupplementaryField5", null, SelectedProject);
        BasicProjectVersion = GetBasicProjectVersion();

        if (SelectedProject != "")
        {
            ProjectShortName = Path.GetFileName(SelectedProject);

            if (IsEditableProject(SelectedProject))
            {
                // if empty, project has never been updated by this tool
                if (ProjectVersionString == "")
                {
                    ProjectVersion = 0;
                }
                else
                {
                    //check if value can be converted to integer, if not it contains text (=wrong)
                    bool canConvert = Int32.TryParse(ProjectVersionString, out ProjectVersion);
                    if (canConvert == false)
                    {
                        MessageBox.Show("Error: Project version (Project User Supplementary Field 5) of project: " + ProjectShortName + " is a non-numeric value!", "Version error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                //Project version is higher than basic project version, should never be the case
                if (ProjectVersion > BasicProjectVersion)
                {
                    MessageBox.Show("Error: Project version (" + ProjectVersion + ") of project: " + ProjectShortName + " is higher than Basic Project version (" + BasicProjectVersion + ").", "Version error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //Project version is lower than basic project version, normal situation with outdated projects
                else if (ProjectVersion < BasicProjectVersion)
                {
                    result = MessageBox.Show("Project version (" + ProjectVersion + ") of project: " + ProjectShortName + " is lower than Basic Project version (" + BasicProjectVersion + "). Update project?", "Version difference", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if (result == DialogResult.Yes)
                    {
                        if (VSM_ImportProjectSettings(SelectedProject, Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), "Basis project Voortman IEC structuur.xml")) &&
                            VSM_ImportProjectUserSettings(SelectedProject, Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), "PcPr.Project.UserSupplementaryField5.xml")) &&
                            ImportBasicProjectELC(SelectedProject, Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), "Basis project Voortman IEC structuur.elc")) &&
                            SetProjectPropertyVersion(SelectedProject, BasicProjectVersion))
                        {
                            MessageBox.Show("Project version of project: " + ProjectShortName + " updated to latest version", "Version update", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                //Project version = the same as basic project version, do nothing
                else
                {
                }
            }
        }
    }

    public bool SetProjectPropertyVersion(string SelectedProject, int ProjectVersion)
    {
        //Import property based on template, write through projectmanagement

        string pathTemplate = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"), "Script templates","SetProjectProperty", "SetProjectProperty_Template.xml");
        string pathScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"), "Script templates", "SetProjectProperty", "SetProjectProperty_Scheme.xml");
        try
        {
            //Read template
            string content = File.ReadAllText(pathTemplate);
            content = content.Replace("%version%", ProjectVersion.ToString());
            
            //Write scheme
            File.WriteAllText(pathScheme, content);
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "SetProjectProperty", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        ActionCallingContext actionCallingContext2 = new ActionCallingContext();
        actionCallingContext2.AddParameter("TYPE", "READPROJECTINFO");
        actionCallingContext2.AddParameter("PROJECTNAME", SelectedProject);
        actionCallingContext2.AddParameter("FILENAME", pathScheme);

        try
        {
            new CommandLineInterpreter().Execute("projectmanagement", actionCallingContext2);
            return true;
        }
        catch (System.Exception ex)
        {
            MessageBox.Show(ex.Message, "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }

    }

     private int GetBasicProjectVersion()
    {
        //Get basic project version from file
        string text = System.IO.File.ReadAllText(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten\Basic project version.txt");
        try
        {
            return Int32.Parse(text);
        }
        catch (FormatException)
        {
            MessageBox.Show("Unable to parse version to integer");
            return 0;
        }
    }
    public bool VSM_ImportProjectSettings(string ProjectName, string SettingsLocation)
    {
        //Import Project settings, the settings in Options>Settings>Project

        ActionCallingContext ReadXMLFile = new ActionCallingContext();
        ReadXMLFile.AddParameter("XMLFile", SettingsLocation);
        ReadXMLFile.AddParameter("Project", ProjectName);

        //Call without variable
        if (!(SettingsLocation == string.Empty))
        {
            //Call with not existing path
            if (File.Exists(SettingsLocation))
            {
                try
                {
                    new CommandLineInterpreter().Execute("XSettingsImport", ReadXMLFile);
                    //MessageBox.Show("Setting '" + SettingsLocation + "' is imported.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                catch (BaseException ex)
                {
                    MessageBox.Show(ex.Message, "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Setting file '" + SettingsLocation + "' does not exist.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        else
        {
            return false;
        }
    }

    public bool VSM_ImportProjectUserSettings(string ProjectName, string SettingsLocation)
    {
        //Import Project user settings, the settings in Project data>Configure properties & Configure property arrangements

        ActionCallingContext ReadXMLFile = new ActionCallingContext();
        ReadXMLFile.AddParameter("XMLFile", SettingsLocation);
        ReadXMLFile.AddParameter("Project", ProjectName);
        ReadXMLFile.AddParameter("Overwrite", "1");

        //Call without variable
        if (!(SettingsLocation == string.Empty))
        {
            //Call with not existing path
            if (File.Exists(SettingsLocation))
            {
                try
                {
                    new CommandLineInterpreter().Execute("XEsUserPropertiesImportAction", ReadXMLFile);
                    //MessageBox.Show("Setting '" + SettingsLocation + "' is imported.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return true;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message, "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Setting file '" + SettingsLocation + "' does not exist.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        else
        {
            return false;
        }
    }

    public async Task SetBasicProjectVersion(int SetBasicProjectVersion)
    {
        //Write file content with version
        File.WriteAllText(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten\Basic project version.txt", SetBasicProjectVersion.ToString());
    }

    private string GetProjectPropertyAction(string id, string index, string SelectedProject)
    {
        //Get Property from project
        string value = null;
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("id", id);
        actionCallingContext.AddParameter("index", index);
        actionCallingContext.AddParameter("SelectedProject", SelectedProject);
        new CommandLineInterpreter().Execute("GetProjectProperty", actionCallingContext);
        actionCallingContext.GetParameter("value", ref value);
        return value;
    }

    [DeclareAction("GetProjectProperty")]
    public void Action(string id, string index, string SelectedProject, out string value)
    {
        //Stolen code :-)
        string pathTemplate = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"), "Script templates",
            "GetProjectProperty", "GetProjectProperty_Template.xml");
        string pathScheme = Path.Combine(PathMap.SubstitutePath("$(MD_SCRIPTS)"), "Script templates",
            "GetProjectProperty", "GetProjectProperty_Scheme.xml");
        bool isUserDefined = string.IsNullOrEmpty(index) || index.ToUpper().Equals("NULL");
        try
        {
            // Set scheme
            const string QUOTE = "\"";
            string content = File.ReadAllText(pathTemplate);
            if (isUserDefined)
            {
                string isSelectedPropertyUserDef =
                    @"<Setting name=" + QUOTE + "SelectedPropertyIdUserDef" + QUOTE + " type=" + QUOTE + "string" + QUOTE + ">" +
                    "<Val>" + id + "</Val>" +
                    "</Setting>";
                content = content.Replace("GetProjectProperty_ID_SelectedPropertyId", "0");
                content = content.Replace("IsSelectedPropertyIdUserDef", isSelectedPropertyUserDef);
                content = content.Replace("GetProjectProperty_INDEX", "0");
                content = content.Replace("GetProjectProperty_ID", id);
            }
            else
            {
                content = content.Replace("GetProjectProperty_ID_SelectedPropertyId", id);
                content = content.Replace("IsSelectedPropertyIdUserDef", "");
                content = content.Replace("GetProjectProperty_INDEX", index);
                content = content.Replace("GetProjectProperty_ID", id);
            }
            File.WriteAllText(pathScheme, content);
            new Settings().ReadSettings(pathScheme);
            string pathOutput = Path.Combine(
                PathMap.SubstitutePath("$(MD_SCRIPTS)"), "GetProjectProperty",
                "GetProjectProperty_Output.txt");
            // Export
            ActionCallingContext actionCallingContext = new ActionCallingContext();
            actionCallingContext.AddParameter("PROJECTNAME", SelectedProject);
            actionCallingContext.AddParameter("configscheme", "GetProjectProperty");
            actionCallingContext.AddParameter("destinationfile", pathOutput);
            actionCallingContext.AddParameter("language", "de_DE");
            new CommandLineInterpreter().Execute("label", actionCallingContext);
            // Read
            value = File.ReadAllLines(pathOutput)[0];
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "GetProjectProperty", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            value = "[Error]";
        }
    }

    public void ExportBasicProjectXML(string ProjectName)
    {
        string BasicProjectXMLFileName = "Basis project Voortman IEC structuur.xml";
        string BasicProjectXMLPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), BasicProjectXMLFileName);

        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("prj", ProjectName);
        actionCallingContext.AddParameter("XMLFile", BasicProjectXMLPath);
        new CommandLineInterpreter().Execute("XSettingsExport", actionCallingContext);
    }

    public void ArchiveBasicProjectXML(int BasicProjectVersion)
    {
        //Archive old BasicProjectXML
        string BasicProjectXMLFileName = "Basis project Voortman IEC structuur.xml";
        string BasicProjectXMLPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), BasicProjectXMLFileName);

        int ArchiveBasicProjectVersion = BasicProjectVersion - 1;

        string ArchiveBasicProjectXMLFileName = "Basis project Voortman IEC structuur " + ArchiveBasicProjectVersion.ToString() + ".xml";
        string ArchiveBasicProjectXMLPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), ArchiveBasicProjectXMLFileName);

        File.Copy(BasicProjectXMLPath, ArchiveBasicProjectXMLPath);
    }
    public void ExportBasicProjectELC(string ProjectName)
    {
        string BasicProjectELCFileName = "Basis project Voortman IEC structuur.elc";
        string BasicProjectELCPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Xml\VAR\Projecten\Basis project settings"), BasicProjectELCFileName);

        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("TYPE", "EXPORT");
        actionCallingContext.AddParameter("PROJECTNAME", ProjectName);
        actionCallingContext.AddParameter("EXPORTFILE", BasicProjectELCPath);
        new CommandLineInterpreter().Execute("GraphicalLayerTable", actionCallingContext);
    }

    public bool ImportBasicProjectELC(string ProjectName, string SettingsLocation)
    {
        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("TYPE", "IMPORT");
        actionCallingContext.AddParameter("PROJECTNAME", ProjectName);
        actionCallingContext.AddParameter("IMPORTFILE", SettingsLocation);

        try
        {
            new CommandLineInterpreter().Execute("GraphicalLayerTable", actionCallingContext);
            return true;
        }
        catch (System.Exception ex)
        {
            MessageBox.Show(ex.Message, "Failure", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }



    public void ArchiveBasicProjectELC(int BasicProjectVersion)
    {
        //Archive old BasicProjectXML
        string BasicProjectELCFileName = "Basis project Voortman IEC structuur.elc";
        string BasicProjectELCPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\ELC\VAR\Projecten\Basis project settings"), BasicProjectELCFileName);

        int ArchiveBasicProjectVersion = BasicProjectVersion - 1;

        string ArchiveBasicProjectELCFileName = "Basis project Voortman IEC structuur " + ArchiveBasicProjectVersion.ToString() + ".elc";
        string ArchiveBasicProjectELCPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\ELC\VAR\Projecten\Basis project settings"), ArchiveBasicProjectELCFileName);

        File.Copy(BasicProjectELCPath, ArchiveBasicProjectELCPath);
    }
    public string VSM_GetSelectedProject()
    {
        //Execute a "selectionset" action to obtain the currently selected project's full name (including folder)
        Eplan.EplApi.ApplicationFramework.ActionManager oMngr = new Eplan.EplApi.ApplicationFramework.ActionManager();
        Eplan.EplApi.ApplicationFramework.Action oSelSetAction = oMngr.FindAction("selectionset");
        string selectedProject = string.Empty;
        if (oSelSetAction != null)
        {
            Eplan.EplApi.ApplicationFramework.ActionCallingContext ctx = new Eplan.EplApi.ApplicationFramework.ActionCallingContext();
            ctx.AddParameter("TYPE", "PROJECT");
            bool bRet = oSelSetAction.Execute(ctx);
            if (bRet)
            {
                ctx.GetParameter("PROJECT", ref selectedProject);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Action could not be executed!");
            }
        }
        else
        {
            System.Windows.Forms.MessageBox.Show("Action Selection could not be executed");
        }

        if (selectedProject == string.Empty)
        {
            //MessageBox.Show("No project selected", "Project selection");
            return selectedProject;
        }
        else
        {
            return selectedProject;
        }

    }

    public void ExportBasicProjectTemplate()
    {
        string BasicProjectTemplateFileName = "Basis project Voortman IEC structuur.ept";
        string BasicProjectTemplatePath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Templates\VAR"), BasicProjectTemplateFileName);

        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("PROJECT", VSM_GetSelectedProject());
        actionCallingContext.AddParameter("TARGET", BasicProjectTemplatePath);
        new CommandLineInterpreter().Execute("XPrjActionProjectCreateProjectTemplate", actionCallingContext);
    }
    public void ArchiveBasicProjectTemplate(int BasicProjectVersion)
    {
        //Archive basic Project tempate
        string BasicProjectTemplateFileName = "Basis project Voortman IEC structuur.ept";
        string BasicProjectTemplatePath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Templates\VAR"), BasicProjectTemplateFileName);

        int ArchiveBasicProjectVersion = BasicProjectVersion - 1;
                
        string ArchiveBasicProjectTemplateFileName = "Basis project Voortman IEC structuur " + ArchiveBasicProjectVersion.ToString() + ".ept";
        string ArchiveBasicProjectTemplatePath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Templates\VAR"), ArchiveBasicProjectTemplateFileName);
        
        File.Copy(BasicProjectTemplatePath, ArchiveBasicProjectTemplatePath);
    }

    public void ExportBasicProjectBackup()
    {
        string ExportBasicProjectBackupFileName = "Basis project Voortman IEC structuur.zw1";
        string ExportBasicProjectBackupPath = @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten";

        ActionCallingContext actionCallingContext = new ActionCallingContext();
        actionCallingContext.AddParameter("TYPE", "PROJECT");
        actionCallingContext.AddParameter("PROJECTNAME", VSM_GetSelectedProject());
        actionCallingContext.AddParameter("ARCHIVENAME", ExportBasicProjectBackupFileName);
        actionCallingContext.AddParameter("DESTINATIONPATH", ExportBasicProjectBackupPath);
        actionCallingContext.AddParameter("BACKUPMEDIA", "DISK");
        actionCallingContext.AddParameter("BACKUPMETHOD", "BACKUP");
        new CommandLineInterpreter().Execute("backup", actionCallingContext);
    }

    public void ArchiveBasicProjectBackup(int BasicProjectVersion)
    {

        string BasicProjectBackupFileName = "Basis project Voortman IEC structuur.zw1";
        string BasicProjectBackupPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten"), BasicProjectBackupFileName);

        int ArchiveBasicProjectVersion = BasicProjectVersion - 1;

        string ArchiveBasicProjectBackupFileName = "Basis project Voortman IEC structuur " + ArchiveBasicProjectVersion.ToString() + ".zw1";
        string ArchiveBasicProjectBackupPath = Path.Combine(PathMap.SubstitutePath(@"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Projects\VAR\Basis projecten"), ArchiveBasicProjectBackupFileName);

        File.Copy(BasicProjectBackupPath, ArchiveBasicProjectBackupPath);
    }

    public bool IsEditableProject(string SelectedProject)
    {
        string Extension = SelectedProject.Substring(SelectedProject.Length - 3);
        if (Extension == "elk" || Extension == "ell")
        {
            return true;
        }
        else
        {
            return false;
        }

    }
}