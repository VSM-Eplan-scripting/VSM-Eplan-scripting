using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Xml;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
using System.Drawing;
using Eplan.EplApi.ApplicationFramework;

public class Voortman_functions
{
	public static void Main()
	{
		;
	}

	[DeclareAction("StartProcess")]
	public void StartProcess(string ProcessName, string Parameter)
	{
		try
		{
			Parameter = PathMap.SubstitutePath(Parameter);
			Process.Start(ProcessName, Parameter);
		}
		catch (Exception ex)
		{
			MessageBox.Show(
			ex.Message,
			"Failed",
			MessageBoxButtons.OK,
			MessageBoxIcon.Error
			);
		}
		return;
	}

	[DeclareAction("OpenFile")]
	public void OpenFile(string FileName)
	{
		try
		{
			File.Open(FileName, FileMode.Open);
		}
		catch (Exception ex)
		{
			MessageBox.Show(
			ex.Message,
			"Failed",
			MessageBoxButtons.OK,
			MessageBoxIcon.Error);
		}
		return;
	}

	public void VSM_UpdateLayerTextHeight(string LayerName, string TextHeight)
	{

		if (LayerName == "" || TextHeight == "")
		{
			return;
		}
		else
		{
			CommandLineInterpreter cli = new CommandLineInterpreter();
			ActionCallingContext acc = new ActionCallingContext();

			acc.AddParameter("LAYER", LayerName);
			acc.AddParameter("TEXTHEIGHT", TextHeight);

			cli.Execute("changeLayer", acc);
		}

	}

	public void VSM_UpdateLayerColor(string LayerName, string Color)
	{
		//-5 = inverse to background

		if (LayerName == "" || Color == "")
		{
			return;
		}
		else
		{
			CommandLineInterpreter cli = new CommandLineInterpreter();
			ActionCallingContext acc = new ActionCallingContext();

			acc.AddParameter("LAYER", LayerName);
			acc.AddParameter("COLORID", Color);

			cli.Execute("changeLayer", acc);
		}

	}

	public void VSM_ImportProjectSettings(string ProjectName, string SettingsLocation)
	{

		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		string strProject = PathMap.SubstitutePath("$(P)");
		ActionCallingContext ReadXMLFile = new ActionCallingContext();
		ReadXMLFile.AddParameter("XMLFile", SettingsLocation);
		ReadXMLFile.AddParameter("Project", ProjectName);

		if (!(SettingsLocation == string.Empty))
		{
			if (File.Exists(SettingsLocation))
			{
				new CommandLineInterpreter().Execute("XSettingsImport", ReadXMLFile);
				MessageBox.Show("Setting '" + SettingsLocation + "' is imported.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			else
			{
				MessageBox.Show("Setting file '" + SettingsLocation + "' does not exist.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		else
		{
			return;
		}
	}

	public void VSM_ImportSettings(string SettingsLocation, bool Feedback)
	{

		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		ActionCallingContext ReadXMLFile = new ActionCallingContext();
		ReadXMLFile.AddParameter("XMLFile", SettingsLocation);
		ReadXMLFile.AddParameter("Option", "OVERWRITE");

		if (!(SettingsLocation == string.Empty))
		{
			if (File.Exists(SettingsLocation))
			{
				new CommandLineInterpreter().Execute("XSettingsImport", ReadXMLFile);
				if (Feedback)
				{
					MessageBox.Show("Setting '" + SettingsLocation + "' is imported.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			else
			{
				MessageBox.Show("Setting file '" + SettingsLocation + "' does not exist.", "Write settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		else
		{
			return;
		}
	}

	public string VSM_GetSelectedProject()
	{
		//Execute a "selectionset" action to obtain the currently selected project's full name (including folder)
		Eplan.EplApi.ApplicationFramework.ActionManager oMngr = new Eplan.EplApi.ApplicationFramework.ActionManager();
		Eplan.EplApi.ApplicationFramework.Action oSelSetAction = oMngr.FindAction("selectionset");
		string sProjectFolder = string.Empty;
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
			MessageBox.Show("No project selected", "Project selection");
			return selectedProject;
		}
		else
		{
			return selectedProject;
		}

	}

	[DeclareAction("VSM_ExportAllModule")]
	public void VSM_ExportAllModule()
	{
		VSM_MarkingMTP();
		VSM_PLCExport();
		VSM_PLCEbusExport();
		VSM_ExportWireTerminalData();
		VSM_PublishProject();
		VSM_ExportPDF();
	}

	[DeclareAction("VSM_ExportAllAssy")]
	public void VSM_ExportAllAssy()
	{
		VSM_MarkingMTP();
		VSM_PLCExport();
		VSM_PLCEbusExport();
		VSM_ExportWireTerminalData();
		VSM_PublishProject();
		VSM_ExportPDF();
		VSM_MachiningFiles();
	}

	[DeclareAction("VSM_ExportBOM")]
	public void VSM_ExportBOM()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "Export BOM");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VSM parts list with bimmer");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "1");
			labelingContext.AddParameter("USESELECTION", "1");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_ExportProjectBOM")]
	public void VSM_ExportProjectBOM()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "Export Project BOM");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VOORTMAN artikellijst veld Projectspecifiek");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "1");
			labelingContext.AddParameter("USESELECTION", "1");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_ExportPDF")]
	public void VSM_ExportPDF()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "Export PDF");
			ActionCallingContext pdfContext = new ActionCallingContext();
			pdfContext.AddParameter("type", "PDFPROJECTSCHEME");
			pdfContext.AddParameter("exportscheme", "VSM PDF export");
			new CommandLineInterpreter().Execute("export", pdfContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_ExportWireTerminalData")]
	public void VSM_ExportWireTerminalData()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "Export Wire Terminal Data");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VSM Wire Terminal data");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "1");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_PLCEbusExport")]
	public void VSM_PLCEbusExport()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "PLC E-bus export");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VSM PLC E-bus export");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "1");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_PLCExport")]
	public void VSM_PLCExport()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "PLC export");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VSM PLC export");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "1");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	[DeclareAction("VSM_MarkingMTP")]
	public void VSM_MarkingMTP()
	{
		Progress progress = new Progress("SimpleProgress");
		progress.BeginPart(100, "");
		progress.SetAllowCancel(true);
		if (!progress.Canceled())
		{
			progress.BeginPart(100, "Generate Marking help file");
			ActionCallingContext labelingContext = new ActionCallingContext();
			labelingContext.AddParameter("CONFIGSCHEME", "VSM Marking MTP help");
			labelingContext.AddParameter("LANGUAGE", "en_US");
			labelingContext.AddParameter("LogMsgActionDone", "true");
			labelingContext.AddParameter("SHOWOUTPUT", "0");
			new CommandLineInterpreter().Execute("label", labelingContext);
			progress.BeginPart(100, "Call Clip Project Marking");
			StartProcess("c:\\temp\\MTP help text.txt", "");
			VSM_CallClipProjectMarking();
			progress.EndPart();
		}
		progress.EndPart(true);
	}

	public void VSM_CallClipProjectMarking()
	{
		ActionCallingContext acc = new ActionCallingContext();
		acc.AddParameter("ExportTo", "CPM");
		new CommandLineInterpreter().Execute("ExportActionClip", acc);
	}

	[DeclareAction("VSM_MachiningFiles")]
	public void VSM_MachiningFiles()
	{
		ActionCallingContext acc = new ActionCallingContext();
		if (MessageBox.Show("Use 'VSM_Layout space' scheme!", "Generate machining files", MessageBoxButtons.OK) == DialogResult.OK)
		{
			new CommandLineInterpreter().Execute("XNCDlgExportKiesling", acc);
		}

	}

	[DeclareAction("VSM_PublishProject")]
	public void VSM_PublishProject()
	{
		ActionCallingContext acc = new ActionCallingContext();
		new CommandLineInterpreter().Execute("XEdaExportAction /SmartWiringExport:1", acc);
	}

	[DeclareAction("VSM_OpenModules")]
	public void VSM_OpenModules()
	{
		ActionCallingContext acc = new ActionCallingContext();
		CommandLineInterpreter cli = new CommandLineInterpreter();

		using (OpenFileDialog openFileDialog = new OpenFileDialog())
		{
			openFileDialog.InitialDirectory = "\\\\vsm-fs-svr03\\data\\eplan electric p8\\gegevens\\projects\\var\\macro projecten\\modules\\";
			openFileDialog.Filter = "Eplan files (???-????*.el*)|???-????*.el*";
			openFileDialog.FilterIndex = 2;
			openFileDialog.RestoreDirectory = true;
			openFileDialog.Multiselect = false;
			openFileDialog.CheckFileExists = true;
			openFileDialog.CustomPlaces.Add(@"\\vsm-fs-svr03\data\eplan electric p8\gegevens\projects\var\macro projecten\modules\");
			openFileDialog.CustomPlaces.Add(@"\\vsm-fs-svr03\data\eplan electric p8\gegevens\projects\var\macro projecten\modules\Modules niet in EEC");
			openFileDialog.CustomPlaces.Add(@"\\vsm-fs-svr03\data\eplan electric p8\gegevens\projects\var\macro projecten\Assembly");

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				acc.AddParameter("Project", openFileDialog.FileName);
				cli.Execute("ProjectOpen", acc);
			}
		}
	}

	[DeclareAction("VSM_OpenCustomers")]
	public void VSM_OpenCustomers()
	{
		ActionCallingContext acc = new ActionCallingContext();
		CommandLineInterpreter cli = new CommandLineInterpreter();

		using (OpenFileDialog openFileDialog = new OpenFileDialog())
		{
			openFileDialog.InitialDirectory = "\\\\vsm-fs-svr03\\Data\\Eplan Electric P8\\Gegevens\\Projects\\VAR\\Klanten projecten";
			openFileDialog.Filter = "Eplan files (*.el*)|*.el*";
			openFileDialog.FilterIndex = 2;
			openFileDialog.RestoreDirectory = true;
			openFileDialog.Multiselect = false;
			openFileDialog.CheckFileExists = true;
			openFileDialog.CustomPlaces.Add(@"\\vsm-fs-svr03\data\eplan electric p8\gegevens\projects\var\Klanten projecten");

			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				acc.AddParameter("Project", openFileDialog.FileName);
				cli.Execute("ProjectOpen", acc);
			}
		}
	}

	//[DeclareAction("VSM_UpdateLayer583")]
	//public void VSM_UpdateLayer583()
	//{
	//	string LayerName = "EPLAN583";
	//	string TextHeight = "5";
	//	string Color = "-5";

	//	VSM_UpdateLayerTextHeight(LayerName, TextHeight);
	//	VSM_UpdateLayerColor(LayerName, Color);
	//	MessageBox.Show("Layer " + LayerName + " updated with fontsize " + TextHeight + " and color " + Color, "Update layer", MessageBoxButtons.OK);
	//}

	//[DeclareAction("VSM_ImportModelViewReports")]
	//public void VSM_ImportModelViewReports()
	//{
	//	string SelectedProject = VSM_GetSelectedProject();

	//	if (!(SelectedProject == string.Empty))
	//	{
	//		//Import Model views
	//		VSM_ImportProjectSettings(SelectedProject, @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Schemes\VAR\Model views\VSM all model view templates.xml");

	//		//Import filter for model views
	//		VSM_ImportProjectSettings(SelectedProject, @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Schemes\VAR\Model views\G3vl.VSM_w_o_mounting_rail.xml");
	//	}
	//}

	//[DeclareAction("VSM_ImportTranslationSettings")]
	//public void VSM_ImportTranslationSettings()
	//{
	//	string SelectedProject = VSM_GetSelectedProject();

	//	if (!(SelectedProject == string.Empty))
	//	{
	//		//Import translation settings
	//		VSM_ImportProjectSettings(SelectedProject, @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Xml\VAR\Vertaling\Projects+Basis project Voortman IEC structuur+Translation.xml");
	//	}
	//}

	[DeclareAction("VSM_ImportProjectsFilter")]
	public void VSM_ImportProjectsFilter()
	{
		SchemeSetting schemeSetting = new SchemeSetting();
		schemeSetting.Init("USER.FormGeneratorGui.FilterScheme.PxfForm_PARTSLIST");

		string value = "VA";

		if (InputBox("Filter string", "Enter filter string:", ref value) == DialogResult.OK)
		{
			VSM_SetVOORTMAN_artikellijst_veld_Projectspecifiek(value);
			VSM_SetVOORTMAN_Kabel_en_veldcodering_Projectspecifiek(value);

			MessageBox.Show("Order specific filter set");
		}
	}

	public string VSM_UpdateExtPlacementFilter(string OrigDocPath, string FilterString)
	{

		string TargetFile = @"C:\temp\" + Path.GetFileName(OrigDocPath);
		string DocString = string.Empty;

		// instantiate XmlDocument and load XML from file
		XmlDocument OrigDoc = new XmlDocument();
		XmlDocument ModifDoc = new XmlDocument();

		OrigDoc.Load(OrigDocPath);

		DocString = OrigDoc.OuterXml;
		DocString = DocString.Replace("nl_NL@VA", string.Concat("nl_NL@", FilterString));

		ModifDoc.LoadXml(DocString);

		// save the XmlDocument back to disk
		ModifDoc.Save(TargetFile);

		return TargetFile;

	}

	public static DialogResult InputBox(string title, string promptText, ref string value)
	{
		Form form = new Form();
		Label label = new Label();
		TextBox textBox = new TextBox();
		Button buttonOk = new Button();
		Button buttonCancel = new Button();

		form.Text = title;
		label.Text = promptText;
		textBox.Text = value;

		buttonOk.Text = "OK";
		buttonCancel.Text = "Cancel";
		buttonOk.DialogResult = DialogResult.OK;
		buttonCancel.DialogResult = DialogResult.Cancel;

		label.SetBounds(9, 15, 372, 13);
		textBox.SetBounds(12, 36, 372, 20);
		buttonOk.SetBounds(228, 72, 75, 23);
		buttonCancel.SetBounds(309, 72, 75, 23);

		label.AutoSize = true;
		textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
		buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

		form.ClientSize = new Size(396, 107);
		form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
		form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
		form.FormBorderStyle = FormBorderStyle.FixedDialog;
		form.StartPosition = FormStartPosition.CenterScreen;
		form.MinimizeBox = false;
		form.MaximizeBox = false;
		form.AcceptButton = buttonOk;
		form.CancelButton = buttonCancel;

		DialogResult dialogResult = form.ShowDialog();
		value = textBox.Text;
		return dialogResult;
	}

	public void VSM_SetVOORTMAN_artikellijst_veld_Projectspecifiek(string FilterValue)
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		oSettings.SetStringSetting("USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeData", "0|1|0|EPLAN.PartRef.UserSupplementaryField4;0|0|nl_NL@" + FilterValue + ";|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|0|0|EPLAN.PartRef.UserSupplementaryField5;0|0|NB|0|1|1|0|0|0;0|", 0);
		oSettings.SetStringSetting("USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeName", "Veld installatie klantspecifiek", 0);

		//For debugging
		//string value = oSettings.GetStringSetting("USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeData", 0);
		//MessageBox.Show("Value of Index " + 1 + ":\n" + value, "USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeData");
	}

	public void VSM_SetVOORTMAN_Kabel_en_veldcodering_Projectspecifiek(string FilterValue)
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		oSettings.SetStringSetting("USER.Labelling.Config.VSM cable and field coding project specific.Data.SortFilter.FilterSchemeData", "0|1|0|22220;0|0|1|0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|22041;0|0|29|0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|22041;0|0|12|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|EPLAN.PartRef.UserSupplementaryField4;0|0|nl_NL@" + FilterValue + ";|0|1|1|0|0|0;0|", 0);
		oSettings.SetStringSetting("USER.Labelling.Config.VSM cable and field coding project specific.Data.SortFilter.FilterSchemeName", "External placement and cables project specific", 0);

		//For debugging
		//string value = oSettings.GetStringSetting("USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeData", 0);
		//MessageBox.Show("Value of Index " + 1 + ":\n" + value, "USER.Labelling.Config.VOORTMAN artikellijst veld Projectspecifiek.Data.SortFilter.FilterSchemeData");
	}

	[DeclareAction("VSM_ExportBOMOnStructure")]
	public void VSM_ExportBOMOnStructure()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeData", "0|1|1|EPLAN.PartRef.UserSupplementaryField4;0|0|VA|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0|#0|1|1|EPLAN.PartRef.UserSupplementaryField5;0|0|??_??@NB;|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0|#0|1|1|19007;0|0|*EEC*|0|1|1|0|0|0;0|", 0);
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeName", "Filter _NB, _VA,_*EEC*", 0);
		VSM_ExportBOM();
	}

	[DeclareAction("VSM_ExportBOMOnExt_Field")]
	public void VSM_ExportBOMOnExt_Field()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeData", "0|1|1|EPLAN.PartRef.UserSupplementaryField5;0|0|??_??@NB;|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|22041;0|0|29|0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|22220;0|0|1|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|1|20494;0|0|1|0|1|1|0|0|0;0|#3|1|0|;0|0||0|1|1|0|0|0;0|#0|1|0|22220;0|0|1|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|1|20450;0|3|1|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|1|20024;0|0|nl_NL@Montageplaat;|0|1|1|0|0|0;0|#7|1|0|;0|0||0|1|1|0|0|0;0|#0|1|1|EPLAN.PartRef.UserSupplementaryField4;0|0|nl_NL@VA;|0|1|1|0|0|0;0|", 0);
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeName", "Filter parts field", 0);
		VSM_ExportBOM();
	}

	[DeclareAction("VSM_ExportBOMOnExt_Panel")]
	public void VSM_ExportBOMOnExt_Panel()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeData", "0|1|0|20484;0|0|0|0|1|1|0|0|0;0|#7|1|0|0;0|0||0|1|1|0|0|0;0|#0|1|1|20494;0|0|1|0|1|1|0|0|0;0|#7|1|0|0;0|0||0|1|1|0|0|0;0|#0|1|1|40305;0|0|NB|0|1|1|0|0|0;0|#7|1|0|0;0|0||0|1|1|0|0|0;0|#0|1|1|20121;0|0|-6|0|1|1|0|0|0;0|#7|1|0|0;0|0||0|1|1|0|0|0;0|#0|1|1|22220;0|0|1|0|1|1|0|0|0;0|", 0);
		oSettings.SetStringSetting("USER.Labelling.Config.VSM parts list with bimmer.Data.SortFilter.FilterSchemeName", "Filter parts cabinet", 0);
		VSM_ExportBOM();
	}


}

