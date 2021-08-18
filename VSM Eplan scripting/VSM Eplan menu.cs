using Eplan.EplApi.Scripting;
using System.Runtime.InteropServices;
using System.Text;
using System;
using System.Windows.Forms;
public class VSM_Eplan_menu
{
	[DeclareMenu]
	public void MenuFunction()
	{
		//The script works as follows;
		//Each menu item needs to have its own menuId, this is done with the menuID1 
		//The main menu is generated after the help menu
		//The ID of this menu is stored in the 2nd variable
		//Further expansion of the menu is done with the 1st variable
		//The ID of each 'to be expanded' menu item is stored in a variable, which is used later to expand this menu, f.i. 'menuId4 = menuId1'

		Eplan.EplApi.Gui.Menu menu = new Eplan.EplApi.Gui.Menu();

		uint menuId1; // Menu id
		uint menuId2;
		uint menuId3;
		uint menuId4;
		uint menuId5;
		uint menuId6;
		uint menuId7;
		uint menuId8;
		uint menuId9;

		menuId1 = 0;
		menuId2 = 0;
		menuId3 = 0;
		menuId4 = 0;
		menuId5 = 0;
		menuId6 = 0;
		menuId7 = 0;
		menuId8 = 0;
		menuId9 = 0;

		//The main menu is generated after the help menu
		menuId1 = menu.AddMainMenu("Voortman", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuHelp, "Elec engineering sharepoint",
		"StartProcess /ProcessName:https://voortmansteelgroup.sharepoint.com/sites/ElektroEngineering /Parameter:", "Open Electrical engineering Sharepoint page", 1);
		menuId2 = menuId1;
		menuId1 = menu.AddPopupMenuItem("Wiring", "Wire size sheet", "StartProcess /ProcessName:https://voortmansteelgroup.sharepoint.com/sites/ElektroEngineering/Engineering%20Documenten%20Bibliotheek/Design%20sheet%20protection%20wire%20size%20NFPA%20IEC%20CEC.pdf /Parameter:",
		"Open wire size sheet", menuId1, 0, false, false);
		menuId1 = menu.AddPopupMenuItem("BOM upload", "BOM upload (+structure/Propanel)", "VSM_ExportBOMOnStructure", "Export BOM with bimmer", menuId2, 0, false, false);
		menuId3 = menuId1;
		menuId1 = menu.AddPopupMenuItem("EDR module/assy exports", "Marking MTP file (M/A)", "VSM_MarkingMTP", "Export MTP file", menuId2, 0, false, false);
		menuId4 = menuId1;
		menuId1 = menu.AddPopupMenuItem("Open projects", "Modules/assy", "VSM_OpenModules", "Open module projects", menuId2, 0, false, false);
		menuId5 = menuId1;
		//menuId1 = menu.AddPopupMenuItem("Update Eplan projects", "Update 3D DT layer", "VSM_UpdateLayer583", "Update 3D DT layer", menuId2, 0, false, false);
		//menuId6 = menuId1;
		//menuId1 = menu.AddPopupMenuItem("Update model views", "Import model view reports", "VSM_ImportModelViewReports", "VSM_ImportModelViewReports", menuId6, 0, false, false);
		//menuId7 = menuId1;
		menuId1 = menu.AddPopupMenuItem("For Projects", "Set 'Order project specific' filter", "VSM_ImportProjectsFilter", "Set 'Order project specific' filter", menuId2, 0, false, false);
		menuId8 = menuId1;

		if(Environment.UserName == "m.pluimers" | Environment.UserName == "r.vandenberg")
		{
			menuId1 = menu.AddPopupMenuItem("Eplan management", "Release basic project", "VSM_ReleaseBasicProject", "Release basic project", menuId2, 0, false, false);
			menuId9 = menuId1;
		}

		// Menu Update Eplan projects
		//menuId1 = menu.AddMenuItem("Update translation settings", "VSM_ImportTranslationSettings", "Import translation settings", menuId6, 0, false, false);

		// Menu BOM upload
		menuId1 = menu.AddMenuItem("BOM upload (ext placement/field)", "VSM_ExportBOMOnExt_Field", "Export all deliverables", menuId3, 0, false, false);
		menuId1 = menu.AddMenuItem("BOM upload (int placement/panel)", "VSM_ExportBOMOnExt_Panel", "Export all deliverables", menuId3, 0, false, false);

		//	EDR Module assy menu
		menuId1 = menu.AddMenuItem("Modules export all", "VSM_ExportAllModule", "Export all deliverables", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("Assembly export all", "VSM_ExportAllAssy", "Export all deliverables", menuId4, 0, false, true);
		menuId1 = menu.AddMenuItem("PLC export (M/A)", "VSM_PLCExport", "Export PLC", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("PLC e-bus export (M/A)", "VSM_PLCEbusExport", "Export PLC e-bus", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("Wiring list (M/A)", "VSM_ExportWireTerminalData", "Export wiring list", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("Machining files (A)", "VSM_MachiningFiles", "Export machining files", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("Publish project (M/A)", "VSM_PublishProject", "Publish project", menuId4, 0, false, false);
		menuId1 = menu.AddMenuItem("PDF export (M/A)", "VSM_ExportPDF", "Export PDF", menuId4, 0, false, false);

		//	Open projects menu
		menuId1 = menu.AddMenuItem("Customers", "VSM_OpenCustomers", "Open customer projects", menuId5, 0, false, false);

		//	Update projects menu
		menuId1 = menu.AddMenuItem("Export Project BOM", "VSM_ExportProjectBOM", "Export Project BOM", menuId8, 0, false, false);


	}

	//[DeclareMenu]
	//public void PolkaMenu()
	//{
	//	Eplan.EplApi.Gui.Menu menu = new Eplan.EplApi.Gui.Menu();
	//	uint menuId1;
	//	menuId1 = menu.AddMenuItem("Play Polka musik", "EPLAN_PlayPolka", "Play Polka musik", 35571, 1, false, false);
	//}

	//[DeclareAction("EPLAN_PlayPolka")]
	//public void EPLAN_PlayPolka()
	//{
	//	Player mp3player = new Player();
	//	mp3player.Open(@"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Scripts\VAR\Special\Polka.mp3");
	//	mp3player.Play();
	//}


}

//public class Player
//{
//	//To import the dll winmn.dll which allows to play mp3 files
//	[DllImport("winmm.dll")]
//	private static extern long mciSendString(string lpstrCommand, StringBuilder lpstrReturnString, int uReturnLength, int hwndCallback);



//	public void Open(string file)
//	{
//		string command = "open \"" + file + "\" type MPEGVideo alias Music";
//		mciSendString(command, null, 0, 0);
//	}

//	public void Play()
//	{
//		string command = "play Music";
//		mciSendString(command, null, 0, 0);
//	}
//}

public class MyEventHandler
{
	[DeclareEventHandler("onActionStart.String.XPamAutoCreateFunctionTemplate")]
	public void Action()
	{
		MessageBox.Show("Weet je het zeker dat je dit wilt? Voor gehaktballen bel de Sterkerij 0548 543 859.");
		return;
	}
}