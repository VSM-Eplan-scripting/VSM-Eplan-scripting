//===================================================
// LUC S.  04-07-2018
// Script Exportiert die Fehlworteliste für die eingestellte Projektsprache
//===================================================
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
//==========================================
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.Scripting;
//==========================================
public class Export_Project_Missing_Translation
{
    [DeclareAction("Export_Project_Missing_Translation")]
    //[Start]
    public void Export_Txt_Fehlworte()
    {
        //=======================================================================  
        // Dialogabfrage
        const string message = "Prüfung von fehlenden Übersetzungen durchführen?";
        const string caption = "Export Fehlworteliste";
        var result = MessageBox.Show(message, caption,
                                     MessageBoxButtons.YesNo,
                                     MessageBoxIcon.Question);
        if (result == DialogResult.No)
        {
            return;
        }
        //=======================================================================
        // aktuelles Projektpfad ermitteln    
        string sProject = Get_Project();
        //sProject = sProject.Replace("File-2", "K-ELT-01");
        if (sProject == "")
        {
            MessageBox.Show("Projekt auswählen !");
            return;
        }
        //MessageBox.Show(sProject);

        // Projektname ermitteln
        string strProjectname = Get_Name(sProject);
        //=======================================================================
        //eingestellte Projektsprache EPLAN ermitteln 
        string strDisplayLanguage = null;
        ActionCallingContext ACCDisplay = new ActionCallingContext();
        new CommandLineInterpreter().Execute("GetDisplayLanguage", ACCDisplay);
        ACCDisplay.GetParameter("value", ref strDisplayLanguage);
        //MessageBox.Show("Language : " + strDisplayLanguage);

        //=======================================================================    
        //Fehlworteliste erzeugen :
        Eplan.EplApi.ApplicationFramework.ActionCallingContext acctranslate = new Eplan.EplApi.ApplicationFramework.ActionCallingContext();
        Eplan.EplApi.ApplicationFramework.CommandLineInterpreter CLItranslate = new Eplan.EplApi.ApplicationFramework.CommandLineInterpreter();
        Eplan.EplApi.Base.Progress progress = new Eplan.EplApi.Base.Progress("SimpleProgress");
        progress.BeginPart(100, "");
        progress.SetAllowCancel(true);
        string MisTranslateFile = @"c:\TEMP\EPLAN\EPLAN_Fehlworteliste_" + strProjectname + "_" + strDisplayLanguage + ".txt";
        acctranslate.AddParameter("TYPE", "EXPORTMISSINGTRANSLATIONS");
        acctranslate.AddParameter("LANGUAGE", strDisplayLanguage);
        acctranslate.AddParameter("EXPORTFILE", MisTranslateFile);
        acctranslate.AddParameter("CONVERTER", "XTrLanguageDbXml2TabConverterImpl");
        bool sRet = CLItranslate.Execute("translate", acctranslate);
        if (!sRet)
        {
            MessageBox.Show("Fehler bei Export fehlende Übersetzungen!");
            return;
        }
        // MessageBox.Show("Fehlende Übersetzungen exportiert in : " + MisTranslateFile);
        //=================================================================
        //Fehlworteliste lesen und Zeilenanzahl ermitteln :
        int counter = 0;

        if (File.Exists(MisTranslateFile))
        {
            using (StreamReader countReader = new StreamReader(MisTranslateFile))
            {
                while (countReader.ReadLine() != null)
                    counter++;
            }
            // MessageBox.Show("Zeilenanzahl in " + MisTranslateFile + " : " + counter);
            if (counter > 1)
            //=================================================================
            //Fehlworteliste öffnen falls Zeilenanzahl > 1 :
            {
                // MessageBox.Show("Fehlende Übersetzungen gefunden !");       
                // Open the txt file with missing translation
                System.Diagnostics.Process.Start("notepad.exe", MisTranslateFile);
            }
        }
        progress.EndPart(true);
        return;
    }
    //=======================================================================
    public string Get_Project()
    {
        try
        {
            // aktuelles Projekt ermitteln
            //==========================================
            Eplan.EplApi.ApplicationFramework.ActionManager oMngr = new Eplan.EplApi.ApplicationFramework.ActionManager();
            Eplan.EplApi.ApplicationFramework.Action oSelSetAction = oMngr.FindAction("selectionset");
            string sProjektT = "";
            if (oMngr != null)
            {
                Eplan.EplApi.ApplicationFramework.ActionCallingContext ctx = new Eplan.EplApi.ApplicationFramework.ActionCallingContext();
                ctx.AddParameter("TYPE", "PROJECT");
                bool sRet = oSelSetAction.Execute(ctx);
                if (sRet)
                { ctx.GetParameter("PROJECT", ref sProjektT); }
                //MessageBox.Show("Projekt: " + sProjektT);
            }
            return sProjektT;
        }
        catch
        { return ""; }
    }
    //################################################################################################
    public string Get_Name(string sProj)
    {
        try
        {
            // Projektname ermitteln
            //==========================================
            int i = sProj.Length - 5;
            string sTemp = sProj.Substring(1, i);
            i = sTemp.LastIndexOf(@"\");
            sTemp = sTemp.Substring(i + 1);
            //MessageBox.Show("Ausgabe: " + sTemp);
            return sTemp;
        }
        catch
        { return "ERROR"; }
    }
}