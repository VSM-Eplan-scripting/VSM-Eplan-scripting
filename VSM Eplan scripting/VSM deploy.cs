/*
 * Created by SharpDevelop.
 * User: m.pluimers
 * Date: 11-5-2021
 * Time: 16:00
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;

public class VSM_deploy
{
	public static void Copy(string sourceDirectory, string targetDirectory)
	{
		DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
		DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

		// get the file attributes for file or directory
		FileAttributes attr = File.GetAttributes(@sourceDirectory);

		if (attr.HasFlag(FileAttributes.Directory))
		{
			//It's a folder
			CopyAll(diSource, diTarget);
		}
		else
		{
			//It's a file
			if (File.Exists(@targetDirectory))
			{
				File.Delete(@targetDirectory);
			}
			File.Copy(sourceDirectory, targetDirectory);
		}
	}

	public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
	{
		Directory.CreateDirectory(target.FullName);

		// Copy each file into the new directory.
		foreach (FileInfo fi in source.GetFiles())
		{
			Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
			MessageBox.Show(Path.Combine(target.FullName, fi.Name));
			if (File.Exists(Path.Combine(target.FullName, fi.Name)))
			{

				File.Delete(Path.Combine(target.FullName, fi.Name));
			}
			fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
		}

		// Copy each subdirectory using recursion.
		foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
		{
			DirectoryInfo nextTargetSubDir =
				target.CreateSubdirectory(diSourceSubDir.Name);
			CopyAll(diSourceSubDir, nextTargetSubDir);
		}
	}

	public string StripIllegalChars(string _input)
	{
		return string.Join("", _input.Split(Path.GetInvalidFileNameChars()));
	}

	[DeclareEventHandler("onAddOnRegistered.String.App")]
	public void IncrementalDeploy(IEventParameter iEventParameter)
	{
		EventParameterString oEventParameterString = new EventParameterString(iEventParameter);
		string Parameter = oEventParameterString.String;
		if (Parameter == @"\\vsm-fs-svr03\Data\Eplan Electric P8\Gegevens\Auto roll-out Eplan settings")
		{
			DeployProgressUI();
		}
	}

	public void DeployProgressUI()
	{
		Form form = new Form();
		TextBox textBox = new TextBox();

		int i = 0;
		int WaitingTimeInSec = 5;

		string XmlLocation = @"\\vsm-fs-svr03\data\Eplan Electric P8\Gegevens\Auto roll-out Eplan settings\DeployData.xml";

		string ProgressText = "VSM File deploy tool reporting:" + Environment.NewLine;
		ProgressText = ProgressText + "Deploy XML file:" + XmlLocation + Environment.NewLine;

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

		//Load XML		
		XmlDocument doc = new XmlDocument();
		doc.Load(XmlLocation);

		form.Show();

		XmlNodeList xmlnode;
		xmlnode = doc.GetElementsByTagName("FileFolder");

		//Loop through files and folders		
		for (i = 0; i <= xmlnode.Count - 1; i++)
		{
			string Source = xmlnode[i].ChildNodes.Item(1).InnerText.Replace("%username%", Environment.UserName);
			string Target = xmlnode[i].ChildNodes.Item(2).InnerText.Replace("%username%", Environment.UserName);
			ProgressText = ProgressText + "Subject: " + xmlnode[i].ChildNodes.Item(0).InnerText + Environment.NewLine;
			ProgressText = ProgressText + "Copy from: " + Source + Environment.NewLine;
			ProgressText = ProgressText + "Copy to: " + Target + Environment.NewLine;
			textBox.Text = ProgressText;
			form.Update();

			Copy(Source, Target);
		}

		//The end
		textBox.Text = ProgressText + "Closing window in " + WaitingTimeInSec + " seconds";
		form.Update();

		Wait(WaitingTimeInSec);
		form.Close();
	}

	public void Wait(int time)
	{
		Thread thread = new Thread(delegate ()
		{
			System.Threading.Thread.Sleep(time * 1000);
		});
		thread.Start();
		while (thread.IsAlive)
			Application.DoEvents();
	}
}