using System;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using System.Text.RegularExpressions; // regex
using OpenNETCF.Win32; // RegistryKey
using System.Runtime.InteropServices; // DllImport
using System.Reflection; // Assembly

namespace Mossywell.CleanActiveSync
{
	#region Structs
	public struct RegistryValues
	{
		public string RegValue;
		public object RegData;

		public RegistryValues(string regvalue, object regdata)
		{
			RegValue = regvalue;
			RegData = regdata;
		}
	}
	#endregion

	#region Enums
	public enum RegistryTree
	{
		LocalMachine,
		CurrentUser,
	}
	#endregion

	public class MainForm : System.Windows.Forms.Form
	{
		#region Externals
		const uint EM_REPLACESEL = 0xC2;
		[DllImport("coredll")]
		extern static IntPtr GetFocus();
		[DllImport("coredll")]
		extern static int SendMessage(IntPtr hWnd, uint Msg, bool WParam, string LParam);
		#endregion

		#region Class Fields
		private int _warnings;
		private int _errors;

		private System.Windows.Forms.MenuItem menuItemGo;
		private System.Windows.Forms.MenuItem menuItemExit;
		private System.Windows.Forms.MainMenu mainMenu;
		private System.Windows.Forms.TextBox textBoxOutput;
		private RegistryValues[] novalues = {
		};
		private RegistryValues[] cuinboxsyncserviceprovidersactivesync = {
		  new RegistryValues("Disabled", 0x00000000U),
			new RegistryValues("DLL", "mailtrns.dll"),
			new RegistryValues("Email", 0x00000001U),
		  new RegistryValues("TYPE", "ActiveSync"),
	  };
		private RegistryValues[] lmproviderscalendar = {
			new RegistryValues("Enabled", 0x00000000U),
			new RegistryValues("ManagerDriven", 0x00000001U),
			new RegistryValues("Name", "Calendar"),
			new RegistryValues("Priority", 80U),
		};
		private RegistryValues[] lmproviderscontacts = {
			new RegistryValues("Enabled", 0x00000000U),
			new RegistryValues("ManagerDriven", 0x00000001U),
			new RegistryValues("Name", "Contacts"),
			new RegistryValues("Priority", 80U),
		};
		private RegistryValues[] lmprovidersinbox = {
			new RegistryValues("Enabled", 0x00000000U),
			new RegistryValues("ManagerDriven", 0x00000001U),
			new RegistryValues("Name", "Inbox"),
			new RegistryValues("Priority", 64U),
		};
		private RegistryValues[] regcalendarcontactsmailvalues = {
			new RegistryValues("SyncSwitchPurge", 0x00000000U),
		};
		private RegistryValues[] regsettingsvalues = {
			new RegistryValues("ConflictResolution", 0x00000001U),
			new RegistryValues("WindowSize", 0x00000064U),
			new RegistryValues("OutboundMailDelay", 0x00000005U),
			new RegistryValues("SyncHierarchy", 0x00000001U),
			new RegistryValues("IncludeRemoteManualSync", 0x00000000U),
			new RegistryValues("IncludeRemoteSync", 0x00000000U),
			new RegistryValues("SyncAfterTimeWhenCradled", 0x00000005U),
			new RegistryValues("AutoSyncWhenCradled", 0x00000001U),
			new RegistryValues("DeviceAddressingMethod", 0x00000002U),
			new RegistryValues("DevicePhoneNumber", ""),
			new RegistryValues("CarrierConnectorList", new string[] {"","",""}),
			new RegistryValues("CarrierConnector", ""),
			new RegistryValues("DeviceSMSAddress", ""),
			new RegistryValues("SendMailItemsImmediately", 0x00000000U),
			new RegistryValues("MSAS-ProtocolVersions", ""),
			new RegistryValues("ClientNegotiated", 0x00000000U),
			new RegistryValues("ClientProtocolVersion", "1.0"),
			new RegistryValues("LaxFrequency", 0x000005a1U),
			new RegistryValues("PeakDays", 0x0000003eU),
			new RegistryValues("PeakEndTime", 0x00000438U),
			new RegistryValues("PeakStartTime", 0x000001e0U),
			new RegistryValues("SyncWhenRoaming", 0x00000000U),
			new RegistryValues("OffPeakFrequency", 0x00000000U),
			new RegistryValues("PeakFrequency", 0x00000000U),
			new RegistryValues("SyncAfterTime", 0x0000001eU),
			new RegistryValues("SyncAfterCount", 0x00000000U),
			new RegistryValues("AutoSync", 0x00000000U),
			new RegistryValues("MailBodyTruncation", 0x00000200U),
			new RegistryValues("DisconnectWhenDone", 0x00000001U),
			new RegistryValues("MailFileAttachments", 0x00000000U),
			new RegistryValues("NotificationsEnabled", 0x00000001U),
			new RegistryValues("SaveDeletedItems", 0x00000001U),
			new RegistryValues("SaveSentItems", 0x00000001U),
			new RegistryValues("BodyTruncation", 0x00001400U),
			new RegistryValues("CalendarAgeFilter", 0x00000004U),
			new RegistryValues("EmailAgeFilter", 0x00000002U),
		};
		private RegistryValues[] regconnectionvalues = {
			new RegistryValues("LstActvSncErrTime", 0x00000000U),
			new RegistryValues("UseSSL", 0x00000001U),
			new RegistryValues("UseURIAsSupplied", 0x00000000U),
			new RegistryValues("UseSSL", 0x00000001U),
			new RegistryValues("URI", "Microsoft-Server-ActiveSync"),
			new RegistryValues("DeviceID", ""),
			new RegistryValues("Device", "SmartPhone"),
			new RegistryValues("ConnectionID", ""),
			new RegistryValues("Domain", ""),
			new RegistryValues("Server", ""),
			new RegistryValues("SavePassword", 0x00000000U),
			new RegistryValues("Password", ""),
			new RegistryValues("User", ""),
		};
		private RegistryValues[] regloggingvalues = {
			new RegistryValues("NumberOfLogs", 0x00000001U),
			new RegistryValues("Enabled", 0x00000000U),
		};
		#endregion

		#region Constructor
	  public MainForm()
		{
			InitializeComponent();
		}
		#endregion

		#region Dispose
		protected override void Dispose( bool disposing )
		{
			base.Dispose( disposing );
		}
		#endregion

		#region Windows Form Designer generated code
		private void InitializeComponent()
		{
			this.mainMenu = new System.Windows.Forms.MainMenu();
			this.menuItemGo = new System.Windows.Forms.MenuItem();
			this.menuItemExit = new System.Windows.Forms.MenuItem();
			this.textBoxOutput = new System.Windows.Forms.TextBox();
			// 
			// mainMenu
			// 
			this.mainMenu.MenuItems.Add(this.menuItemGo);
			this.mainMenu.MenuItems.Add(this.menuItemExit);
			// 
			// menuItemGo
			// 
			this.menuItemGo.Text = "&Go";
			this.menuItemGo.Click += new System.EventHandler(this.menuItemGo_Click);
			// 
			// menuItemExit
			// 
			this.menuItemExit.Text = "E&xit";
			this.menuItemExit.Click += new System.EventHandler(this.menuItemExit_Click);
			// 
			// textBoxOutput
			// 
			this.textBoxOutput.BackColor = System.Drawing.Color.White;
			this.textBoxOutput.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular);
			this.textBoxOutput.Location = new System.Drawing.Point(-1, 0);
			this.textBoxOutput.Multiline = true;
			this.textBoxOutput.ReadOnly = true;
			this.textBoxOutput.Size = new System.Drawing.Size(188, 180);
			this.textBoxOutput.Text = "";
			// 
			// MainForm
			// 
			this.Controls.Add(this.textBoxOutput);
			this.Menu = this.mainMenu;
			this.Text = "Clean ActiveSync";
			this.Load += new System.EventHandler(this.MainForm_Load);

		}
		#endregion

		#region Main Entry Point
		static void Main() 
		{
			Application.Run(new MainForm());
		}
		#endregion

		#region Events
		private void MainForm_Load(object sender, System.EventArgs e)
		{
			_warnings = 0;
			_errors = 0;
			string version = Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() + "." + 
				Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString();
			AddToTextBox("Clean ActiveSync " + version + " loaded. Click \"Go\" to start.\r\n\r\n" +
				"WARNING: This is irreversable, so make sure that you:\r\n" +
        "  1. Really want to do this!\r\n" +
        "  2. Have backed up the phone's registry.\r\n\r\n" +
				"Disclaimer: Although I've not had any problems running this on my own phone, I can't " +
        "guarantee that you won't, so running this is your responsibility, not mine, and all " +
        "the usual disclaimers apply here.");
		}

		private void menuItemExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void menuItemGo_Click(object sender, System.EventArgs e)
		{
			this.menuItemExit.Enabled = false;
			this.menuItemGo.Enabled = false;
			AddToTextBox("\r\nStarting...\r\n");
			DoIdent();
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\Inbox\SyncServiceProviders\ActiveSync", cuinboxsyncserviceprovidersactivesync);
			DoClearAndReplace(RegistryTree.LocalMachine, @"SOFTWARE\Microsoft\AirSync\SyncMgr\Providers\{208067F4-CD6D-4042-AA0D-EBEA00714318}", lmproviderscalendar);
			DoClearAndReplace(RegistryTree.LocalMachine, @"SOFTWARE\Microsoft\AirSync\SyncMgr\Providers\{3085FACE-4C83-497d-80EC-E2DE2B7295C9}", lmproviderscontacts);
			DoClearAndReplace(RegistryTree.LocalMachine, @"SOFTWARE\Microsoft\AirSync\SyncMgr\Providers\{944c16c0-dc2b-4c00-91b5-c914a440473b}", lmprovidersinbox);
      DoClearAndReplace(RegistryTree.LocalMachine, @"Software\Microsoft\Windows CE Services\Partners", novalues);
			DoClearAndReplace(RegistryTree.LocalMachine, @"Software\Microsoft\Windows CE Services\Partners\P1", novalues);
      DoClearAndReplace(RegistryTree.LocalMachine, @"Software\Microsoft\Windows CE Services\Partners\P2", novalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync", novalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Settings", regsettingsvalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Settings\Mail", regcalendarcontactsmailvalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Settings\Contacts", regcalendarcontactsmailvalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Settings\Calendar", regcalendarcontactsmailvalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Connection", regconnectionvalues);
			DoClearAndReplace(RegistryTree.CurrentUser, @"Software\Microsoft\AirSync\Logging", regloggingvalues);
			AddToTextBox("Errors: " + _errors.ToString() + ", Warnings: " + _warnings.ToString() + "\r\n");
			AddToTextBox("Finished! Please reboot your mobile device immediately.");
			this.menuItemGo.Enabled = true;
			this.menuItemExit.Enabled = true;
		}
    #endregion

		#region Utility Methods
		private void DoIdent()
		{
			// Hard-coded to sort out the phone's ident
		  RegistryKey rk = null;
			string strOrigName;

			rk = Registry.LocalMachine.OpenSubKey("Ident");
			// Did the key exist?
			if(rk == null)
			{
				_errors++;
				AddToTextBox("ERROR: The Ident registry key was not found and so is ignored.\r\n");
			}
			else
			{
				// Backup the orig name
				strOrigName = (rk.GetValue("OrigName", String.Empty)).ToString();

				// Was it found? If not we'll have to use a default value instead
				if(strOrigName == String.Empty)
				{
					_warnings++;
					AddToTextBox("WARNING: The Ident OrigName was not found, so using the " +
						"name \"MYDEVICE\" instead.");
					strOrigName = "MYDEVICE";
				}

				// Create the string array replacement values
				RegistryValues[] identvalues = {
				  new RegistryValues("Desc", ""),
					new RegistryValues("Name", strOrigName),
					new RegistryValues("OrigName", strOrigName),
					new RegistryValues("Username", "guest")
			  };

				// Clear out the key
				DoClearAndReplace(RegistryTree.LocalMachine, @"Ident", identvalues);
			}
		}

		private void DoClearAndReplace(RegistryTree tree, string keytodo, RegistryValues[] regvalues)
		{
			string keyname = "";
			RegistryKey rk = null;

			// Parse out the key name just for logging
			int pos = keytodo.LastIndexOf(@"\");
			if(pos == -1)
				keyname = keytodo;
			else
				keyname = keytodo.Substring(pos + 1);

			// Initialise the registry stuff
			switch(tree)
			{
				case RegistryTree.LocalMachine:
					rk = Registry.LocalMachine.OpenSubKey(keytodo, true);
					break;
				case RegistryTree.CurrentUser:
				  rk = Registry.CurrentUser.OpenSubKey(keytodo, true);
					break;
				default:
					_errors++;
					throw new ArgumentException("ERROR: Only LocalMachine and CurrentUser are supported.");
			}

			// Did the key exist?
			if(rk == null)
			{
				_warnings++;
				AddToTextBox("WARNING: The " + keyname + " registry key was not found and so is ignored.\r\n");
			}
			else
			{
				// Clear out all the values
				AddToTextBox("Clearing all " + keyname + " values...");
				foreach(string regvalue in rk.GetValueNames())
				{
					rk.DeleteValue(regvalue);
				}
				AddToTextBox(keyname + " values cleared.\r\n");

				// Now populate it with the new fresh values
				if(regvalues.Length != 0)
				{
					AddToTextBox("Adding new " + keyname + " values...");
					foreach(RegistryValues regval in regvalues)
					{
						object regdata = regval.RegData;
						switch(regdata.GetType().FullName)
						{
							case "System.String":   // SZ
							case "System.UInt32":   // DWORD
							case "System.String[]": // MULTI_SZ
								rk.SetValue(regval.RegValue, regval.RegData);
								break;
							default:
								_errors++;
								throw new ArgumentException("Only String, UInt32 and String[] are allowed.");
						}
					}
					AddToTextBox("New " + keyname + " values added.\r\n");

					// Now check what's in there
					AddToTextBox("This is what has been added...");
					foreach(string regval in rk.GetValueNames())
					{
						object regdata = rk.GetValue(regval);
						switch(regdata.GetType().FullName)
						{
							case "System.String": // SZ
								AddToTextBox(regval + ": \"" + regdata.ToString() + "\"");
								break;
							case "System.UInt32": // DWORD
								AddToTextBox(regval + ": " + regdata.ToString());
								break;
							case "System.String[]": // MULTI_SZ
								string str = "";
								foreach(string s in (string[])regdata)
								{
									str += "\"" + s + "\" ";
								}
								AddToTextBox(regval + ": " + str);
								break;
							default:
								_errors++;
								throw new Exception("Unknown type retrieved from registry.");
						}
					}
					AddToTextBox("\r\n");
				}
				rk.Close();
			}
		}

		private void AddToTextBox(string message)
		{
			this.textBoxOutput.Focus();
			IntPtr hWnd = GetFocus();
			SendMessage(hWnd, EM_REPLACESEL, false, message + "\r\n");
		}
		#endregion
	}
}
