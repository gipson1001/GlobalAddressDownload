using System.Diagnostics;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace GlobalAddressListSample
{
	[DebuggerDisplay("DisplayName = {DisplayName}, Mail = {Mail}")]
	public class Address
	{
		internal Address(DirectoryEntry entry,int index,int size)
		{
			email = (string)entry.Properties["mail"].Value;
			department = (string)entry.Properties["department"].Value;
			company = (string)entry.Properties["company"].Value;
			string lastName1 = (string)entry.Properties["sn"].Value;
			string lastName2 = (string)entry.Properties["name"].Value;
			if(email != null && email.Contains("quantacn")) {
				//return;
			}
			if(email != null && email.Equals("Wenchi.Lan@quantatw.com")) {
			}
			if(company != null && !company.Contains("QCI") && !company.Contains("qci")) {
			}

			if(department != null && department.EndsWith("-QSMC")) {
				//return;
			}
			object directReports = entry.Properties["directReports"].Value;
			if(directReports != null) {
				try {
					object[] tmp;
					if(directReports.GetType().FullName.Equals("System.String")) {
						tmp = new object[1];
						tmp[0] = directReports;
					} else {
						tmp = (object[])directReports;
					}
					reported = new Person[tmp.Length];
					for (int i = 0; i < tmp.Length; i++) {
						reported[i] = new Person();
						string man = tmp[i] as string;
						if(man.Contains(" (")) {
							int l = man.IndexOf(" (");
							int r = man.IndexOf(')');
							reported[i].name = man.Substring(3, l - 3);
							reported[i].chtName = man.Substring(l + 2, r - l - 2);
						} else if(man.Contains(")(")) {
							int l = man.IndexOf('(');
							int m = man.IndexOf(")(");
							int r = man.IndexOf("),");
							reported[i].name = man.Substring(3, l - 3);
							reported[i].chtName = man.Substring(l + 1, m - l - 1);
							reported[i].alias = man.Substring(m + 2, r - m - 2);
						} else if(!man.Contains("(") && !man.Contains(")")) {
							reported[i].name = man.Substring(3, man.IndexOf(',') - 3);
						} else if(man.Contains("(")) {
							int l = man.IndexOf('(');
							int r = man.IndexOf(')');
							reported[i].name = man.Substring(3, l - 3);
							reported[i].chtName = man.Substring(l + 1, r - l - 1);
						} else {							
						}
					}
				} catch {
				}
			}
			number = (string)entry.Properties["telephoneNumber"].Value;
			alias = (string)entry.Properties["mailNickname"].Value;
			titleName = (string)entry.Properties["title"].Value;
			if(company == null && titleName == null) {
				return;
			}
			/*try {
			if(company != null) {
				foreach (string k in entry.Properties.PropertyNames) {
					System.Console.WriteLine(k + " : " + entry.Properties[k].Value);
				}
			}
			}catch{
				
			}*/
			if(titleName == null) {
				System.Console.WriteLine("[" + titleName + "][" + number + "][" + company + "][" + lastName1 + "]");
				return;
			}
			int left = titleName.IndexOf('(');
			int right = titleName.IndexOf(')');					
			if(left == -1) {
				if(titleName == null || number == null || company == null || lastName1 == null) {
					System.Console.WriteLine("[" + titleName + "][" + number + "][" + company + "][" + lastName1 + "]");
					return;
				} else {
					titleName = "不明";
				}
			}
			name = (string)entry.Properties["title"].Value;
			if(name.Contains(" ")) {
				name = name.Substring(0, name.IndexOf(' '));
			}
			if(name.Contains("\u6052")) {
				name = name.Replace("\u6052", "\u6046");
			}
			if(left >= 0 && right >= 0) {
				titleName = titleName.Substring(left + 1, right - left - 1);
			}
			firstName = (string)entry.Properties["displayName"].Value;
			if(firstName.Contains("(")) {
				//System.Console.WriteLine(firstName);
				if(firstName.IndexOf('(') == 0) {
					firstName = (string)entry.Properties["sn"].Value;
					firstName = firstName.Substring(0, firstName.IndexOf('.'));
				} else {
					firstName = firstName.Substring(0, firstName.IndexOf('('));
				}
			}
			lastName = "";
			if(lastName2.Contains("(")) {
				lastName2 = lastName2.Substring(0, lastName2.LastIndexOf('(')).Trim(new char[]{ ' ' });
			}
			if(lastName2.Contains(" ")) {
				lastName2 = lastName2.Substring(lastName2.LastIndexOf(' ') + 1);
			}
			if(lastName2.Contains(".")) {
				lastName2 = lastName2.Substring(lastName2.LastIndexOf('.') + 1);
			}
			if(!lastName1.Equals(lastName2) && lastName2.Length > 0) {
				lastName2 = lastName2.Substring(0, 1).ToUpper() + lastName2.Substring(1).ToLower();
			}
			if(!lastName1.Equals(lastName2)) {
				if(name.Contains("鍾") || name.Contains("鐘")) {
					lastName2 = lastName1;
				}
			}
			if(firstName.Contains(" " + lastName1)) {
				firstName = firstName.Replace(" " + lastName1, "");
				lastName = lastName1;
			} else if(firstName.Contains(lastName1 + " ")) {
				firstName = firstName.Replace(lastName1 + " ", "");
				lastName = lastName1;
			} else if(firstName.Contains(" " + lastName2)) {
				firstName = firstName.Replace(" " + lastName2, "");
				lastName = lastName2;
			} else if(firstName.Contains(lastName2 + " ")) {
				firstName = firstName.Replace(lastName2 + " ", "");
				lastName = lastName2;
			}
			firstName = firstName.Trim(new char[]{ ' ' });
			if(name.Equals("\u9577\u5b89\u5f18\u52f3")) {
				firstName = "Hiroisa";
				lastName = "Nagayasu";
			}
			object obj = entry.Properties["msExchUMDtmfMap"];
			office = (string)entry.Properties["physicalDeliveryOfficeName"].Value;
			string manager = (string)entry.Properties["manager"].Value;
			if(entry.Properties["thumbnailPhoto"].Value != null) {
				//byte[] bytes = (byte[])entry.Properties["thumbnailPhoto"].Value;
				//outputJpg(bytes, alias + ".jpg");
			}
			if(manager != null && manager.Contains(",")) {
				manager = manager.Substring(3, manager.IndexOf(',') - 3);
			} else {
				/*foreach (string k in entry.Properties.PropertyNames) {
					System.Console.WriteLine(k + " : " + entry.Properties[k].Value);
				}*/
				manager = "--";
			}
			boss = new Person();
			if(manager.Contains("-")) {
				manager = manager.Replace("-", " ");
			}
			if(manager.Contains("Yung Chang Chin")) {
				manager = manager.Replace("Yung Chang Chin", "Y. C. Chin");
			}
			if(manager.Contains(" (") && manager.Contains(")")) {
				int leftP = manager.IndexOf(" (");
				int rightP = manager.IndexOf(')');
				boss.name = manager.Substring(0, leftP).Replace(" ", "").Replace(".", "").ToLower();
				boss.chtName = manager.Substring(leftP + 2, rightP - leftP - 2);
			} else if(manager.Contains("(") && manager.Contains(")")) {
				int leftP = manager.IndexOf('(');
				int rightP = manager.IndexOf(')');
				boss.name = manager.Substring(0, leftP).Replace(" ", "").Replace(".", "").ToLower();			
				boss.chtName = manager.Substring(leftP + 1, rightP - leftP - 1);
			} else if(manager != null) {
				boss.name = manager.Replace(" ", "").Replace(".", "").ToLower();
			}
			if(boss.chtName != null && boss.name.Equals("at") && boss.chtName.Equals("\u8521\u80b2\u5b97")) {
				boss.name = "attsai";
			}
			if(boss.chtName != null && boss.name.Equals("aaron") && boss.chtName.Equals("\u8b1d\u5fb7\u8cb4")) {
				boss.name = "aaronhsieh";
			}
			if(boss.name.Equals("jwhwu") && boss.chtName == null) {
				boss.name = "jeffreyhwu";
			}
			if(boss.chtName != null && boss.name.Equals("jrwenjeng") && boss.chtName.Equals("\u912d\u667a\u6587")) {
				boss.name = "jwjeng";
			}
			if(boss.chtName != null && boss.name.Equals("vencentwcc") && boss.chtName.Equals("\u738b\u5efa\u660c")) {
				boss.name = "vencentwang";
			}
			if(boss.chtName != null && boss.name.Equals("joeycchen") && boss.chtName.Equals("\u9673\u8000\u632f")) {
				boss.name = "joechen";
			}
			if(name != null) {
				string msg = "Processing " + index + " / " + size + " " + firstName + " " + lastName + "         ";
				System.Console.Write(msg);
				for (int i = 0; i < msg.Length; i++) {
					System.Console.Write("\b");			
				}
			}
		}
		
		public string alias;
		public string number;
		public string office;
		public Person boss;
		public Person[] reported;
		public string email;
		public string department;
		public string company;
		public string titleName;
		public string lastName;
		public string firstName;
		public string name;

		private static void outputJpg(byte[] data, string fileName)
		{
			using (Image image = Image.FromStream(new MemoryStream(data))) {
				image.Save("D:\\jpg\\" + fileName, ImageFormat.Jpeg);
			}
		}
	}
	
	public class Person
	{
		public string alias;
		public string name;
		public string chtName;
	}
}
