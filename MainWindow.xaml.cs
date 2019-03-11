using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Text;

namespace GlobalAddressListSample
{
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			DateTime start = DateTime.Now;
			Address[] addressList = GetGlobalAddressList();
			var relationMap = new Dictionary<string, Person>(); // alias -> boss
			var nameMap = new Dictionary<string, List<string>>(); // name -> alias
			var titleMap = new Dictionary<string, string>(); // alias -> title
			var addressMap = new Dictionary<string, Address>(); // name -> alias
			foreach (Address item in addressList) {
				if(item.alias != null && item.boss != null) {
					relationMap.Add(item.alias, item.boss);
					string newKey1 = item.firstName + " " + item.lastName + item.name;
					string newKey2 = item.firstName + " " + item.lastName;
					newKey1 = newKey1.Replace("-", " ").Replace(" ", "").Replace(".", "").ToLower();
					newKey2 = newKey2.Replace("-", " ").Replace(" ", "").Replace(".", "").ToLower();
					if(!nameMap.ContainsKey(newKey1)) {
						var aliasList = new List<string>();
						nameMap.Add(newKey1, aliasList);
					}
					if(!nameMap.ContainsKey(newKey2)) {
						var aliasList = new List<string>();
						nameMap.Add(newKey2, aliasList);
					}
					nameMap[newKey1].Add(item.alias);
					nameMap[newKey2].Add(item.alias);
					titleMap.Add(item.alias, item.titleName);
					addressMap.Add(item.alias, item);
				}
			}
			foreach (string item in relationMap.Keys) {
				if(nameMap.ContainsKey(relationMap[item].name + relationMap[item].chtName)) {
					var list = nameMap[relationMap[item].name + relationMap[item].chtName];
					if(list.Count == 1) {
						relationMap[item].alias = list[0];
						continue;
					} else {
						foreach (string alias in list) { // make sure the same depart for relation
							if(addressMap[alias].department.Equals(addressMap[item].department)) {
								relationMap[item].alias = alias;
								break;
							}
							if(addressMap[alias].office.Equals(addressMap[item].office)) {
								relationMap[item].alias = alias;
								break;
							}
							if(addressMap[item].department.Contains(addressMap[alias].department)) {
								relationMap[item].alias = alias;
								break;
							}
						}
						if(relationMap[item].alias == null) {
							foreach (string alias in list) { // make sure the same depart for relation
								if(getWeight(titleMap[item]) < getWeight(titleMap[alias])) {
									relationMap[item].alias = alias;
									break;
								}
							}
						} else {
							continue;
						}
					}
				}
				if(nameMap.ContainsKey(relationMap[item].name)) {
					var list = nameMap[relationMap[item].name];
					if(list.Count == 1) {
						relationMap[item].alias = list[0];
						continue;
					} else {
						foreach (string alias in list) { // make sure the same depart for relation
							if(addressMap[alias].department.Equals(addressMap[item].department)) {
								relationMap[item].alias = alias;
								break;
							}
							if(addressMap[alias].office.Equals(addressMap[item].office)) {
								relationMap[item].alias = alias;
								break;
							}
							if(addressMap[item].department.Contains(addressMap[alias].department)) {
								relationMap[item].alias = alias;
								break;
							}
						}
						if(relationMap[item].alias == null) {
							foreach (string alias in list) { // make sure the same depart for relation
								if(getWeight(titleMap[item]) < getWeight(titleMap[alias])) {
									relationMap[item].alias = alias;
									break;
								}
							}
						} else {
							continue;
						}
					}
				}
			}
			var finalList = new List<Address>();
			foreach (Address item in addressList) {
				if(item.email == null) {
					continue;
				}
				if(item.alias == null) {
					continue;
				}
				if(item.alias != null && relationMap.ContainsKey(item.alias) && relationMap[item.alias].alias != null) {
					item.boss.alias = relationMap[item.alias].alias;
				} else {
					
				}
				finalList.Add(item);
			}
			outputList(finalList);
			DateTime end = DateTime.Now;
			System.Console.WriteLine(end.Subtract(start).TotalSeconds);
			this.InitializeComponent();
			Close();
		}
		
		private static void outputList(List<Address> list)
		{
			list.Sort(delegate(Address a, Address b) {
				if(a.alias == null && b.alias == null)
					return 0;
				else if(a.alias == null)
					return -1;
				else if(b.alias == null)
					return 1;
				else
					return a.alias.CompareTo(b.alias);
			});
			DateTime now = DateTime.Now;
			string fileName = "quanta_" + string.Format("{0:0000}{1:00}{2:00}{3:00}{4:00}", now.Year, now.Month, now.Day, now.Hour, now.Minute) + ".csv";
			var sb = new StringBuilder();
			foreach (Address item in list) {
				if(item.number != null && item.name != null && item.firstName != null && item.firstName.Length > 0) {
					sb.Append("\"" + item.alias + "\",");
					sb.Append("\"" + item.email + "\",");
					if(item.number.Length >= 3) {
						sb.Append("\"" + item.number + "\",");
					} else {
						sb.Append("\"\",");
					}
					sb.Append("\"" + item.office + "\",");
					sb.Append("\"" + (item.boss.alias??"--") + "\",");
					sb.Append("\"" + item.firstName + "\",");
					sb.Append("\"" + item.lastName + "\",");
					sb.Append("\"" + item.titleName + "\",");
					sb.Append("\"" + item.name + "\",");
					sb.Append("\"" + item.department + "\"");
					sb.AppendLine();
				}
			}
			File.WriteAllText(fileName, sb.ToString());
		}

		public static Address[] GetGlobalAddressList()
		{
			using (var searcher = new DirectorySearcher()) {
				using (var entry = new DirectoryEntry(searcher.SearchRoot.Path)) {
					searcher.Filter = "(&(mailnickname=*)(objectClass=user))";
					searcher.PropertiesToLoad.Add("cn");
					searcher.PropertyNamesOnly = true;
					searcher.SearchScope = SearchScope.Subtree;
					searcher.Sort.Direction = SortDirection.Ascending;
					searcher.Sort.PropertyName = "cn";
					SearchResultCollection ret = searcher.FindAll();
					int size = ret.Count;
					int index = 0;
					return ret.Cast<SearchResult>().Select(result => 
						new Address(result.GetDirectoryEntry(), index++, size)
					).ToArray();
				}
			}
		}
		
		public static int getWeight(String t)
		{
			switch (t) {
				case "董事長":
					return 1000;
				case "副董事長":
					return 833;
				case "CTO VP&GM":
					return 666;
				case "VP&GM":
					return 666;
				case "AVP&GM":
					return 666;
				case "Sr. VP&GM":
					return 666;
				case "資深副總經理":
					return 583;
				case "副總經理暨財務長":
				case "執行副總經理":
					return 533;
				case "副總經理":
					return 500;
				case "協理":
				case "技術協理":
				case "專案協理":
					return 417;
				case "顧問":
				case "顧問A":
				case "顧問B":
				case "處長":
				case "技術處長":
				case "專案處長":
					return 367;
				case "副處長":
				case "專案副處長":
				case "技術副處長":
					return 333;
				case "部經理":
				case "專案部經理":
				case "技術部經理":
					return 300;
				case "資深經理":
				case "專案資深經理":
				case "技術資深經理":
					return 267;
				case "經理":
				case "專案經理":
				case "技術經理":
					return 233;
				case "副理":
				case "技術副理":
				case "專案副理":
					return 200;
				case "一級專員":
				case "資深組長":
					return 167;
				case "二級專員":
				case "儲備幹部":
					return 150;
				case "工程師":
				case "管理師":
					return 133;
				case "助理工程師":
				case "助理管理師":
				case "管理員":
				case "組長":
					return 100;
				case "技術員":
					return 83;
				case "助理技術員":
					return 67;
				case "工讀生":
					return 33;
				default: 
					return 100;
			}
		}
	}
}
