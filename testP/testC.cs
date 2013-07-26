using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SilkTest.Ntf;

namespace testN {

	[TestClass]
	public class testC {

		private readonly Desktop _desktop = Agent.Desktop;

		[TestInitialize]
		public void Initialize() {
			BaseState baseState = new BaseState("%ProgramFiles%\\Microsoft Office\\Office12\\OUTLOOK.EXE", "/Window[@caption='Outlook Today - Microsoft Outlook']");
			baseState.WorkingDirectory = "C:\\Windows\\system32";
			baseState.Execute();
		}

		[TestMethod]
		public void outbox() {
			Window outlookTodayMicrosoftOutlook = _desktop.Window("@caption='*Microsoft Outlook'");
			//outlookTodayMicrosoftOutlook.Control("[@windowClassName='NetUIHWND'][1]").Click(MouseButton.Left, new Point(77, 139), ModifierKeys.None);
			outlookTodayMicrosoftOutlook.Control("[@windowClassName='NetUIHWND'][1]").Click(MouseButton.Left);
            //outlookTodayMicrosoftOutlook.Control("[*]").GetPropertyList()[1];

            Console.WriteLine(outlookTodayMicrosoftOutlook.Control("@windowClassName='NetUIHWND'"));
            
		}

	}

}