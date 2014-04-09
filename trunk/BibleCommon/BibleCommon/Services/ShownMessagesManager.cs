using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Services
{
    public static class ShownMessagesManager
    {
        public static class MessagesCodes
        {
            public static int SuggestUsingFolderForNotesPages = 1;
            public static int NewVersionInfo = 2;
        }

        public static bool GetMessageWasShown(int code)
        {
            return SettingsManager.Instance.ShownMessages.Contains(code);
        }

        public static void SetMessageWasShown(int code)
        {
            if (!SettingsManager.Instance.ShownMessages.Contains(code))
                SettingsManager.Instance.ShownMessages.Add(code);            
        }

        public static void ClearMessageWasShown(int code)
        {
            if (SettingsManager.Instance.ShownMessages.Contains(code))
                SettingsManager.Instance.ShownMessages.Remove(code);
        }
    }
}
