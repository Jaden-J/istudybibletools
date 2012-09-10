﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleVersePointer
{
    public static class Logger
    {
        public static bool WasLogged = false;

        public static void Initialize()
        {
            WasLogged = false;
        }

        public static void LogError(string message)
        {
            MessageBox.Show(message);
            WasLogged = true;
        }
    }
}