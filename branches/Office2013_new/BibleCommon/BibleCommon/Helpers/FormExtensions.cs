using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Management;
using System.Security;
using System.Diagnostics;
using System.Security.Principal;
using System.Reflection;
using System.Threading;

namespace BibleCommon.Helpers
{
    public static class FormExtensions
    {
        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // When you don't want the ProcessId, use this overload and pass IntPtr.Zero for the second parameter
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        /// <summary>The GetForegroundWindow function returns a handle to the foreground window.</summary>
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool BringWindowToTop(HandleRef hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        public static void SetFocus(this Form form)
        {
            uint foreThread = GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero);
            uint appThread = GetCurrentThreadId();
            const uint SW_SHOW = 5;
            if (foreThread != appThread)
            {
                AttachThreadInput(foreThread, appThread, true);
                BringWindowToTop(form.Handle);
                ShowWindow(form.Handle, SW_SHOW);
                AttachThreadInput(foreThread, appThread, false);
            }
            else
            {
                BringWindowToTop(form.Handle);
                ShowWindow(form.Handle, SW_SHOW);
            }
            form.Activate();
        }


        public static void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            }
            else
            {
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
            }
        }

        public static void Invoke(this Control control, Action action)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new MethodInvoker(action), null);
            }
            else
            {
                action.Invoke();
            }
        }

        public static void RunSingleInstance(string mutexId, string messageIfSecondInstance, Action singleAction, bool silent = false, string additionalMutexId = null)
        {
            string appGuid = ((GuidAttribute)Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(GuidAttribute), false).GetValue(0)).Value.ToString();

            Mutex mutex = null;
            Mutex additionalMutex = null;

            try
            {
                mutex = new Mutex(false, string.Format("Global\\{0}_{1}", appGuid, mutexId));
                if (!string.IsNullOrEmpty(additionalMutexId))
                    additionalMutex = new Mutex(false, string.Format("Global\\{0}_{1}", appGuid, additionalMutexId));
                
                if (mutex.WaitOne(0, false) && (additionalMutex == null || additionalMutex.WaitOne(0, false)))
                {
                    singleAction();                    
                }
                else
                {
                    if (!silent)
                        MessageBox.Show(messageIfSecondInstance, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    BibleCommon.Services.Logger.Done();
                }
            }
            finally
            {
                if (mutex != null)
                    mutex.Close();

                if (additionalMutex != null)
                    additionalMutex.Close();
            }
        }

        public static void SetToolTip(Control c, string toolTip)
        {
            var _toolTip = new ToolTip();

            _toolTip.AutoPopDelay = 5000;
            _toolTip.InitialDelay = 1000;
            _toolTip.ReshowDelay = 500;
            _toolTip.ShowAlways = true;

            _toolTip.SetToolTip(c, toolTip);
        }

        public static void EnableAll(bool enabled, Control.ControlCollection controls, params Control[] except)
        {
            foreach (Control control in controls)
            {
                EnableAll(enabled, control.Controls, except);

                if (!except.Contains(control))
                    control.Enabled = enabled;
            }
        }
    }
}




