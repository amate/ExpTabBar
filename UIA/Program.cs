using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace UIA
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetScrollPos((IntPtr)0x002206D2);
            SetScrollPos((IntPtr)0x02010FEC, 6000);
        }

        [DllExport]
        public static int GetScrollPos(IntPtr hwndListView)
        {
            AutomationElement scrollbarElm = null;
            try
            {
                var listViewElm = AutomationElement.FromHandle(hwndListView);
                var condScrollbarCtrl = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ScrollBar);
                scrollbarElm = listViewElm.FindFirst(TreeScope.Children, condScrollbarCtrl);
                if (scrollbarElm == null)
                {
                    return 0;
                }
                var rangeValuePattern = (RangeValuePattern)scrollbarElm.GetCurrentPattern(RangeValuePattern.Pattern);
                //Console.WriteLine("Hello World! {0}", rangeValuePattern.Current.Value);
                //rangeValuePattern.SetValue(50000);
                return (int)rangeValuePattern.Current.Value;
            }
            catch (Exception )
            {
                try
                {
                    var valuePattern = (ValuePattern)scrollbarElm.GetCurrentPattern(ValuePattern.Pattern);
                    int pos = 0;
                    int.TryParse(valuePattern.Current.Value, out pos);
                    return pos;
                } catch(Exception )
                {
                    return 0;
                }
            }
        }

        [DllExport]
        public static bool SetScrollPos(IntPtr hwndListView, int scrollPos)
        {
            AutomationElement scrollbarElm = null;
            try
            {
                var listViewElm = AutomationElement.FromHandle(hwndListView);
                var condScrollbarCtrl = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ScrollBar);
                scrollbarElm = listViewElm.FindFirst(TreeScope.Children, condScrollbarCtrl);
                if (scrollbarElm == null)
                {
                    return false;
                }
                var rangeValuePattern = (RangeValuePattern)scrollbarElm.GetCurrentPattern(RangeValuePattern.Pattern);
                rangeValuePattern.SetValue(scrollPos);
                return true;
            }
            catch(Exception )
            {
                try
                {
                    //var valuePattern = (ValuePattern)scrollbarElm.GetCurrentPattern(ValuePattern.Pattern);
                    //valuePattern.SetValue(scrollPos.ToString());
                }
                catch (Exception)
                {
                }
                return false;
            }
        }

        [DllExport]
        public static double GetUIItemsViewVerticalScrollPercent(IntPtr hwndUIItemsView)
        {
            try
            {
                var UIItemsView = AutomationElement.FromHandle(hwndUIItemsView);
                var scrollPattern = (ScrollPattern)UIItemsView.GetCurrentPattern(ScrollPattern.Pattern);
                if (!scrollPattern.Current.VerticallyScrollable)
                {
                    return 0.0;
                }
                return scrollPattern.Current.VerticalScrollPercent;

            }
            catch(Exception)
            {
                return 0.0;
            }

        }

        [DllExport]
        public static void SetUIItemsViewVerticalScrollPercent(IntPtr hwndUIItemsView, double virticalScrollPercent)
        {
            try
            {
                var UIItemsView = AutomationElement.FromHandle(hwndUIItemsView);
                var scrollPattern = (ScrollPattern)UIItemsView.GetCurrentPattern(ScrollPattern.Pattern);
                scrollPattern.SetScrollPercent(ScrollPattern.NoScroll, virticalScrollPercent);
            }
            catch (Exception)
            {
            }
        }
    }
}
