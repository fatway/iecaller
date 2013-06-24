using System;
using System.Collections.Generic;
using System.Text;
using mshtml;

namespace IECall
{
    class ieHelper
    {

        /// <summary>
        /// 获取当前所有已打开的IE窗口的浏览URL
        /// </summary>
        /// <returns>List&lt;string&rt;</returns>
        public static List<string> GetIhtmlUrls()
        {
            List<string> urls = new List<string>();

            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindowsClass();
            foreach (SHDocVw.InternetExplorer ie in shellWindows)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
                if (filename.Equals("iexplore"))
                {
                    urls.Add(ie.LocationURL);
                }
            }
            return urls;
        }



        /// <summary>
        /// 返回指定Url的IE窗口下的 IHTMLDocument2 对象。
        /// </summary>
        /// <returns>IHTMLDocument2</returns>
        public static IHTMLDocument2 GetIHTMLDocument2ByUrl(string url)
        {
            SHDocVw.ShellWindows shellWindows = new SHDocVw.ShellWindowsClass();

            foreach (SHDocVw.InternetExplorer ie in shellWindows)
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(ie.FullName).ToLower();
                if (filename.Equals("iexplore") && ie.LocationURL == url)
                {
                    return ie.Document as IHTMLDocument2;
                }
            }

            return null;
        }



        /// <summary>
        /// 填写表单内容
        /// </summary>
        /// <param name="htmlDoc">已获取的IE窗体</param>
        /// <param name="editName">表单控件名称（如username）</param>
        /// <param name="editValue">需要填写的内容</param>
        public static void InputValueOnIHTMLDocument(IHTMLDocument2 htmlDoc, string editName, string editValue)
        {
            IHTMLInputElement input = (IHTMLInputElement)htmlDoc.all.item(editName, 0);
            input.value = editValue;
        }



        /// <summary>
        /// 点击已获取网页内的指定按钮
        /// </summary>
        /// <param name="htmlDoc">已获取的IE窗体</param>
        /// <param name="btnName">按钮显示名称</param>
        /// <param name="tagName">按钮标签（默认为input）</param>
        /// <returns>bool</returns>
        public static bool PressBtnInHTMLDocument(IHTMLDocument2 htmlDoc, string btnName, string tagName="input")
        {
            HTMLDocumentClass obj = (HTMLDocumentClass)htmlDoc;
            IHTMLElement iHTMLElement = null;
            IHTMLElementCollection c = obj.getElementsByTagName(tagName);

            foreach (IHTMLElement element in c)
            {
                if (element.outerHTML.IndexOf(btnName) != -1)
                {
                    iHTMLElement = element;
                    break;
                }
            }

            if (iHTMLElement != null)
            {
                iHTMLElement.click();
                return true;
            }

            return false;
        }

    }
}
