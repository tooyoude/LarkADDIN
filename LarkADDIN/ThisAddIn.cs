using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace LarkADDIN
{
    
    public partial class ThisAddIn
    {
        private Word.Application wordapp;
        private Word.Document doc;
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.wordapp = this.Application;
            this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(WorkWithDocument);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(WorkWithDocument);

        }

        private void WorkWithDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            try
            {
                this.doc = Doc;

                Word.Range rng = Doc.Range(0, 0);
                
                rng.Text = "Hello World";

                this.doc.PageSetup.TopMargin = this.wordapp.CentimetersToPoints(float.Parse("2.54"));// 上边距
                this.doc.PageSetup.BottomMargin = this.wordapp.CentimetersToPoints(float.Parse("2.54"));// 下边距
                this.doc.PageSetup.LeftMargin = this.wordapp.CentimetersToPoints(float.Parse("4.17"));// 左边距
                this.doc.PageSetup.RightMargin = this.wordapp.CentimetersToPoints(float.Parse("3.17"));// 右边距 

                this.doc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;// 页面为纵向
                this.doc.PageSetup.GutterPos = Word.WdGutterStyle.wdGutterPosLeft;//装订线位于左侧
                this.doc.PageSetup.PageWidth = this.wordapp.CentimetersToPoints(float.Parse("21"));// 纸张宽度
                this.doc.PageSetup.PageHeight = this.wordapp.CentimetersToPoints(float.Parse("29.7"));// 纸张高度

                this.doc.PageSetup.HeaderDistance = this.wordapp.CentimetersToPoints(float.Parse("1.5"));//页眉顶端距离
                this.doc.PageSetup.FooterDistance = this.wordapp.CentimetersToPoints(float.Parse("1.75"));//页脚底端距离


                //rng.Select();
            }
            catch (Exception ex)
            {
                // Handle exception if for some reason the document is not available.
            }
        }

        private void 

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
