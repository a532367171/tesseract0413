using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using O2S.Components.PDFRender4NET;
using Tesseract;

namespace tesseract0413
{
    public partial class Form1 : Form
    {

        提取表格类 _提取表格类;
        public Form1()
        {
            InitializeComponent();
            _提取表格类 = new 提取表格类();

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)

            {

                var img = new Bitmap(openFileDialog1.FileName);

                // var ocr = new TesseractEngine(@"C:\Program Files (x86)\Tesseract-OCR\tessdata", "eng", EngineMode.TesseractAndCube);

                var ocr = new TesseractEngine("./tessdata", "chi_sim", EngineMode.TesseractOnly);

                //ocr.SetVariable("tessedit_char_whitelist", "0123456789");

                var page = ocr.Process(img);

                txtResult.Text = page.GetText();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<List<Bitmap>> xList1 = new List<List<Bitmap>>();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //ConvertPDF2Image(openFileDialog1.FileName, "D:\\", "A", 1, 5, System.Drawing.Imaging.ImageFormat.Jpeg,Definition.Five);
                xList1= ConvertPDF2ListImage(openFileDialog1.FileName, "D:\\", "A", 1, 5, System.Drawing.Imaging.ImageFormat.Jpeg, Definition.Five);
            }
        }



        public enum Definition

        {

            One = 1, Two = 2, Three = 3, Four = 4, Five = 5, Six = 6, Seven = 7, Eight = 8, Nine = 9, Ten = 10

        }

        /// <summary>

        /// 将PDF文档转换为图片的方法

        /// </summary>

        /// <param name="pdfInputPath">PDF文件路径</param>

        /// <param name="imageOutputPath">图片输出路径</param>

        /// <param name="imageName">生成图片的名字</param>

        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>

        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>

        /// <param name="imageFormat">设置所需图片格式</param>

        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>

        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath, string imageName, int startPageNum, int endPageNum, System.Drawing.Imaging.ImageFormat imageFormat, Definition definition)

        {

            PDFFile pdfFile = PDFFile.Open(pdfInputPath);

            if (!Directory.Exists(imageOutputPath))

            {

                Directory.CreateDirectory(imageOutputPath);

            }

            // validate pageNum

            if (startPageNum <= 0)

            {

                startPageNum = 1;

            }

            if (endPageNum > pdfFile.PageCount)

            {

                endPageNum = pdfFile.PageCount;

            }

            if (startPageNum > endPageNum)

            {

                int tempPageNum = startPageNum;

                startPageNum = endPageNum;

                endPageNum = startPageNum;

            }

            // start to convert each page

            for (int i = startPageNum; i <= endPageNum; i++)
            {


                Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * (int)definition);

                pageImage.Save(imageOutputPath + imageName + i.ToString() + "." + imageFormat.ToString(), imageFormat);

                pageImage.Dispose();

            }

            pdfFile.Dispose();

        }

        public  List<List<Bitmap>> ConvertPDF2ListImage(string pdfInputPath, string imageOutputPath, string imageName, int startPageNum, int endPageNum, System.Drawing.Imaging.ImageFormat imageFormat, Definition definition)
        {

            List<List<Bitmap>> xList = new List<List<Bitmap>>();

            PDFFile pdfFile = PDFFile.Open(pdfInputPath);

            if (!Directory.Exists(imageOutputPath))

            {

                Directory.CreateDirectory(imageOutputPath);

            }

            // validate pageNum

            if (startPageNum <= 0)

            {

                startPageNum = 1;

            }

            if (endPageNum > pdfFile.PageCount)

            {

                endPageNum = pdfFile.PageCount;

            }

            if (startPageNum > endPageNum)

            {

                int tempPageNum = startPageNum;

                startPageNum = endPageNum;

                endPageNum = startPageNum;

            }

            // start to convert each page

            for (int i = startPageNum; i <= endPageNum; i++)
            {


                Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * (int)definition);

                var x = _提取表格类.ConvertImage2form(pageImage);
                //pageImage.Save(imageOutputPath + imageName + i.ToString() + "." + imageFormat.ToString(), imageFormat);

                xList.Add(x);

                pageImage.Dispose();

            }

            pdfFile.Dispose();
            return xList;
        }



    }
}


