using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace tesseract0413
{
  public  class 提取表格类
    {
        public List<Bitmap> ConvertImage2form(Bitmap srcImage1)
        {

            Mat srcImage=new Mat();
            //Mat srcImage2 = new Mat();

            var bb = BitmapToBytes(srcImage1);

            srcImage = Mat.ImDecode(bb);

            List<Bitmap> listBitmaps=new List<Bitmap>();

            //List<Rect> list01 = new List<Rect>();
            //List<Rect> list02 = new List<Rect>();
            //List<Rect> list03 = new List<Rect>();
            //List<Rect> list04 = new List<Rect>();

            //Mat srcImage = Cv2.ImRead(@"D:\A1.Jpeg");

            //Cv2.GaussianBlur(srcImage, srcImage,new OpenCvSharp.Size(3, 3), 0, 0);

            ////拉普拉斯锐化
            //Mat kernel =new Mat(3, 3, MatType.CV_32F, new Scalar(-1));
            //kernel.At<float>(1, 1) X = 8.9;

            Mat dsrImage = new Mat();
            //缩小
            OpenCvSharp.Size dsize = new OpenCvSharp.Size(0.5 * srcImage.Cols, 0.5 * srcImage.Rows);
            Cv2.Resize(srcImage, dsrImage, dsize, 0, 0, InterpolationFlags.Linear);
            //var window缩小 = new Window("缩小");
            //window缩小.Image = dsrImage;

            //灰度
            Mat dsrImage1 = new Mat();
            Cv2.CvtColor(dsrImage, dsrImage1, ColorConversionCodes.BGR2GRAY);
            //var window灰度 = new Window("灰度");
            //window灰度.Image = dsrImage1;

            //二值化后的图片是黑底白字
            Mat bw1 = new Mat();

            Mat bw = new Mat();


            Cv2.Threshold(~dsrImage1, bw, 90, 255, ThresholdTypes.Binary);
            Cv2.AdaptiveThreshold(~dsrImage1, bw1, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 15, -2);

            var window直接阈值化 = new Window("直接阈值化");
            window直接阈值化.Image = bw1;

            var window自适应阈值化 = new Window("自适应阈值化");
            window自适应阈值化.Image = bw;

            //查找轮廓
            OpenCvSharp.Point[][] contours = null;
            HierarchyIndex[] hierarcy = null;
            Cv2.FindContours(bw1, out contours, out hierarcy, RetrievalModes.External, ContourApproximationModes.ApproxNone);

            Mat RoiSrcImg1 = new Mat(bw1.Rows, bw1.Cols, MatType.CV_8UC3);

            //创建一个旋转后的图像  
            Mat RatationedImg = new Mat(bw1.Rows, bw1.Cols, MatType.CV_8UC1);

            for (int i = 0; i < contours.Length; i++)
            {

                RotatedRect rect = Cv2.MinAreaRect(contours[i]);

                Rect rect1 = Cv2.BoundingRect(contours[i]);




                RoiSrcImg1.Rectangle(rect1, Scalar.Red);

                //Cv2.rectangle(image, (x, y), (x + w, y + h), (153, 153, 0), 5)




                if (rect.Size.Height * rect.Size.Width > 8000)
                {

                    RoiSrcImg1.Rectangle(rect1, Scalar.Blue);

                    float angle = rect.Angle;
                    float angle1 = 0;

                    if (angle != 0.0)
                    {
                        // 此处可通过cv::warpAffine进行旋转矫正，本例不需要
                        bw1.Rectangle(rect1, Scalar.Red);
                        if (angle > 0)
                        {
                            angle1 = angle;
                        }
                        else
                        {
                            angle1 = 90 + angle;
                        }

                    }

                    ////Mat RoiSrcImg(srcImg.rows, srcImg.cols, CV_8UC3); //注意这里必须选CV_8UC3
                    //Mat RoiSrcImg = new Mat(bw1.Rows, bw1.Cols, MatType.CV_8UC3);
                    //RoiSrcImg.SetTo(0);


                    //Cv2.DrawContours(bw1, contours, -1, new Scalar(255), Cv2.FILLED);

                    //dsrImage.CopyTo(RoiSrcImg, bw1);


                    //var window感兴趣的区域 = new Window("感兴趣的区域");
                    //window感兴趣的区域.Image = RoiSrcImg;



                    RatationedImg.SetTo(0);

                    Point2f center = rect.Center;  //中心点  
                    Mat M2 = Cv2.GetRotationMatrix2D(center, angle1, 1);//计算旋转加缩放的变换矩阵 

                    Cv2.WarpAffine(bw1, RatationedImg, M2, bw1.Size(), InterpolationFlags.Linear, 0, new Scalar(0));//仿射变换 

                    var window旋转之后 = new Window("旋转之后");
                    window旋转之后.Image = RatationedImg;
                }
            }

            Mat horizontal = RatationedImg.Clone();
            Mat vertical = RatationedImg.Clone();

            int scale = 20; //这个值越大，检测到的直线越多
            int horizontalsize = horizontal.Cols / scale;

            // 为了获取横向的表格线，设置腐蚀和膨胀的操作区域为一个比较大的横向直条
            Mat horizontalStructure = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(horizontalsize, 1));

            // 先腐蚀再膨胀
            Cv2.Erode(horizontal, horizontal, horizontalStructure, new OpenCvSharp.Point(-1, -1));
            Cv2.Dilate(horizontal, horizontal, horizontalStructure, new OpenCvSharp.Point(-1, -1));

            //imshow("horizontal", horizontal);
            var window先腐蚀再膨胀横向 = new Window("先腐蚀再膨胀横向");
            window先腐蚀再膨胀横向.Image = horizontal;



            int verticalsize = vertical.Rows / scale;
            Mat verticalStructure = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(1, verticalsize));
            Cv2.Erode(vertical, vertical, verticalStructure, new OpenCvSharp.Point(-1, -1));
            Cv2.Dilate(vertical, vertical, verticalStructure, new OpenCvSharp.Point(-1, -1));
            var window先腐蚀再膨胀纵向 = new Window("先腐蚀再膨胀纵向");
            window先腐蚀再膨胀纵向.Image = vertical;

            Mat mask = horizontal + vertical;
            var window横纵向 = new Window("横纵向");
            window横纵向.Image = mask;

            Mat joints = new Mat();
            Cv2.BitwiseAnd(horizontal, vertical, joints);
            var window交点 = new Window("交点");
            window交点.Image = joints;


            OpenCvSharp.Point[][] contours1 = null;
            HierarchyIndex[] hierarcy1 = null;
            Cv2.FindContours(mask, out contours1, out hierarcy1, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);

            Rect[] rects;
            //ArrayList list = new ArrayList();

            List<Rect> list = new List<Rect>();

            for (int i = 0; i < contours1.Length; i++)
            {

                //RotatedRect rect = Cv2.MinAreaRect(contours1[i]);

                Rect rect1 = Cv2.BoundingRect(contours1[i]);
                list.Add(rect1);
                //var window = new Window(i.ToString());


                //Mat image_roi = new Mat(RatationedImg, rect1);

                //window.Image = image_roi;

            }


            list.Sort((a, b) => a.Y.CompareTo(b.Y));


            List<Rect> list01 = new List<Rect>();
            List<Rect> list02 = new List<Rect>();

            List<Rect> list03 = new List<Rect>();

            List<Rect> list04 = new List<Rect>();



            for (int i = 1; i < 5; i++)
            {
                list01.Add(list[i]);
            }

            for (int i = 5; i < 9; i++)
            {
                list02.Add(list[i]);
            }

            for (int i = 9; i < 13; i++)
            {
                list03.Add(list[i]);
            }

            for (int i = 13; i < 17; i++)
            {
                list04.Add(list[i]);
            }


            list01.Sort((a, b) => a.X.CompareTo(b.X));
            list02.Sort((a, b) => a.X.CompareTo(b.X));
            list03.Sort((a, b) => a.X.CompareTo(b.X));
            list04.Sort((a, b) => a.X.CompareTo(b.X));


            for (int i = 1; i < 4; i += 2)
            {
                ;

                var window = new Window("01" + i.ToString());


                Mat image_roi = new Mat(RatationedImg, list01[i]);

                byte[] bytes = image_roi.ImEncode(".jpg");

                var  x=   BytesToBitmap(bytes);

                listBitmaps.Add(x);

                window.Image = image_roi;


            }


            for (int i = 1; i < 4; i += 2)
            {
                ;

                var window = new Window("02" + i.ToString());


                Mat image_roi = new Mat(RatationedImg, list02[i]);


                byte[] bytes = image_roi.ImEncode(".jpg");

                var x = BytesToBitmap(bytes);

                listBitmaps.Add(x);

                window.Image = image_roi;


            }

            for (int i = 1; i < 4; i += 2)
            {
                ;

                var window = new Window("03" + i.ToString());


                Mat image_roi = new Mat(RatationedImg, list03[i]);

                byte[] bytes = image_roi.ImEncode(".jpg");

                var x = BytesToBitmap(bytes);

                listBitmaps.Add(x);

                window.Image = image_roi;


            }


            for (int i = 1; i < 4; i += 2)
            {
                ;

                var window = new Window("04" + i.ToString());


                Mat image_roi = new Mat(RatationedImg, list04[i]);

                byte[] bytes = image_roi.ImEncode(".jpg");

                var x = BytesToBitmap(bytes);

                listBitmaps.Add(x);

                window.Image = image_roi;


            }






            return listBitmaps;
        }

        public static Bitmap BytesToBitmap(byte[] Bytes)
        {
            MemoryStream stream = null;
            try
            {
                stream = new MemoryStream(Bytes);
                return new Bitmap((Image)new Bitmap(stream));
            }
            catch (ArgumentNullException ex)
            {
                throw ex;
            }
            catch (ArgumentException ex)
            {
                throw ex;
            }
            finally
            {
                stream.Close();
            }
        }

        public static byte[] BitmapToBytes(Bitmap Bitmap)
        {
            MemoryStream ms = null;
            try
            {
                ms = new MemoryStream();
                Bitmap.Save(ms, Bitmap.RawFormat);
                byte[] byteImage = new Byte[ms.Length];
                byteImage = ms.ToArray();
                return byteImage;
            }
            catch (ArgumentNullException ex)
            {
                throw ex;
            }
            finally
            {
                ms.Close();
            }
        }

    }
}
