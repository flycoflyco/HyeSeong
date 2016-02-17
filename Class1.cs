using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
//여기 수정
using EXCEL = Microsoft.Office.Interop.Excel;


using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Interop;


using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.Runtime;
using Autodesk.Civil.Settings;
using Autodesk.Civil.DatabaseServices.Styles;




namespace ClassLibrary1
{

   // public delegate void StrAddHandler(String str);
     
    public class Class1 : Form1  // 오늘 회의 필요한 서류가 하나라면 그 서류만 달라고 마누라에게 전화를 하고..
    {                               // 가방에 필요한 서류가 이것 저것 많을 땐 가방채로 메고 출근을 해야할듯... 너무 많이 필요해
        Form1 dlg;
       // Alignment currAl = new Alignment();

        //public static event StrAddHandler ItemStr;  // 테스트용

       // public event EventHandler test;
        

        [DllImport("User32.dll")]
        private static extern int MessageBox(int h, string m, string c, int type);


        [CommandMethod("ShowDialog")]
      /*  public static void STR()
        {
            ItemStr("abcde");
        }*/
        public void ShowDialog1()
        {
           
            if (dlg == null)
            {
                dlg = new Form1();
                dlg.Show();
            }
            dlg.Focus();



            dlg.Click3 += openFilePointGroup; // 점 추가하기 버튼 눌렀을시 실행되는 메쏘드 (Class.cs -> form1)




           Autodesk.AutoCAD.ApplicationServices.Application.ShowModelessDialog(dlg);
        }
        public void openFilePointGroup(object obj, EventArgs e)
        {
            OpenFileDialog OPFile = new OpenFileDialog();
            String fName;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument; // 현재 도큐먼트를 가져옴
            DocumentLock loc = doc.LockDocument();
            CivilDocument Cdoc = CivilApplication.ActiveDocument; // tinsurface 위해서 생성


            if (OPFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                System.IO.StreamReader sr = new
                System.IO.StreamReader(OPFile.FileName);
                //System.Windows.Forms.MessageBox.Show(OPFile.FileName);//OPFile.FileName //System.Windows.Forms.MessageBox.Show(sr.ReadToEnd());
                fName = OPFile.FileName;
                sr.Close();
                using (Transaction ts1 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
                {
                    PointFileFormatCollection ptFileFormats = PointFileFormatCollection.GetPointFileFormats(HostApplicationServices.WorkingDatabase); // 현재 가지고 있는 파일 포맷을 다 가져옴
                    PointFileFormat ptFormatId = ptFileFormats["PNEZD(쉼표 구분)"]; // 그중 하나를 선택
                    uint result = CogoPointCollection.ImportPoints(fName, ptFormatId); // 파일 이름이랑 포맷형식을 인자로 가지고 점그룹 추가
                    ts1.Commit();
                }
                using (Transaction ts2 = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction()) // 지표면에 원래 있던 점그룹을 추가하는 트랜잭션
                {
                    ObjectId surfaceStyleID = Cdoc.Styles.SurfaceStyles[3]; // 지표면 스타일 (이름을 쓰거나 인덱스)

                    ObjectId pointG = Cdoc.PointGroups.AllPointsPointGroupId; // 모든점을 ObjectId 클래스 변수에 저장

                    ObjectId surfId = TinSurface.Create("기본지표면", surfaceStyleID); // surfID에 지표면 이름이랑 스타일 값을 인자로 넘겨주고 지표면 생성


                    TinSurface Tsur = surfId.GetObject(OpenMode.ForRead) as TinSurface; // ObjectId 값을 as 로 명시적 형태 변환 해서 사용

                    Tsur.PointGroupsDefinition.AddPointGroup(pointG); // 모든점을 지금 추가한 지표면에 추가함

                    ts2.Commit(); // 두번째 트랜잭션 시작
                }
                Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen();
                doc.SendStringToExecute("._zoom _e", true, false, true); // 리습명령어로 줌 사용

            }
            else
            {
                System.Windows.Forms.MessageBox.Show("파일을 다시 열어주세요");
            }
        }
        public void testpara(ListView mylist)
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            DocumentLock loc = acDoc.LockDocument();
            Polyline test = new Polyline();

            using (loc)
            {
                Database acCurDb = acDoc.Database;
                PromptPointOptions pPtOpts = new PromptPointOptions(" 점을 입력하세요");
                PromptPointResult pPtRes = acDoc.Editor.GetPoint(pPtOpts);// PromptPointResult pPtRes = acDoc.Editor.GetPoint(pPtOpts); // 점을 사용자에게 가져온다
                Point2dCollection ptStart = new Point2dCollection(); // Point3d 의 배열형이라 생각하면 쉽다.

                pPtOpts.BasePoint = pPtRes.Value; // 기준점 (점선)
                pPtOpts.UseBasePoint = true; 

                //pPtOpts.UseDashedLine = true;
                int count = 0; // 점갯수 count

                ptStart.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));

                while (pPtRes.Status == PromptStatus.OK) // 사용자한테 받는값이 있을때까지
                {
                    if (pPtRes.Status == PromptStatus.Cancel) { return; } // 캔슬시 리턴
                    pPtOpts.BasePoint = pPtRes.Value;
                    pPtOpts.Message = " 점을 입력하세요"; // ㅇㅇ 
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts);  // 사용자에게 점값을 받아온다 

                    ptStart.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y)); // 그값을 콜렉션에다가 배열로 저장
                    count++;

                }

                test.SetDatabaseDefaults(); // 폴리라인 초기화(버튼 두번누르면 초기화 안되므로)

                for (int i = 0; i < count; i++)
                {
                    test.AddVertexAt(i, ptStart[i], 0, 0, 0);
                }

                //          test.Closed = true; // 폴리라인을 폐합선으로 만들어줌 (도형)
                //         


                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction()) // 트랜잭션으로 묵어줌 라인그리기 위해서
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                 OpenMode.ForRead) as BlockTable;

                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    acBlkTblRec.AppendEntity(test); // 폴리라인을 블럭테이블 레코드에 그려준다.
                    acTrans.AddNewlyCreatedDBObject(test, true);//디비 데이터에 넣어줌
                    acTrans.Commit(); // 트랜잭션 실행
                }

            }
            Autodesk.AutoCAD.ApplicationServices.Application.UpdateScreen(); // 화면 한번 업뎃 해주고

            Alignment al = this.CreateAlign(test); // 만들어논 메소드를 이용해서 만들어진 평면선형 클래스의 변수값을 가져온다.

                int temp=0;

                temp = 0;
                foreach(AlignmentEntity myAe in al.Entities) //  일반 라인과 호의 순서가 맞진 않지만 정보는 다 맞음
                {                                                   ///추후에 정교한 작업 필요할듯
                    temp++;
                    string msg1 = "";
                    mylist.BeginUpdate();
                    ListViewItem testLvi2 = new ListViewItem(string.Format("{0}",temp));   // 리스트뷰 값 넣어주는 부분
                    switch(myAe.EntityType)
                    {
                        case AlignmentEntityType.Line:
                            AlignmentLine myLine = myAe as AlignmentLine;
                            msg1 = myLine.StartPoint.ToString(); // 시작점
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myLine.EndPoint.ToString(); // 끝점
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myLine.Length.ToString(); // 길이
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myLine.StartStation.ToString(); // 시작 스테이션
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myLine.EndStation.ToString(); // 끝 스테이션
                            testLvi2.SubItems.Add(msg1);


                            mylist.Items.Add(testLvi2);
                            break;
                        case AlignmentEntityType.Arc:
                            AlignmentArc myArc = myAe as AlignmentArc;
                            msg1 = myArc.StartPoint.ToString(); // 시작점
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myArc.EndPoint.ToString(); // 끝점
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myArc.Length.ToString(); // 길이
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myArc.StartStation.ToString(); // 시작 스테이션
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myArc.EndStation.ToString(); // 끝 스테이션
                            testLvi2.SubItems.Add(msg1);
                            msg1 = myArc.Radius.ToString(); // r 값
                            testLvi2.SubItems.Add(msg1);


                            mylist.Items.Add(testLvi2);
                            break;
                        default:
                            mylist.Items.Add("");
                            break;
                    }

                    mylist.EndUpdate(); // 포문안에 beginupdata 가 있으므로 마찬가지로 안쪽에 위치해줘야한다.
                }
             



        }
        public Alignment CreateAlign(Polyline guid)  // 만들어진 폴리라인을 가지고 선형을 만들어주는 메소드 선형뿐만아니라 그리드 뷰까지 찍어줌
        {
            Document dc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument; // 현재 도큐먼트를 가져옴
            Database db = dc.Database; // 현재 데이터베이스를 가져옴

           
            CivilDocument doc = CivilApplication.ActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                BlockTable acBlktbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable; // 선을 그리기 위한 블럭테이블 생성
                BlockTableRecord acblkTblrec = acTrans.GetObject(acBlktbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord; // 어디 도면에 그릴지 선택


                PolylineOptions plops = new PolylineOptions(); // 폴리라인 옵션 지정

                plops.AddCurvesBetweenTangents = true;
                plops.EraseExistingEntities = true;
                plops.PlineId = guid.ObjectId;

                ObjectId testAlignmentID = Alignment.Create(doc, plops, "내가만든 선형", null, "0", "Proposed", "All Labels");


                Alignment oAlignment = acTrans.GetObject(testAlignmentID, OpenMode.ForRead) as Alignment;

                
                ObjectId layerId = oAlignment.LayerId;
                // get first surface in the document
                ObjectId surfaceId = doc.GetSurfaceIds()[0];
                // get first style in the document
                ObjectId styleId = doc.Styles.ProfileStyles[0];
                // get the first label set style in the document
                ObjectId labelSetId = doc.Styles.LabelSetStyles.ProfileLabelSetStyles[0];


                try
                {
                    ObjectId profileId = Profile.CreateFromSurface("My Profile", testAlignmentID, surfaceId, layerId, styleId, labelSetId);

                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    ed.WriteMessage(e.Message);
                }


                PromptPointResult pPtRes;
                PromptPointOptions pPtOpts = new PromptPointOptions("종단뷰를 그릴 위치를 찍어주세요~");
                pPtRes = dc.Editor.GetPoint(pPtOpts);
                Point3d ptStart = pPtRes.Value;
                if (pPtRes.Status == PromptStatus.Cancel) return null;

                // ObjectId ProfileViewId = ProfileView.Create(alignID, ptStart);
                ObjectId pfrVBSStyleId = doc.Styles.ProfileViewBandSetStyles[8];

                ObjectId ProfileViewId2 = ProfileView.Create(doc, "My Profile", pfrVBSStyleId, testAlignmentID, ptStart);
                //doc, "My Profile View", pfrVBSStyleId, alignID, ptInsert
                acTrans.Commit();
                return oAlignment;
            }
            

        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // lv1
            // 
            this.lv1.SelectedIndexChanged += new System.EventHandler(this.lv1_SelectedIndexChanged);
            // 
            // Class1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.ClientSize = new System.Drawing.Size(678, 406);
            this.Name = "Class1";
            this.ResumeLayout(false);

        }

        private void lv1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
