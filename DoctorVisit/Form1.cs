using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DoctorVisit.Utility;
using System.Data.OleDb;
using System.Drawing.Imaging;
using Spire.Pdf;
using Spire.Pdf.Grid;
using Spire.Pdf.Graphics;
using System.IO;
using System.Diagnostics;

namespace DoctorVisit
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }


        public string newline = Environment.NewLine;

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            string nodata = "";
            string[] cdcarray = new string[482] 
            {

"1113705267275",
"1113705267435",
"1113705268029",
"1113705268387",
"1113705274240",
"1113705274609",
"1113705274657",
"1113705274744",
"1113705275097",
"1113705277464",
"1113705277603",
"1113705278937",
"1113705279700",
"1113705279781",
"1113705280819",
"1113705283060",
"1113705287656",
"1113705289475",
"1113705289621",
"1113705289952",
"1113705339834",
"1113705346937",
"1113705347105",
"1113705347536",
"1113705348764",
"1113705348790",
"1113705350226",
"1113705350618",
"1113705350729",
"1113705351172",
"1113705351299",
"1113705352610",
"1113705353661",
"1113705356987",
"1113705357571",
"1113705358419",
"1113705358571",
"1113705361035",
"1113705361425",
"1113705363198",
"1113705365883",
"1113705370287",
"1113705370425",
"1113705372615",
"1113705377082",
"1113705378077",
"1113705378147",
"1113705378341",
"1113705381477",
"1113705382616",
"1113705382788",
"1113705385027",
"1113705386983",
"1113705388994",
"1113705389461",
"1113705391255",
"1113705391282",
"1113705391287",
"1113705391288",
"1113705392475",
"1113705393177",
"1113705393223",
"1113705394583",
"1113705394855",
"1113705395657",
"1113705396427",
"1113705396467",
"1113705397625",
"1113705397983",
"1113705398382",
"1113705398997",
"1113705399402",
"1113705399507",
"1113705401067",
"1113705402423",
"1113705405062",
"1113705406082",
"1113705406418",
"1113705408743",
"1113705410931",
"1113705411314",
"1113705411322",
"1113705411446",
"1113705415018",
"1113705415166",
"1113705415920",
"1113705417363",
"1113705418139",
"1113705419218",
"1113705497829",
"1113705497877",
"1113705497915",
"1113705498202",
"1113705498699",
"1113705500376",
"1113705502878",
"1113705503092",
"1113705503356",
"1113705503495",
"1113705503555",
"1113705503666",
"1113705504958",
"1113705504985",
"1113705505132",
"1113705513433",
"1113705535986",
"1113705536791",
"1113705539470",
"1113705544525",
"1113705545549",
"1113705545971",
"1113705546104",
"1113705546108",
"1113705546184",
"1113705546282",
"1113705548678",
"1113705548977",
"1113705549617",
"1113705549781",
"1113705552899",
"1113705553372",
"1113705557424",
"1113705558012",
"1113705559431",
"1113705560019",
"1113705560608",
"1113705560811",
"1113705561300",
"1113705561906",
"1113705561945",
"1113705563902",
"1113705564420",
"1113705565885",
"1113705566289",
"1113705566755",
"1113705567109",
"1113705567263",
"1113705569055",
"1113705569964",
"1113705570425",
"1113705571486",
"1113705571635",
"1113705573245",
"1113705573475",
"1113705573546",
"1113705573955",
"1113705574153",
"1113705574179",
"1113705574650",
"1113705574993",
"1113705575085",
"1113705575351",
"1113705575708",
"1113705578953",
"1113705579120",
"1113705579819",
"1113705581538",
"1113705584432",
"1113705585870",
"1113705587195",
"1113705587876",
"1113705591210",
"1113705592408",
"1113705593403",
"1113705593538",
"1113705594794",
"1113705595358",
"1113705595411",
"1113705595784",
"1113705596085",
"1113705597238",
"1113705598133",
"1113705599108",
"1113705600086",
"1113705601675",
"1113705604664",
"1113705607179",
"1113705607336",
"1113705607357",
"1113705609286",
"1113705609729",
"1113705610628",
"1113705611717",
"1113705611775",
"1113705613707",
"1113705614243",
"1113705616492",
"1113705618887",
"1113705619058",
"1113705621204",
"1113705622633",
"1113705622662",
"1113705623223",
"1113705623308",
"1113705623723",
"1113705623995",
"1113705624019",
"1113705627239",
"1113705632391",
"1113705634039",
"1113705635252",
"1113705635526",
"1113705636776",
"1113705637412",
"1113705639229",
"1113705640125",
"1113705640414",
"1113705641672",
"1113705642224",
"1113705642767",
"1113705643580",
"1113705643929",
"1113705645177",
"1113705645297",
"1113705646332",
"1113705646858",
"1113705647919",
"1113705648423",
"1113705657022",
"1113705657141",
"1113705664810",
"1113705664974",
"1113705665692",
"1113705665704",
"1113705665791",
"1113705665804",
"1113705665805",
"1113705668923",
"1113705670553",
"1113705670851",
"1113705678398",
"1113705678619",
"1113705678647",
"1113705679248",
"1113705679251",
"1113705679409",
"1113705680530",
"1113705682876",
"1113705686420",
"1113705686814",
"1113705687189",
"1113705687985",
"1113705688332",
"1113705690573",
"1113705732373",
"1113705737410",
"1113705737826",
"1113705738029",
"1113705740326",
"1113705741226",
"1113705743773",
"1113705748342",
"1113705749043",
"1113705749453",
"1113705749783",
"1113705751172",
"1113705751839",
"1113705752861",
"1113705753627",
"1113705753642",
"1113705754580",
"1113705755860",
"1113705764155",
"1113705764179",
"1113705772336",
"1113705772551",
"1113705776029",
"1113705778353",
"1113705778601",
"1113705780600",
"1113705780771",
"1113705782436",
"1113705782458",
"1113705782687",
"1113705782700",
"1113705782716",
"1113705783038",
"1113705783154",
"1113705783401",
"1113705784026",
"1113705784071",
"1113705784090",
"1113705784281",
"1113705784284",
"1113705784405",
"1113705786784",
"1113705787773",
"1113705787957",
"1113705788368",
"1113705788527",
"1113705788572",
"1113705789224",
"1113705789271",
"1113705790401",
"1113705792572",
"1113705795219",
"1113705796555",
"1113705797195",
"1113705799909",
"1113705799922",
"1113705800109",
"1113705800353",
"1113705800399",
"1113705800821",
"1113705801757",
"1113705803530",
"1113705804742",
"1113705805475",
"1113705805626",
"1113705806093",
"1113705809058",
"1113705811280",
"1113705811366",
"1113705812564",
"1113705812625",
"1113705812914",
"1113705813642",
"1113705813944",
"1113705814039",
"1113705815867",
"1113705816278",
"1113705817126",
"1113705826872",
"1113705827644",
"1113705827667",
"1113705827742",
"1113705828485",
"1113705828749",
"1113705829262",
"1113705829438",
"1113705829470",
"1113705829673",
"1113705830107",
"1113705830496",
"1113705830668",
"1113705830838",
"1113705831265",
"1113705832607",
"1113705838160",
"1113705839467",
"1113705840462",
"1113705843603",
"1113705844217",
"1113705844339",
"1113705844701",
"1113705848189",
"1113705851592",
"1113705852812",
"1113705853234",
"1113705853842",
"1113705854001",
"1113705855399",
"1113705855502",
"1113705855586",
"1113705855623",
"1113705855741",
"1113705855814",
"1113705855816",
"1113705856345",
"1113705860682",
"1113705862088",
"1113705862250",
"1113705862435",
"1113705864051",
"1113705864306",
"1113705865327",
"1113705866549",
"1113705866608",
"1113705866722",
"1113705868433",
"1113705874587",
"1113705877718",
"1113705879001",
"1113705880838",
"1113705882298",
"1113705882424",
"1113705882556",
"1113705883517",
"1113705885874",
"1113705886203",
"1113705887883",
"1113705888175",
"1113705888639",
"1113705888677",
"1113705889549",
"1113705889979",
"1113705912220",
"1113705912221",
"1113705912230",
"1113705912233",
"1113705912235",
"1113705912240",
"1113705912253",
"1113705912291",
"1113705912292",
"1113705912299",
"1113705912300",
"1113705912302",
"1113705912308",
"1113705912321",
"1113705912323",
"1113705912330",
"1113705912341",
"1113705912342",
"1113705912344",
"1113705912346",
"1113705912349",
"1113705912370",
"1113705912371",
"1113705912373",
"1113705915222",
"1113705915263",
"1113705915496",
"1113705915878",
"1113705915883",
"1113705915993",
"1113705917827",
"1113705918124",
"1113705919283",
"1113705919469",
"1113705919763",
"1113705921788",
"1113705922142",
"1113705922985",
"1113705922997",
"1113705925708",
"1113705926050",
"1113705926274",
"1113705926332",
"1113705926353",
"1113705926478",
"1113705927464",
"1113705929843",
"1113705929851",
"1113705929976",
"1113705930266",
"1113705930936",
"1113705930996",
"1113705935701",
"1113705937435",
"1113705939172",
"1113705939792",
"1113705941409",
"1113705942174",
"1113705942330",
"1113705942455",
"1113705944201",
"1113705945959",
"1113705947762",
"1113705948729",
"1113705949961",
"1113705951006",
"1113705951023",
"1113705953241",
"1113705954818",
"1113705956289",
"1113705956839",
"1113705964482",
"1113705966921",
"1113705967128",
"1113705967371",
"1113705967466",
"1113705967676",
"1113705967841",
"1113705968719",
"1113705969091",
"1113705969099",
"1113705969159",
"1113705969508",
"1113705976994",
"1113705977984",
"1113705978797",
"1113705978894",
"1113705979097",
"1113705979168",
"1113705979298",
"1113705979834",
"1113705979996",
"1113705980563",
"1113705981153",
"1113705982382",
"1113705983716"



        };
            try
            {
                string[] datanulllist = new string[] { };

                for (int i = 0; i< cdcarray.Length;i++)
                {
                    string cdcreportno = cdcarray[i];
                    Cursor.Current = Cursors.WaitCursor;                

                    string startdate = "111/08/01";
                    string enddate = "111/08/31";    
                    string documentdate = "";
                    string documentdateyear = "";
                    string documentdatemonth = "";
                    documentdateyear = startdate.Substring(0, 3);
                    documentdatemonth = startdate.Substring(4, 2); 
                    int dateyearpercent = Convert.ToInt32(documentdateyear) % 10;
                    string startdateyearlabel = "";

                    switch (dateyearpercent)
                    {
                        case 0:
                            startdateyearlabel = "J";
                            break;
                        case 1:
                            startdateyearlabel = "";
                            break;
                        case 2:
                            startdateyearlabel = "B";
                            break;
                        case 3:
                            startdateyearlabel = "C";
                            break;
                        case 4:
                            startdateyearlabel = "D";
                            break;
                        case 5:
                            startdateyearlabel = "E";
                            break;
                        case 6:
                            startdateyearlabel = "F";
                            break;
                        case 7:
                            startdateyearlabel = "G";
                            break;
                        case 8:
                            startdateyearlabel = "H";
                            break;
                        case 9:
                            startdateyearlabel = "I";
                            break;
                        default:                            
                            break;
                    }
                    string textlabel4 = Querydata1(cdcreportno, startdateyearlabel, documentdatemonth);
                    DataTable dg3 = new DataTable();
                    DataTable dg4 = new DataTable();
                    DataTable dg5 = new DataTable();
                    richTextBox3.BackColor = Color.LightGray;
                    richTextBox3.Text = "SUPER";
                    label4.Text = textlabel4;
                    richTextBox1.Text = startdate;
                    richTextBox2.Text = enddate;

                    richTextBox5.Text = Querydata2(cdcreportno);
                    richTextBox6.Text = Querydata3(cdcreportno, startdateyearlabel, documentdatemonth, documentdateyear);
                    richTextBox7.Text = Querydata4(cdcreportno, startdateyearlabel, documentdatemonth, documentdateyear);
                    dg3 = Querydata5(cdcreportno, startdateyearlabel, documentdatemonth, documentdateyear);
                    
                    if (dg3.Rows.Count != 0)
                    {
                        dataGridView3.DataSource = dg3;
                        try
                        {
                        dataGridView3.CurrentCell.Selected = false;
                        dataGridView3.AllowUserToAddRows = false;
                        dataGridView3.RowHeadersVisible = false;
                        }
                        catch(Exception ex)
                        {
                        string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                        MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        finally
                        {
                            ;
                        }
                    }
                    else
                    {
                        dataGridView3.DataSource = dg3;
                    }

                    dg4 = Querydata6(cdcreportno, startdateyearlabel, documentdatemonth, documentdateyear);      
                    if (dg4.Rows.Count != 0)
                    {
                        dataGridView4.DataSource = dg4;
                        try
                        {
                            dataGridView4.CurrentCell.Selected = false;
                            dataGridView4.AllowUserToAddRows = false;
                            dataGridView4.RowHeadersVisible = false;
                        }
                        catch (Exception ex)
                        {
                            string error =
                                    "{Message}" + ex.Message + newline +
                                    "{Stacktrace}" + ex.StackTrace + newline +
                                    "{Targetsite}" + ex.TargetSite + newline +
                                    "{Tostring}" + ex.ToString();

                            MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        finally
                        {
                            ;
                        }
                    }
                    else
                    {
                        dataGridView4.DataSource = dg4;
                    }

                    dg5 = Querydata7(cdcreportno, startdateyearlabel, documentdatemonth);
                    
                    if (dg5.Rows.Count != 0)
                    {
                        dataGridView5.DataSource = dg5;
                        try
                        {
                            dataGridView5.CurrentCell.Selected = false;
                            dataGridView5.AllowUserToAddRows = false;
                            dataGridView5.RowHeadersVisible = false;
                        }
                        catch (Exception ex)
                        {
                            string error =
                                    "{Message}" + ex.Message + newline +
                                    "{Stacktrace}" + ex.StackTrace + newline +
                                    "{Targetsite}" + ex.TargetSite + newline +
                                    "{Tostring}" + ex.ToString();

                            MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        finally
                        {
                            ;
                        }
                    }
                    else
                    {
                        dataGridView5.DataSource = dg5;
                    }

                    documentdate = Querydata8(cdcreportno, startdateyearlabel, documentdatemonth);

                    if(dg5.Rows.Count == 0) 
                    {
                       nodata += cdcreportno;
                       nodata += newline;
                    }


                    dataGridView3.ColumnHeadersHeight = 20;
                    dataGridView3.Columns[0].Width = 120;
                    dataGridView3.Columns[1].Width = 400;
                    dataGridView3.Columns[0].Name = "疾病代碼";
                    dataGridView3.Columns[1].Name = "疾病名稱";
                    dataGridView3.Columns[2].Name = " ";
                    dataGridView3.Columns[0].DataPropertyName = "疾病代碼";
                    dataGridView3.Columns[1].DataPropertyName = "疾病名稱";
                    dataGridView3.Columns[2].DataPropertyName = " ";

                    dataGridView4.Columns[0].Width = 100;
                    dataGridView4.Columns[1].Width = 700;
                    dataGridView4.Columns[2].Width = 100;
                    dataGridView4.Columns[3].Width = 100;
                    dataGridView4.Columns[4].Width = 100;
                    dataGridView4.Columns[5].Width = 100;
                    dataGridView4.Columns[6].Width = 100;
                    dataGridView4.Columns[7].Width = 100;
                    dataGridView4.Columns[8].Width = 100;
                    dataGridView4.Columns[9].Width = 100;
                    dataGridView4.Columns[10].Width = 100;
                    dataGridView4.Columns[11].Width = 100;
                    dataGridView4.Columns[12].Width = 100;
                    dataGridView4.Columns[13].Width = 100;
                    dataGridView4.Columns[14].Width = 100;
                    dataGridView4.Columns[0].Name = "批價碼";
                    dataGridView4.Columns[1].Name = "名稱";
                    dataGridView4.Columns[2].Name = "用量";
                    dataGridView4.Columns[3].Name = "單位";
                    dataGridView4.Columns[4].Name = "頻次";
                    dataGridView4.Columns[5].Name = "途徑";
                    dataGridView4.Columns[6].Name = "天數";
                    dataGridView4.Columns[7].Name = "總量";
                    dataGridView4.Columns[8].Name = "外加天數";
                    dataGridView4.Columns[9].Name = "外加總量";
                    dataGridView4.Columns[10].Name = "計價方式";
                    dataGridView4.Columns[11].Name = "急件";
                    dataGridView4.Columns[12].Name = "慢箋";
                    dataGridView4.Columns[13].Name = "交付";
                    dataGridView4.Columns[14].Name = " ";   
                    dataGridView4.Columns[0].DataPropertyName = "批價碼";
                    dataGridView4.Columns[1].DataPropertyName = "名稱";
                    dataGridView4.Columns[2].DataPropertyName = "用量";
                    dataGridView4.Columns[3].DataPropertyName = "單位";
                    dataGridView4.Columns[4].DataPropertyName = "頻次";
                    dataGridView4.Columns[5].DataPropertyName = "途徑";
                    dataGridView4.Columns[6].DataPropertyName = "天數";
                    dataGridView4.Columns[7].DataPropertyName = "總量";
                    dataGridView4.Columns[8].DataPropertyName = "外加天數";
                    dataGridView4.Columns[9].DataPropertyName = "外加總量";
                    dataGridView4.Columns[10].DataPropertyName = "計價方式";
                    dataGridView4.Columns[11].DataPropertyName = "急件";
                    dataGridView4.Columns[12].DataPropertyName = "慢箋";
                    dataGridView4.Columns[13].DataPropertyName = "交付";
                    dataGridView4.Columns[14].DataPropertyName = " ";

                    using (Bitmap bmp = new Bitmap(this.Width, this.Height))
                    {
                        this.DrawToBitmap(bmp, new Rectangle(Point.Empty, bmp.Size));
                        string bmpsave = string.Format(@"C:\Users\3732\Desktop\CDC\{0}.png", cdcreportno);
                        
                        bmp.Save(bmpsave, ImageFormat.Png);
                    }

                    if (Directory.Exists(string.Format(@"C:\Users\3732\Desktop\CDC\{0}", documentdate)))
                    {
                        //資料夾存在
                    }
                    else
                    {
                        //新增資料夾
                        Directory.CreateDirectory(string.Format(@"C:\Users\3732\Desktop\CDC\{0}", documentdate));
                    }



                    

                    try
                    {
                        //Create a PdfDocument instance			
                        PdfDocument pdf = new PdfDocument();

                        //pdf.LoadFromFile
                        PdfPageBase page1 = pdf.Pages.Add();
                        PdfPageBase page2 = pdf.Pages.Add();

                        //Load an image
                        string bmppdfsave = string.Format(@"C:\Users\3732\Desktop\CDC\{0}.png", cdcreportno);
                        PdfImage image = PdfImage.FromFile(bmppdfsave);

                        //Specify the width and height of the image area on the page			
                        float width = image.Width * 0.250f;
                        float height = image.Height * 0.250f;

                        //Specify the X and Y coordinates to start drawing the image			
                        float x = 1f;
                        float y = 1f;

                        //Draw the image at a specified location on the page			
                        page2.Canvas.DrawImage(image, x, y, width, height);
                        pdf.Pages.Remove(page1);

                        //Save the result document
                        //string order = "_醫囑單";
                        string bmppdf = string.Format(@"C:\Users\3732\Desktop\CDC\{0}\{1}_醫囑單.pdf", documentdate, cdcreportno); ;
                        pdf.SaveToFile(bmppdf, FileFormat.PDF);
                        ;
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                
                    finally
                        {

                        dataGridView3.DataSource = null;
                        dataGridView3.Rows.Clear();
                        dataGridView4.DataSource = null;
                        dataGridView4.Rows.Clear();
                        dataGridView5.DataSource = null;
                        dataGridView5.Rows.Clear();
                    }
                }
            }
            catch(Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                ;
            }
            MessageBox.Show(nodata, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }


        public string Querydata1(string test, string year, string month)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
                        @" SELECT Concat(Concat(Concat(Ltrim(Substr(Lpad(c.clinic_date, 7, '0'), 1, 3), '0'), '/'),Concat(Ltrim(Substr(Lpad(c.clinic_date, 7, '0'), 4, 2), '0'), '/')), Ltrim(Substr(Lpad(c.clinic_date, 7, '0'), 6, 2), '0')) AS ""看診日期""," +
                        @"        CONCAT(TRIM(c.doctor_name),'醫師')  AS ""醫師"",                                                                                                                                                                             " +
                        @"        e.div_short_name AS ""科別""                                                                                                                                                                                                 " +
                        @" FROM   cdreport.ptcdreport C                                                                                                                                                                                                        " +
                        @"        LEFT JOIN ohis{1}{2}.ptopd A                                                                                                                                                                                                     " +
                        @"               ON c.clinic_date = a.clinic_date                                                                                                                                                                                      " +
                        @"                  AND c.duplicate_no = a.duplicate_no                                                                                                                                                                                " +
                        @"                  AND c.chart_no = a.chart_no                                                                                                                                                                                        " +
                        @"        LEFT JOIN onh{1}{2}.ptopd D                                                                                                                                                                                                      " +
                        @"               ON D.clinic_date = a.clinic_date                                                                                                                                                                                      " +
                        @"                  AND D.duplicate_no = a.duplicate_no                                                                                                                                                                                " +
                        @"                  AND D.chart_no = a.chart_no                                                                                                                                                                                        " +
                        @"        LEFT JOIN mast.doctor B                                                                                                                                                                                                      " +
                        @"               ON a.doctor_no = b.doctor_no                                                                                                                                                                                          " +
                        @"        LEFT JOIN mast.div E                                                                                                                                                                                                         " +
                        @"               ON A.div_no = E.div_no                                                                                                                                                                                                " +
                        @" WHERE  c.cdc_report_no = '{0}'                                                                                                                                                                                                      ";
                        
            string sql = string.Format(tempsql, test, year, month);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("textlabel4 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }

            string clinicdate1 = "";
            string doctor1 = "";
            string div1 = "";
            string textlabel4;

            foreach (DataRow dr in rt.Rows)
            {
                clinicdate1 = dr["看診日期"].ToString();
                doctor1     = dr["醫師"].ToString();
                div1        = dr["科別"].ToString();
            }
            textlabel4 = clinicdate1 + ",  " + doctor1 + ",  " + div1;
            return textlabel4;
        }

        public string Querydata2(string test)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
                            @" SELECT Trim(C.id_no) AS ""身分證字號""     " +
                            @" FROM   cdreport.ptcdreport C              " +
                            @" WHERE  c.cdc_report_no = '{0}'            " ;
            string sql = string.Format(tempsql, test);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("richTextBox5 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch(Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }

            string id = "";
            foreach (DataRow dr in rt.Rows)
            {
                id = dr["身分證字號"].ToString();
            }
            return id;
        }

        public string Querydata3(string test, string yearlabel, string month, string year)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
            @" SELECT                                         " +
            @" CONCAT(CASE A.sub_type                         " +
            @"        WHEN 'CC' THEN 'Chief Complaints:'      " +
            @"         WHEN 'PI' THEN 'Present Illness:'      " +
            @"         WHEN 'AL' THEN 'Allergy Hx:'           " +
            @"         WHEN 'PS' THEN 'Past Hx:'              " +
            @"         WHEN 'OP' THEN 'Operation Hx:'         " +
            @"         WHEN ' ' THEN 'Others:'                " +
            @"         WHEN 'VS' THEN 'Vital Signs:'          " +
            @"         WHEN 'EA' THEN 'Ext. Appearance:'      " +
            @"         WHEN 'HE' THEN 'Head Eye ENT:'         " +
            @"         WHEN 'NE' THEN 'Neck:'                 " +
            @"         WHEN 'CL' THEN 'Chest And Lung:'       " +
            @"         WHEN 'HR' THEN 'Heart:'                " +
            @"         WHEN 'AB' THEN 'Abdomen:'              " +
            @"         WHEN 'EX' THEN 'Extremities:'          " +
            @"         WHEN 'BS' THEN 'Back And Spine:'       " +
            @"         WHEN 'NX' THEN 'Neuro. Exam.:'         " +
            @"       END,                                     " +
            @"       A.content) AS ""S""                      " +
            @" FROM   cdreport.ptcdreport C                   " +
            @"       LEFT JOIN onh{1}{2}.ordsoopd A           " +
            @"              ON C.chart_no = A.chart_no        " +
            @"             AND C.clinic_date = A.clinic_date  " +
            @"             AND C.duplicate_no = A.duplicate_no" +
            @"       LEFT JOIN mast.doctor B                  " +
            @"              ON C.vs_no = B.doctor_no          " +
            @"       LEFT JOIN mast.div D                     " +
            @"              ON C.div_no = D.div_no            " +
            @" WHERE  C.clinic_date LIKE '{3}{2}%'            " +
            @"       AND c.cdc_report_no = '{0}'              " +
            @"       AND A.SO_TYPE = 'S'                      " +
            @"       AND A.CONTENT <> ' '                     " ;

            string sql = string.Format(tempsql, test, yearlabel, month, year);

            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("richTextBox6 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }

            string s = "";
            foreach (DataRow dr in rt.Rows)
            {
                s += dr["S"].ToString() + newline;
            }
            return s;
        }

        public string Querydata4(string test, string yearlabel, string month, string year)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
            @" SELECT                                            " +
            @" CONCAT(                                           " +
            @" CASE A.sub_type                                   " +
            @"          WHEN 'CC' THEN 'Chief Complaints:'       " +
            @"          WHEN 'PI' THEN 'Present Illness:'        " +
            @"          WHEN 'AL' THEN 'Allergy Hx:'             " +
            @"          WHEN 'PS' THEN 'Past Hx:'                " +
            @"          WHEN 'OP' THEN 'Operation Hx:'           " +
            @"          WHEN ' '  THEN 'Others:'                 " +
            @"          WHEN 'VS' THEN 'Vital Signs:'            " +
            @"          WHEN 'EA' THEN 'Ext. Appearance:'        " +
            @"          WHEN 'HE' THEN 'Head Eye ENT:'           " +
            @"          WHEN 'NE' THEN 'Neck:'                   " +
            @"          WHEN 'CL' THEN 'Chest And Lung:'         " +
            @"          WHEN 'HR' THEN 'Heart:'                  " +
            @"          WHEN 'AB' THEN 'Abdomen:'                " +
            @"          WHEN 'EX' THEN 'Extremities:'            " +
            @"          WHEN 'BS' THEN 'Back And Spine:'         " +
            @"          WHEN 'NX' THEN 'Neuro. Exam:'            " +
            @"        END,                                       " +
            @"        A.content) AS ""O""                        " +
            @" FROM   cdreport.ptcdreport C                      " +
            @"        LEFT JOIN onh{1}{2}.ordsoopd A             " +
            @"               ON C.chart_no = A.chart_no          " +
            @"               AND C.clinic_date = A.clinic_date   " +
            @"               AND C.duplicate_no = A.duplicate_no " +
            @"        LEFT JOIN mast.doctor B                    " +
            @"               ON C.vs_no = B.doctor_no            " +
            @"        LEFT JOIN mast.div D                       " +
            @"               ON C.div_no = D.div_no              " +
            @"                                                   " +
            @" WHERE  C.clinic_date LIKE '{3}{2}%'               " +
            @"        AND c.cdc_report_no = '{0}'                " +
            @"        AND A.SO_TYPE = 'O'                        " +
            @"        AND A.CONTENT <> ' '                       " +
            @"                                                   " ;


            string sql = string.Format(tempsql, test, yearlabel, month, year);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("richTextBox7 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }

            string s = "";
            foreach (DataRow dr in rt.Rows)
            {
                s += dr["O"].ToString() + newline;
            }
            return s;
        }

        public DataTable Querydata5(string test, string yearlabel, string month, string year)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
            @"SELECT A.disease_code AS ""疾病代碼"",                       " +
            @"       A.disease_name AS ""疾病名稱"",                        " +
            @"       ''             AS "" ""                               " +
            @"FROM   cdreport.ptcdreport C                                 " +
            @"       LEFT JOIN onh{1}{2}.ORDAOPD10 A                       " +
            @"              ON C.clinic_date = A.clinic_date               " +
            @"                 AND C.chart_no = A.chart_no                 " +
            @"                 AND C.duplicate_no = A.duplicate_no         " +
            @"WHERE  C.clinic_date LIKE '{3}{2}%'                          " +
            @"       AND c.cdc_report_no = '{0}'                           " ;

            string sql = string.Format(tempsql, test, yearlabel, month, year);

            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("dataGridView3 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }
            return rt;
        }


        public DataTable Querydata6(string test, string yearlabel, string month, string year)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
            //6秒版
            /*
            @"SELECT A.code                                             AS ""批價碼"",     " +
            @"       B.english_name                                     AS ""名稱"",       " +
            @"       A.dose_qty                                         AS ""用量"",       " +
            @"       A.dose_unit                                        AS ""單位"",       " +
            @"       A.frequency_code                                   AS ""頻次"",       " +
            @"       A.method_code                                      AS ""途徑"",       " +
            @"       A.days                                             AS ""天數"",       " +
            @"       ROUND(A.total_qty)                                 AS ""總量"",       " +
            @"       ''                                                 AS ""外加天數"",   " +
            @"       ''                                                 AS ""外加總量"",   " +
            @"       CONCAT(concat(a.pricing_flag,'-'),                                     " +
            @"       CASE a.pricing_flag                                                    " +
            @"        WHEN 'Y' THEN                                                         " +
            @"         '自費計價(由電腦依計價原則設定)'                                       " +
            @"        WHEN 'N' THEN                                                         " +
            @"         '健保申報(由電腦依計價原則設定)'                                       " +
            @"        WHEN 'H' THEN                                                         " +
            @"         '健保申報不計價(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'h' THEN                                                         " +
            @"         '健保申報不計價(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'S' THEN                                                         " +
            @"         '一律自費'                                                           " +
            @"        WHEN 's' THEN                                                         " +
            @"         '一律自費'                                                           " +
            @"        WHEN 'X' THEN                                                         " +
            @"         '不計價且不申報(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'x' THEN                                                         " +
            @"         '不計價且不申報(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'Z' THEN                                                         " +
            @"         '自費病人自費,健保病人不申報不計價'                                     " +
            @"        WHEN 'V' THEN                                                         " +
            @"         '虛醫令,交付調劑之藥品空針'                                            " +
            @"        END)                                              AS ""計價方式"",     " +
            @"       A.emg_flag                                         AS ""急件"",         " +
            @"       CASE A.remark                                                          " +
            @"         WHEN 'C' THEN 'Y'                                                    " +
            @"         WHEN 'D' THEN 'Y'                                                    " +
            @"         ELSE ''                                                              " +
            @"       END                                                AS ""慢箋"",        " +
            @"       CASE A.remark                                                          " +
            @"         WHEN 'V' THEN 'Y'                                                    " +
            @"         WHEN 'D' THEN 'Y'                                                    " +
            @"         ELSE ''                                                              " +
            @"       END                                                AS ""交付"",        " +
            @"       ' '                                                AS ""final""       " +
            @"FROM   cdreport.ptcdreport C                                                  " +
            @"       LEFT JOIN ohis11.acntopd A                                             " +
            @"              ON c.clinic_date = a.clinic_date                                " +
            @"                 AND c.duplicate_no = a.duplicate_no                          " +
            @"                 AND c.chart_no = a.chart_no                                  " +
            @"       LEFT JOIN mast.price b                                                 " +
            @"              ON A.code = B.code                                              " +
            @"WHERE  a.clinic_date LIKE '11111%'                                            " +
            @"       AND C.cdc_report_no = '1113703030457'                                  " +
            @"       AND B.effective_date = (SELECT Max(B.effective_date)                   " +
            @"                             FROM   mast.price b                              " +
            @"                             WHERE  A.code = B.code)                          " ;
            */
            
            //0.6秒版            
            @"SELECT A.code                                             AS ""批價碼"",      " +
            @"       C.english_name                                     AS ""名稱"",        " +
            @"       A.dose_qty                                         AS ""用量"",        " +
            @"       A.dose_unit                                        AS ""單位"",        " +
            @"       A.frequency_code                                   AS ""頻次"",        " +
            @"       A.method_code                                      AS ""途徑"",        " +
            @"       A.days                                             AS ""天數"",        " +
            @"       ROUND(A.total_qty)                                 AS ""總量"",        " +
            @"       ''                                                 AS ""外加天數"",     " +
            @"       ''                                                 AS ""外加總量"",     " +
            @"       CONCAT(concat(a.pricing_flag,'-'),                                     " +
            @"       CASE a.pricing_flag                                                    " +
            @"        WHEN 'Y' THEN                                                         " +
            @"         '自費計價(由電腦依計價原則設定)'                                       " +
            @"        WHEN 'N' THEN                                                         " +
            @"         '健保申報(由電腦依計價原則設定)'                                       " +
            @"        WHEN 'H' THEN                                                         " +
            @"         '健保申報不計價(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'h' THEN                                                         " +
            @"         '健保申報不計價(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'S' THEN                                                         " +
            @"         '一律自費'                                                           " +
            @"        WHEN 's' THEN                                                         " +
            @"         '一律自費'                                                           " +
            @"        WHEN 'X' THEN                                                         " +
            @"         '不計價且不申報(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'x' THEN                                                         " +
            @"         '不計價且不申報(由電腦依計價原則設定,不開放選擇)'                        " +
            @"        WHEN 'Z' THEN                                                         " +
            @"         '自費病人自費,健保病人不申報不計價'                                     " +
            @"        WHEN 'V' THEN                                                         " +
            @"         '虛醫令,交付調劑之藥品空針'                                            " +
            @"        END)                                              AS ""計價方式"",     " +
            @"       A.emg_flag                                         AS ""急件"",         " +
            @"       CASE A.remark                                                          " +
            @"         WHEN 'C' THEN 'Y'                                                    " +
            @"         WHEN 'D' THEN 'Y'                                                    " +
            @"         ELSE ''                                                              " +
            @"       END                                                AS ""慢箋"",        " +
            @"       CASE A.remark                                                          " +
            @"         WHEN 'V' THEN 'Y'                                                    " +
            @"         WHEN 'D' THEN 'Y'                                                    " +
            @"         ELSE ''                                                              " +
            @"       END                                                AS ""交付"",        " +
            @"       ''                                                AS "" ""              " +
            @"FROM   cdreport.ptcdreport C                                                  " +
            @"       LEFT JOIN ohis{1}{2}.acntopd A                                         " +
            @"              ON c.clinic_date = a.clinic_date                                " +
            @"                 AND c.duplicate_no = a.duplicate_no                          " +
            @"                 AND c.chart_no = a.chart_no                                  " +
            @"LEFT JOIN (SELECT A.code,Max(A.effective_date) AS effective_date              " +
            @"           FROM   mast.price A                                                " +
            @"           GROUP  BY A.code)B                                                 " +
            @"  ON A.code = B.code                                                          " +
            @"LEFT JOIN (SELECT A.CODE,A.ENGLISH_NAME,A.effective_date                      " +
            @"           FROM   mast.price A)C                                              " +
            @"  ON B.code = C.code                                                          " +
            @"  AND B.effective_date = C.effective_date                                     " +
            @"WHERE  a.clinic_date LIKE '{3}{2}%'                                           " +
            @"       AND c.cdc_report_no = '{0}'                                            " +
            @"       AND B.effective_date = (SELECT Max(B.effective_date)                   " +
            @"                             FROM   mast.price b                              " +
            @"                             WHERE  A.code = B.code)                          " ;

            string sql = string.Format(tempsql, test, yearlabel, month, year);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("dataGridView4 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }
            return rt;
        }


        public DataTable Querydata7(string test, string yearlabel, string month)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
            @" SELECT                                                                                                                                                                                                                   "+        
            @"       Concat(Concat(Concat(LTRIM(SUBSTR(LPAD(c.clinic_date, 7, '0'), 1, 3), '0'), '/'), Concat(LTRIM(SUBSTR(LPAD(c.clinic_date, 7, '0'), 4, 2), '0'), '/')), LTRIM(SUBSTR(LPAD(c.clinic_date, 7, '0'), 6, 2), '0')) AS ""看診日期"",  "+
            @"       c.chart_no                                                                                                                                                                                         AS ""病歷號碼"",  " +
            @"       c.duplicate_no                                                                                                                                                                                     AS ""重複序號"",  " +
            @"       d.nh_clinic_seq                                                                                                                                                                                    AS ""卡號"",      " +
            @"       c.pt_name                                                                                                                                                                                          AS ""姓名"",      " +
            @"       CASE c.sex                                                                                                                                                                                                           " +
            @"         WHEN 'F'                                                                                                                                                                                                           " +
            @"           THEN '女'                                                                                                                                                                                                        " +
            @"         WHEN 'M'                                                                                                                                                                                                           " +
            @"           THEN '男'                                                                                                                                                                                                        " +
            @"         END                                                                                                                                                                                              AS ""性別"",      " +
            @"       Concat(Concat(Concat(LTRIM(SUBSTR(LPAD(c.birth_date, 7, '0'),1,3),'0'), '/'),Concat(LTRIM(SUBSTR(LPAD(c.birth_date, 7, '0'),4,2),'0'), '/')),LTRIM(SUBSTR(LPAD(c.birth_date, 7, '0'),6,2),'0'))    AS ""生日"",      " +
            @"       e.div_short_name                                                                                                                                                                                   AS ""科別"",      " +
            @"       c.doctor_name                                                                                                                                                                                      AS ""醫師""       " +
            @" FROM   cdreport.ptcdreport C                                                                                                                                                                                               " +
            @"       LEFT JOIN ohis{1}{2}.ptopd A                                                                                                                                                                                             " +
            @"              ON c.clinic_date = a.clinic_date                                                                                                                                                                              " +
            @"                 AND c.duplicate_no = a.duplicate_no                                                                                                                                                                        " +
            @"                 AND c.chart_no = a.chart_no                                                                                                                                                                                " +
            @"       LEFT JOIN onh{1}{2}.ptopd D                                                                                                                                                                                              " +
            @"              ON D.clinic_date = a.clinic_date                                                                                                                                                                              " +
            @"                 AND D.duplicate_no = a.duplicate_no                                                                                                                                                                        " +
            @"                 AND D.chart_no = a.chart_no                                                                                                                                                                                " +
            @"       LEFT JOIN mast.doctor B                                                                                                                                                                                              " +
            @"              ON a.doctor_no = b.doctor_no                                                                                                                                                                                  " +
            @"       LEFT JOIN mast.div E                                                                                                                                                                                                 " +
            @"              ON A.Div_No = E.div_no                                                                                                                                                                                        " +
            @" WHERE  c.cdc_report_no = '{0}'                                                                                                                                                                                              ";

            string sql = string.Format(tempsql, test, yearlabel, month);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("dataGridView5 "+test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }
            return rt;

        }

        public string Querydata8(string test, string yearlabel, string month)
        {
            //CDC Report No Ver:    c.cdc_report_no = '{0}'
            //ID No Ver        :    c.id_no = '{0}'
            string tempsql =
                        @" SELECT trim(c.report_date) AS ""資料夾日期""            " +                   
                        @" FROM   cdreport.ptcdreport C                           " +
                        @"        LEFT JOIN ohis{1}{2}.ptopd A                    " +
                        @"               ON c.clinic_date = a.clinic_date         " +
                        @"                  AND c.duplicate_no = a.duplicate_no   " +
                        @"                  AND c.chart_no = a.chart_no           " +
                        @"        LEFT JOIN onh{1}{2}.ptopd D                     " +
                        @"               ON D.clinic_date = a.clinic_date         " +
                        @"                  AND D.duplicate_no = a.duplicate_no   " +
                        @"                  AND D.chart_no = a.chart_no           " +
                        @"        LEFT JOIN mast.doctor B                         " +
                        @"               ON a.doctor_no = b.doctor_no             " +
                        @"        LEFT JOIN mast.div E                            " +
                        @"               ON A.div_no = E.div_no                   " +
                        @" WHERE  c.cdc_report_no = '{0}'                         " ;

            string sql = string.Format(tempsql, test, yearlabel, month);
            Oledb oledb = new Oledb();
            OleDbConnection odcnn = oledb.getoledbconnection("ORACLE_DB_HO", "mast");
            OleDbCommand odcmm = new OleDbCommand();
            DataTable rt = new DataTable();
            odcmm.Connection = odcnn;
            odcmm.CommandType = CommandType.Text;
            odcmm.CommandText = sql;
            OleDbDataAdapter oddad;

            try
            {
                odcnn.Open();
                oddad = new OleDbDataAdapter(odcmm);
                oddad.Fill(rt);

                if (rt.Rows.Count == 0)
                {
                    MessageBox.Show("textlabel4 " + test, "查無資料", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                odcnn.Close();
            }

            catch (Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                odcnn.Dispose();
                odcmm.Dispose();
            }
                string documentdate="";
            try
            {
                
                documentdate = rt.Rows[0].Field<string>(0);
            }
            catch(Exception ex)
            {
                string error =
                                "{Message}" + ex.Message + newline +
                                "{Stacktrace}" + ex.StackTrace + newline +
                                "{Targetsite}" + ex.TargetSite + newline +
                                "{Tostring}" + ex.ToString();

                MessageBox.Show(error, "提醒", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                ;
            }
            
            return documentdate;
        }
    }
}
