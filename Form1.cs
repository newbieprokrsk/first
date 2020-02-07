using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace Compare_price_markets
{
    public partial class Form1 : Form
    {
        public Excel.Application excelapp;
        public Excel.Workbook excelappworkbook;     //Ссылка на созданный объект(1 книга)
        public Excel.Worksheet excelworksheet;      //рабочий лист в данный момент

        public Excel.Application excelapp2;
        public Excel.Workbook excelappworkbook2;
        public Excel.Worksheet excelworksheet2;

        public string[] decomposition_of_Pitstop(string full_name)
        {
            string size = "", brand = "", model = "", ship = "", kom = "", index = "", name = "", brand_model = "", kamera = "";
            int i = 0;
            string[] brands = new string[] {"cordiant","toyo","nokian","nitto","yokohama","kumho","bridgestone","hankook","viatti","tunga","firestone",
            "белшина","нкшз","amtel","doublestar","michelin","continental","crossleader","marshal","барнаул", "кшз","волшз","formula","dunlop"};
            full_name = full_name.ToLower();

            if (full_name.Contains(", шт"))
            {
                full_name = full_name.Replace(", шт", "");
            }

            while (full_name[i] != ' ')
            {
                size = size + full_name[i];
                i++;
            }
            //name = full_name.Substring(0, size.Length);
            name = full_name;
            name = name.Replace(size + " ", "");
            //обрезаем шип 4шт и л в разных вариациях
            if (name.Contains("шип.") || name.Contains("шип"))
            {
                if (name.Contains("шип."))
                {
                    ship = "шип.";
                    name = name.Replace(" шип.", "");
                }
                if (name.Contains("шип"))
                {
                    ship = "шип";
                    name = name.Replace(" шип", "");
                }

            }
            else
                ship = "";
            if (name.Contains("з. "))
                name = name.Replace("з.", " ");
            if (name.Contains("л."))
                name = name.Replace("л.", "");
            if (name.Contains(" л "))
                name = name.Replace(" л ", " ");
            if (name.Contains("(4шт)"))
                kom = "(4шт)";
            name = name.Replace(" (4шт)", "");
            if (name.Contains("(2шт)"))
                kom = "(2шт)";
            name = name.Replace(" (2шт)", "");
            if (name.Contains(" в "))
                name = name.Replace(" в ", " ");

            foreach (string element in brands)
                if (name.Contains(element))
                {
                    brand = element;
                    break;
                }
            name = name.Replace(brand + " ", "");


            //Business CA -1 112R
            //достаем индекс
            i = name.Length - 1;
            while (name[i] != ' ')
            {
                index = index + name[i];
                i--;
            }
            //развернем строку
            StringBuilder sb = new StringBuilder(index.Length);
            for (i = index.Length; i-- != 0;)
                sb.Append(index[i]);
            index = sb.ToString();
            //удалим оставшейся индекс в конце
            name = name.Replace(" " + index, "");
            model = name;

            // size = size.Replace("/", " ");
            //size = size.Replace("C", "");
            //size = size.Replace("С", "");
            model = model.Replace("-", " ");
            
            return new string[] { size, brand, model, kom, ship, index };
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button_find_Olta_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Olta.Text = dlg.FileName;
            }
        }
        private void button_find_Shintorg_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Shintorg.Text = dlg.FileName;
            }
        }
        private void button_find_Skot_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Skot.Text = dlg.FileName;
            }
        }
        private void button_find_Nortek_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Nortek.Text = dlg.FileName;
            }
        }
        private void button_find_Pitstop_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                textBox_Pitstop.Text = dlg.FileName;
            }
        }
        private void textBox_Olta_TextChanged(object sender, EventArgs e)
        {

        }
        //-----------------------------------------------------------------------------------------------------------------------------------------
        public string[] Fix_Olta (string[] for_fix,string kind)
        {
            string current_name_Olta = for_fix[2];
            string Pitstop_brand = for_fix[0]; //parts_of_decomposition_Pitstop[2]
            string Pitstop_moodel = for_fix[1]; //parts_of_decomposition_Pitstop[3]


            if (current_name_Olta.Contains("кама") && Pitstop_brand.Contains("кама"))
            {
                Pitstop_moodel = Pitstop_moodel.Replace("нкшз", "");
                current_name_Olta = current_name_Olta.Replace("кама euro", "кама евро");
            }
            if (current_name_Olta.Contains("formula"))
            {
                if (Pitstop_brand.Contains(" xl"))
                    Pitstop_brand = Pitstop_brand.Replace(" xl", "");
                if (current_name_Olta.Contains(" xl"))
                    current_name_Olta = current_name_Olta.Replace(" xl", "");
            }
            if (current_name_Olta.Contains("nokian"))
            {
                if (Pitstop_moodel.Contains("xl"))
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                if (current_name_Olta.Contains(" xl"))
                    current_name_Olta = current_name_Olta.Replace(" xl", "");
            }
            if (current_name_Olta.Contains("toyo"))
            {
                current_name_Olta = current_name_Olta.Replace(" xl", "");
                if (current_name_Olta.Contains("observe garit giz"))
                    current_name_Olta = current_name_Olta.Replace("observe garit giz", "obgiz");
                if (current_name_Olta.Contains("observe "))
                    current_name_Olta = current_name_Olta.Replace("observe ", "ob");
                /*if (current_name_Olta.Contains("obifa"))
                    current_name_Olta = current_name_Olta.Replace("obifa", "obg3sa");*/
                if (current_name_Olta.Contains("g3 ice"))
                    current_name_Olta = current_name_Olta.Replace("g3 ice", "g3s");
            }
            if (current_name_Olta.Contains("yokohama"))
            {
                if (current_name_Olta.Contains("ig"))
                    current_name_Olta = current_name_Olta.Replace("ig", "ig ");
                else
                    current_name_Olta = current_name_Olta.Replace("g0", "g 0");
            }
            if (current_name_Olta.Contains("bridgestone"))
            {
                current_name_Olta = current_name_Olta.Replace(" xl", "");
                Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                if (current_name_Olta.Contains("blizzak "))
                {
                    current_name_Olta = current_name_Olta.Replace("blizzak ", "");
                }
                if (current_name_Olta.Contains("ice cruiser 7000"))
                    current_name_Olta = current_name_Olta.Replace("ice cruiser 7000", "ic7000");
                if (current_name_Olta.Contains("dm v2"))
                {
                    current_name_Olta = current_name_Olta.Replace("dm v2", "dmv 2");
                }
            }
            if (current_name_Olta.Contains("tunga"))
            {
                if (Pitstop_moodel == "nordway 2")
                {
                    if (!current_name_Olta.Contains("nordway 2"))
                    {
                        current_name_Olta = current_name_Olta.Replace("nordway", "");
                    }
                }
                else if (Pitstop_moodel == "nordway")
                {
                    if (current_name_Olta.Contains("nordway 2"))
                        current_name_Olta = current_name_Olta.Replace("nordway 2", "");
                }
            }
            if (current_name_Olta.Contains("cordiant"))
            {
                current_name_Olta = current_name_Olta.Replace(" pw 2", "");



                if (Pitstop_moodel == "snow cross 2 suv")
                {
                    if (!current_name_Olta.Contains("snow cross 2 suv"))
                    {
                        current_name_Olta = current_name_Olta.Replace("snow cross", "");
                    }
                }
                else if (Pitstop_moodel == "snow cross 2")
                {
                    if (current_name_Olta.Contains("snow cross 2 suv"))
                        current_name_Olta = current_name_Olta.Replace("snow cross 2 suv", "");
                    else
                        if (!current_name_Olta.Contains("snow cross 2"))
                        current_name_Olta = current_name_Olta.Replace("snow cross", "");
                }
                else if (Pitstop_moodel == "snow cross")
                {
                    if (current_name_Olta.Contains("snow cross 2"))
                        current_name_Olta = current_name_Olta.Replace("snow cross", "");
                }

            }
            if (current_name_Olta.Contains("nitto"))
            {
                if (current_name_Olta.Contains("winter "))
                    current_name_Olta = current_name_Olta.Replace("winter ", "nt");
            }
            if (current_name_Olta.Contains("hankook"))
            {
                if (Pitstop_moodel.Contains("stud"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace(" stud", "");
                }

                if (Pitstop_moodel.Contains("xl"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
            }

                return for_fix;
        }

        public string[] Fix_Nortek(string[] for_fix, string kind)
        {
            string current_name_Nortek = for_fix[2];
            string Pitstop_brand = for_fix[0]; //parts_of_decomposition_Pitstop[2]
            string Pitstop_moodel = for_fix[1]; //parts_of_decomposition_Pitstop[3]

            if (current_name_Nortek.Contains("viatti") && Pitstop_brand.Contains("viatti"))
            {
                if (current_name_Nortek.Contains("v 521") && Pitstop_moodel.Contains("v 521"))
                {
                    current_name_Nortek = current_name_Nortek.Replace(" nordico", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" nordico", "");
                }
            }
            if (current_name_Nortek.Contains("nokian"))
            {
                if (current_name_Nortek.Contains("rs2") && Pitstop_moodel.Contains("rs2 xl"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
                if (current_name_Nortek.Contains("hakkapeliitta 8") && Pitstop_moodel.Contains("hkpl 8"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("hakkapeliitta 8", "hkpl 8");
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
                if (current_name_Nortek.Contains("hakkapeliitta 9") && Pitstop_moodel.Contains("hkpl 9"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("hakkapeliitta 9", "hkpl 9");
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
                if (current_name_Nortek.Contains("nordman 7") && Pitstop_moodel.Contains("nordman 7"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
                if (current_name_Nortek.Contains("rs2") && Pitstop_moodel.Contains("rs2"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace("   xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                }
                if (current_name_Nortek.Contains("nordman 5") && Pitstop_moodel.Contains("nordman 5"))
                {
                    Pitstop_moodel = Pitstop_moodel.Replace("  xl", "");
                    Pitstop_moodel = Pitstop_moodel.Replace(" xl", "");
                }
            }

            if (current_name_Nortek.Contains("кама"))
            {
                if (current_name_Nortek.Contains("euro") && current_name_Nortek.Contains("519"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("кама", "кама евро 519");
                }
            }
            if (current_name_Nortek.Contains("nitto"))
            {
                if (current_name_Nortek.Contains("therma spike"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("therma spike", "ntspk");
                }
            }
            if (current_name_Nortek.Contains("bridgestone"))
            {
                if (current_name_Nortek.Contains("ice cruiser"))
                {

                    current_name_Nortek = current_name_Nortek.Replace("ice cruiser ", "ic");
                }
                if (current_name_Nortek.Contains("spike 02") && current_name_Nortek.Contains("suv") && current_name_Nortek.Contains("xl"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("suv", "");
                    current_name_Nortek = current_name_Nortek.Replace("xl", "");
                    current_name_Nortek = current_name_Nortek.Replace("spike 02", "spike 02 suv xl");
                }
            }
            if (current_name_Nortek.Contains("cordiant"))
            {
                if (current_name_Nortek.Contains("snow cross 2") && !Pitstop_moodel.Contains("snow cross 2"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("snow cross 2", "");
                }
            }
            if (current_name_Nortek.Contains("toyo"))
            {
                if (current_name_Nortek.Contains("observe garit giz"))
                {
                    current_name_Nortek = current_name_Nortek.Replace("observe garit giz", "obgiz");
                }
            }

            return for_fix;
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------
        public bool Compare_Pitstop_Shintorg(string[] Pitstop, string current_name_Shintorg)
        {

            return true;
        }

        public bool Compare_Pitstop_Nortek(string[] Pitstop, string current_name_Nortek)
        {
            current_name_Nortek = current_name_Nortek.Replace("_", " ");
            current_name_Nortek = current_name_Nortek.Replace("-", " ");
            string razmer = Pitstop[0];
            int size = razmer.Length - 1;
            StringBuilder someString = new StringBuilder(razmer);
            while (razmer[size] != '/')
            {
                size--;
            }
            someString[size] = ' ';
            razmer = someString.ToString();
            razmer = razmer.Replace(" ", " r");

            if (!current_name_Nortek.Contains(razmer))
                return false;

            string[] for_fix = new string[] { Pitstop[1], Pitstop[2], current_name_Nortek };
            for_fix = Fix_Olta(for_fix, "brand");
            current_name_Nortek = for_fix[2];
            Pitstop[1] = for_fix[0];
            Pitstop[2] = for_fix[1];

            if (!current_name_Nortek.Contains(Pitstop[1]))
                return false;

            if (!current_name_Nortek.Contains(Pitstop[2]))
                return false;
            else return true;

        }

        public bool Compare_Pitstop_Olta(string[] Pitstop, string current_name_Olta )
        {
            current_name_Olta = current_name_Olta.Replace("_", " ");
            current_name_Olta = current_name_Olta.Replace("-", " ");
            string razmer = Pitstop[0];
            int size = razmer.Length-1;
            StringBuilder someString = new StringBuilder(razmer);
            while (razmer[size] != '/')
            {
                size--;
            }
            someString[size] = ' ';
            razmer = someString.ToString();
            razmer = razmer.Replace(" "," r");

            if (!current_name_Olta.Contains(razmer))
                return false;

            string[] for_fix = new string[] { Pitstop[1],Pitstop[2], current_name_Olta};
            for_fix = Fix_Olta(for_fix,"brand");
            current_name_Olta = for_fix[2];
            Pitstop[1] = for_fix[0];
            Pitstop[2] = for_fix[1];

            if (!current_name_Olta.Contains(Pitstop[1]))
                return false;

            if (!current_name_Olta.Contains(Pitstop[2]))
                return false;
            else return true;
        }
        //------------------------------------------------------------------------------------------------------------------------------------------
        public int[] Discovering_position_in_Olta ()
            {
            excelapp2 = new Excel.Application();                               //Создание объкта EXCEL
            excelappworkbook2 = excelapp2.Workbooks.Open(textBox_Olta.Text);   //Обращение через объект к книге EXCEL 
            excelworksheet2 = excelappworkbook2.ActiveSheet;
            int current_row_Olta = 1, column_of_name_Olta = 1, row_of_names_Olta = 0, column_of_price_Olta=1;
            string find_aim = "";
            int[] pos_Olta = new int[] { 0, 1, 2, 3,4};
            
            //1) -------------------------------------------------------------------
            while (find_aim != "Номенклатура")
            {
                if (excelworksheet2.Cells[current_row_Olta, column_of_name_Olta].Value != null)
                {
                    find_aim = excelworksheet2.Cells[current_row_Olta, column_of_name_Olta].Value.ToString();
                }
                current_row_Olta++;
                if (current_row_Olta >= 20)
                {
                    column_of_name_Olta++;
                    current_row_Olta = 1;
                }
            }
            row_of_names_Olta = current_row_Olta - 1;
            pos_Olta[0] = current_row_Olta;
            pos_Olta[1] = column_of_name_Olta;

            current_row_Olta = 1;

            while ((find_aim != "Цена") && !find_aim.Contains("Розни") && !find_aim.Contains("розни"))
            {
                if (excelworksheet2.Cells[row_of_names_Olta, column_of_price_Olta].Value != null)
                {
                    if (excelworksheet2.Cells[row_of_names_Olta, column_of_price_Olta].Value.ToString().Contains("Цена"))
                    {
                        find_aim = "Цена";
                    }
                    column_of_price_Olta++;
                }
                else
                    column_of_price_Olta++;
            }
            column_of_price_Olta--;
            
            
            pos_Olta[2] = current_row_Olta;
            pos_Olta[3] = column_of_price_Olta;

            current_row_Olta = 1;

            while (excelworksheet2.Cells[current_row_Olta, column_of_name_Olta].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 1, column_of_name_Olta].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 2, column_of_name_Olta].Value != null ||
                excelworksheet2.Cells[current_row_Olta, column_of_name_Olta + 1].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 1, column_of_name_Olta + 1].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 2, column_of_name_Olta + 1].Value != null ||
                excelworksheet2.Cells[current_row_Olta, column_of_name_Olta + 2].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 1, column_of_name_Olta + 2].Value != null ||
                excelworksheet2.Cells[current_row_Olta + 2, column_of_name_Olta + 2].Value != null)
            {
                current_row_Olta++;

            }
            pos_Olta[4] = current_row_Olta;

                //excelappworkbook2.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                /*System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworksheet2);
                excelappworkbook2.Close(0);
                excelapp2.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelapp2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelappworkbook2);
                excelapp2 = null;
                excelappworkbook2 = null;
                excelworksheet2 = null;*/
                return pos_Olta;
            }

        public int[] Discovering_position_in_Pitstop(string name_of_store)
        {
            excelapp = new Excel.Application();                               //Создание объкта EXCEL
            excelappworkbook = excelapp.Workbooks.Open(textBox_Pitstop.Text);   //Обращение через объект к книге EXCEL 
            excelworksheet = excelappworkbook.ActiveSheet;
            int current_row_Pitstop = 1, column_of_name_Pitstop = 1, row_of_names_Pitstop = 0, column_of_price_Pitstop = 1;
            string find_aim = "";
            int[] pos_Pitstop = new int[] { 0, 1, 2, 3,4,5};

            while (!find_aim.Contains("Номенклатура"))
            {
                if (excelworksheet.Cells[current_row_Pitstop, column_of_name_Pitstop].Value != null)
                {
                    find_aim = excelworksheet.Cells[current_row_Pitstop, column_of_name_Pitstop].Value.ToString();
                }
                current_row_Pitstop++;
                if (current_row_Pitstop >= 20)
                {
                    column_of_name_Pitstop++;
                    current_row_Pitstop = 1;
                }
            }
            row_of_names_Pitstop= current_row_Pitstop - 1;
            pos_Pitstop[0] = current_row_Pitstop - 1;
            pos_Pitstop[1] = column_of_name_Pitstop;
            current_row_Pitstop = 1;
            
            while (!find_aim.Contains("Цена"))
            {
                if (excelworksheet.Cells[current_row_Pitstop, column_of_price_Pitstop].Value != null)
                {
                    find_aim = excelworksheet.Cells[current_row_Pitstop, column_of_price_Pitstop].Value.ToString();
                }
                current_row_Pitstop++;
                if (current_row_Pitstop >= 20)
                {
                    column_of_price_Pitstop++;
                    current_row_Pitstop = 1;
                }
            }
            pos_Pitstop[3] = column_of_price_Pitstop;

            int check_nulls = 0;
            while (check_nulls<10)
            {
                if (excelworksheet.Cells[current_row_Pitstop, column_of_price_Pitstop].Value != null)
                {
                    current_row_Pitstop++;
                    check_nulls = 0;
                }
                else
                {
                    check_nulls++;
                    current_row_Pitstop++;
                }
            }

            pos_Pitstop[4] = current_row_Pitstop-10;
            check_nulls = 0;
            while (check_nulls!=4)
            {
                if (excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop].Value != null)
                {
                    column_of_price_Pitstop++;
                }
                else
                    check_nulls++;
            }
            if (name_of_store == "Olta")
            {
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop].Value = "Олта";
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop + 1].Value = "Олта2";
            }
            if (name_of_store == "Nortek")
            {
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop].Value = "Нортек";
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop + 1].Value = "Нортек2";
            }
            if (name_of_store == "Shintorg")
            {
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop].Value = "Шинторг";
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop + 1].Value = "Шинторг2";
            }
            if (name_of_store == "Scot")
            {
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop].Value = "Скотченко";
                excelworksheet.Cells[row_of_names_Pitstop, column_of_price_Pitstop + 1].Value = "Скотченко2";
            }
            pos_Pitstop[5] = column_of_price_Pitstop; 
            return pos_Pitstop;
        }

        public int[] Discovering_position_in_Nortek()
        {
            excelapp2 = new Excel.Application();                               //Создание объкта EXCEL
            excelappworkbook2 = excelapp2.Workbooks.Open(textBox_Nortek.Text);   //Обращение через объект к книге EXCEL 
            excelworksheet2 = excelappworkbook2.ActiveSheet;

            int current_row_Nortek = 1, column_of_name_Nortek = 1, row_of_names_Nortek = 0, column_of_price_Nortek = 1;
            string find_aim = "";
            int[] pos_Nortek = new int[] { 0, 1, 2, 3, 4 };
            //1) -------------------------------------------------------------------
            while (find_aim != "Номенклатура")
            {
                if (excelworksheet2.Cells[current_row_Nortek, column_of_name_Nortek].Value != null)
                {
                    find_aim = excelworksheet2.Cells[current_row_Nortek, column_of_name_Nortek].Value.ToString();
                }
                current_row_Nortek++;
                if (current_row_Nortek >= 20)
                {
                    column_of_name_Nortek++;
                    current_row_Nortek = 1;
                }
            }
            row_of_names_Nortek = current_row_Nortek - 1;
            pos_Nortek[0] = current_row_Nortek;
            pos_Nortek[1] = column_of_name_Nortek;

            current_row_Nortek = 1;

            while ((find_aim != "Цена") && !find_aim.Contains("Розни") && !find_aim.Contains("розни"))
            {
                if (excelworksheet2.Cells[row_of_names_Nortek, column_of_price_Nortek].Value != null)
                {
                    if (excelworksheet2.Cells[row_of_names_Nortek, column_of_price_Nortek].Value.ToString().Contains("Цена"))
                    {
                        find_aim = "Цена";
                    }
                    column_of_price_Nortek++;
                }
                else
                    column_of_price_Nortek++;
            }
            column_of_price_Nortek--;


            pos_Nortek[2] = current_row_Nortek;
            pos_Nortek[3] = column_of_price_Nortek;

            current_row_Nortek = 1;

            int check_nulls = 0;
            while (check_nulls < 10)
            {
                if (excelworksheet2.Cells[current_row_Nortek, column_of_price_Nortek].Value != null)
                {
                    current_row_Nortek++;
                    check_nulls = 0;
                }
                else
                {
                    check_nulls++;
                    current_row_Nortek++;
                }
            }
            pos_Nortek[4] = current_row_Nortek-10;

            return pos_Nortek;
        }
        //------------------------------------------------------------------------------------------------------------------------------------------
        private void button_compare_Olta_Click(object sender, EventArgs e)
        {
            string[] parts_of_decomposition_Pitstop = new string[] { "razmer", "brend", "model", "komplekt", "ship", "index" };
            int[] pos_Pitstop = new int[] { 0, 1, 2, 3, 4, 5 };
            int[] pos_Olta = new int[] { 0, 1, 2, 3, 4 };            
            //string[] useless = new string[] { "matador", "continental", "laufenn", "pirelli", "gislaved", "sava", "dunlop", "goodyear", "maxxis", "michelin", "nexen"};
            string current_name_Olta = "";
            bool check_simmiliar = false, IsNum = false;
            int i = 0;
            //-----------------------------------------------------------------------------------------------------------------------------------
            pos_Olta = Discovering_position_in_Olta();
            //1) Ищем колонку наименований и цен в олте
            //[0] ряд имен row_of_names_Olta;
            //[1] колонка имен column_of_names_Olta;
            //[2] ряд цен   row_of_price_Olta;
            //[3] колонка цен  column_of_price_Olta;
            //[4] колонка конца файла
            int row_of_names_Olta = pos_Olta[0];
            int column_of_name_Olta = pos_Olta[1];
            //begin_row_Olta = pos_Olta[0] + 1;
            //current_row_Olta = begin_row_Olta;
            int column_price_Olta = pos_Olta[3];
            int end_of_rows_Olta = pos_Olta[4];
            int current_row_Olta = row_of_names_Olta;
            //-----------------------------------------------------------------------------------------------------------------------------------
            pos_Pitstop = Discovering_position_in_Pitstop("Olta");
            //1) Ищем колонку наименований и цен в олте
            //[0] ряд имен row_of_names_Pitstop;1
            //[1] колонка имен column_of_names_Pitstop;2
            //[2] ряд цен   row_of_price_Pitstop;---
            //[3] колонка цен  column_of_price_Pitstop;3
            //[4] колонка конца файла end_of price_Pitstop
            //[5] номер колонки куда вставлять цены олты в питстопе
            int row_of_names_PitStop = pos_Pitstop[0];
            int column_of_names_Pitstop = pos_Pitstop[1];
            int colum_of_price_Pitstop = pos_Pitstop[3];
            int end_of_rows_Pitstop = pos_Pitstop[4];
            int column_new_price_Pitstop = pos_Pitstop[5];
            int current_row_Pitstop = row_of_names_PitStop;
            //---------------------------------------------------------------------------------------------------------------------------------------
            //Переходим к обработке каждой строки питстопа(Общий цикл)    
            while (current_row_Pitstop < end_of_rows_Pitstop) 
            {
                //Ищем позицию с ценой в столбце с ценой (если енсть цена значит проверяем товар на сравнение)
                while (IsNum == false && current_row_Pitstop<end_of_rows_Pitstop)
                {
                    if (excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value != null)
                    {
                        string stroka = excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value.ToString();
                        for (i = 0; i < stroka.Length; i++)
                            if (stroka[i] >= '0' && stroka[i] <= '9' || stroka[i] == ',')
                            {
                                IsNum = true;
                            }
                        else
                            {
                                current_row_Pitstop++;
                                IsNum = false;
                                break;
                            }
                    }
                    else
                    {
                        current_row_Pitstop++;
                    }
                }
                IsNum = false;
                //декомпозируем строку на отдельные части в ПИТСТОПЕ
                parts_of_decomposition_Pitstop = decomposition_of_Pitstop(excelworksheet.Cells[current_row_Pitstop, column_of_names_Pitstop].Value.ToString());
                while (check_simmiliar!=true && current_row_Olta<=end_of_rows_Olta)
                {//вторая часть цикла по поиску наименований в ОЛТЕ
                    //Ищем позицию с ценой для проверки
                    while (IsNum == false && current_row_Olta<end_of_rows_Olta)
                    {
                        if (excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                        {
                            string stroka = excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value.ToString();
                            for (i = 0; i < stroka.Length; i++)
                                if (stroka[i] >= '0' && stroka[i] <= '9' || stroka[i] == ',')
                                {
                                    IsNum = true;
                                }
                                else
                                {
                                    current_row_Olta++;
                                    IsNum = false;
                                    break;
                                }
                        }
                        else
                        {
                            current_row_Olta++;
                        }
                    }
                    IsNum = false;

                    if (excelworksheet2.Cells[current_row_Olta, column_of_name_Olta].Value != null)
                        current_name_Olta = excelworksheet2.Cells[current_row_Olta, column_of_name_Olta].Value.ToString();
                    current_name_Olta = current_name_Olta.ToLower();
                    check_simmiliar = Compare_Pitstop_Olta(parts_of_decomposition_Pitstop,current_name_Olta);
                    if (check_simmiliar == true)
                    {
                        if (excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value == null)
                        {
                            if (excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value = excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value;
                            else
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value = "Нет цены";
                        }
                        else
                                    if (excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                            excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop + 1].Value = excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value.ToString();
                        else
                            excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop + 1].Value = "Нет цены";
                        if (excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value != null && excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                        {
                            int pit = Convert.ToInt32(excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value.ToString());

                            int olt = Convert.ToInt32(excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value);
                            if (pit > olt)
                            {
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Interior.Color = 16776960;
                            }
                        }
                    }
                    else
                        current_row_Olta++;
                    textBox_process.Text = current_row_Pitstop.ToString() +" В Питстоп и " + current_row_Olta.ToString() + "В олте";
                }
                current_row_Pitstop++;
                check_simmiliar = false;
                current_row_Olta = 1;
            }
            excelappworkbook.Save();
            excelappworkbook.Close();
            excelappworkbook2.Save();
            excelappworkbook2.Close();

            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
            textBox_process.Text = "Работа завершена в Олта";
        }

        private void button_compare_Nortek_Click(object sender, EventArgs e)
        {
            string[] useless = new string[] { "matador", "continental", "laufenn", "pirelli", "gislaved", "sava", "dunlop", "goodyear", "maxxis", "michelin", "nexen", "алтайшина", "nortec", "annaite", "aufine", "power", "compasal", "goodtyre" };
            string[] parts_of_decomposition_Pitstop = new string[] { "razmer", "brend", "model", "komplekt", "ship", "index" };
            string  current_name_Nortek = "";
            bool check_simmiliar = false, IsNum = false;
            int i = 0;
            int[] pos_Pitstop = new int[] { 0, 1, 2, 3, 4, 5 };
            int[] pos_Nortek = new int[] { 0, 1, 2, 3, 4 };
            //-----------------------------------------------------------------------------------------------------------------------------------
            pos_Nortek = Discovering_position_in_Nortek();
            //1) Ищем колонку наименований и цен в олте
            //[0] ряд имен row_of_names_Olta;
            //[1] колонка имен column_of_names_Olta;
            //[2] ряд цен   row_of_price_Olta;
            //[3] колонка цен  column_of_price_Olta;
            //[4] колонка конца файла
            int row_of_names_Nortek = pos_Nortek[0];
            int column_of_name_Nortek = pos_Nortek[1];
            //begin_row_Olta = pos_Olta[0] + 1;
            //current_row_Olta = begin_row_Olta;
            int column_price_Nortek = pos_Nortek[3];
            int end_of_rows_Nortek = pos_Nortek[4];
            int current_row_Nortek = row_of_names_Nortek;
            //-----------------------------------------------------------------------------------------------------------------------------------
            pos_Pitstop = Discovering_position_in_Pitstop("Nortek");
            //1) Ищем колонку наименований и цен в олте
            //[0] ряд имен row_of_names_Pitstop;1
            //[1] колонка имен column_of_names_Pitstop;2
            //[2] ряд цен   row_of_price_Pitstop;---
            //[3] колонка цен  column_of_price_Pitstop;3
            //[4] колонка конца файла end_of price_Pitstop
            //[5] номер колонки куда вставлять цены олты в питстопе
            int row_of_names_PitStop = pos_Pitstop[0];
            int column_of_names_Pitstop = pos_Pitstop[1];
            int colum_of_price_Pitstop = pos_Pitstop[3];
            int end_of_rows_Pitstop = pos_Pitstop[4];
            int column_new_price_Pitstop = pos_Pitstop[5];
            int current_row_Pitstop = row_of_names_PitStop;
            //------------------------------------------------------ОБЩИЙ ЦИКЛ---------------------------------------------------------------
            textBox_process.Text = "Начинаем сравнение...";
            while (current_row_Pitstop < end_of_rows_Pitstop)
            {
                //Ищем позицию с ценой в столбце с ценой (если енсть цена значит проверяем товар на сравнение)
                while (IsNum == false && current_row_Pitstop < end_of_rows_Pitstop)
                {
                    if (excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value != null)
                    {
                        string stroka = excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value.ToString();
                        for (i = 0; i < stroka.Length; i++)
                            if (stroka[i] >= '0' && stroka[i] <= '9' || stroka[i] == ',')
                            {
                                IsNum = true;
                            }
                            else
                            {
                                current_row_Pitstop++;
                                IsNum = false;
                                break;
                            }
                    }
                    else
                    {
                        current_row_Pitstop++;
                    }
                }
                IsNum = false;
                //декомпозируем строку на отдельные части в ПИТСТОПЕ
                parts_of_decomposition_Pitstop = decomposition_of_Pitstop(excelworksheet.Cells[current_row_Pitstop, column_of_names_Pitstop].Value.ToString());
                while (check_simmiliar != true && current_row_Nortek <= end_of_rows_Nortek)
                {//вторая часть цикла по поиску наименований в ОЛТЕ
                    //Ищем позицию с ценой для проверки
                    while (IsNum == false && current_row_Nortek < end_of_rows_Nortek)
                    {
                        if (excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value != null)
                        {
                            string stroka = excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value.ToString();
                            for (i = 0; i < stroka.Length; i++)
                                if (stroka[i] >= '0' && stroka[i] <= '9' || stroka[i] == ',')
                                {
                                    IsNum = true;
                                }
                                else
                                {
                                    current_row_Nortek++;
                                    IsNum = false;
                                    break;
                                }
                        }
                        else
                        {
                            current_row_Nortek++;
                        }
                    }
                    IsNum = false;

                    if (excelworksheet2.Cells[current_row_Nortek, column_of_name_Nortek].Value != null)
                        current_name_Nortek = excelworksheet2.Cells[current_row_Nortek, column_of_name_Nortek].Value.ToString();
                    current_name_Nortek = current_name_Nortek.ToLower();
                    check_simmiliar = Compare_Pitstop_Nortek(parts_of_decomposition_Pitstop, current_name_Nortek);
                    if (check_simmiliar == true)
                    {
                        if (excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value == null)
                        {
                            if (excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value != null)
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value = excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value;
                            else
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Value = "Нет цены";
                        }
                        else
                                    if (excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value != null)
                            excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop + 1].Value = excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value.ToString();
                        else
                            excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop + 1].Value = "Нет цены";
                        if (excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value != null && excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value != null)
                        {
                            int pit = Convert.ToInt32(excelworksheet.Cells[current_row_Pitstop, colum_of_price_Pitstop].Value.ToString());

                            int olt = Convert.ToInt32(excelworksheet2.Cells[current_row_Nortek, column_price_Nortek].Value);
                            if (pit > olt)
                            {
                                excelworksheet.Cells[current_row_Pitstop, column_new_price_Pitstop].Interior.Color = 16776960;
                            }
                        }
                    }
                    else
                        current_row_Nortek++;
                    textBox_process.Text = current_row_Pitstop.ToString() + " В Питстоп и " + current_row_Nortek.ToString() + "В Нортек";
                }
                current_row_Pitstop++;
                check_simmiliar = false;
                current_row_Nortek = 1;
            }

            excelappworkbook.Save();
            excelappworkbook.Close();
            excelappworkbook2.Save();
            excelappworkbook2.Close();


            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
            textBox_process.Text = "Работа завершена в нортек";
        }

        private void button_compare_Shintorg_Click(object sender, EventArgs e)
        {
            excelapp = new Excel.Application();                               //Создание объкта EXCEL
            excelappworkbook = excelapp.Workbooks.Open(textBox_Pitstop.Text);   //Обращение через объект к книге EXCEL 
            excelworksheet = excelappworkbook.ActiveSheet;

            excelapp2 = new Excel.Application();                               //Создание объкта EXCEL
            excelappworkbook2 = excelapp2.Workbooks.Open(textBox_Shintorg.Text);   //Обращение через объект к книге EXCEL 
            excelworksheet2 = excelappworkbook2.ActiveSheet;

            string[] parts_of_decomposition_Pitstop = new string[] { "razmer", "brend+model", "brend", "model", "komplekt", "ship", "index" };
            string[] parts_of_decomposition_Olta = new string[] { "", "brend+model", "brend", "model", "komplekt", "ship", "index" };
            string[] useless = new string[] { "matador", "continental", "laufenn", "pirelli", "gislaved", "sava", "dunlop", "goodyear", "maxxis", "michelin", "nexen", "алтайшина", "nortec", "annaite", "aufine", "power", "compasal", "goodtyre" };
            string find_aim = "", current_name_Olta = "";
            int current_row_Pitstop = 1, column_name_Pitstop = 1, begin_row_Pitstop = 0, row_of_names_PitStop = 0, end_of_rows_Pitstop = 0,
                current_column_Pitstop = 1;
            bool check_simmiliar = false, check_del = false;
            int current_row_Olta = 1, current_column_Olta = 1, column_name_Olta = 1, begin_row_Olta = 0, row_of_names_Olta = 0;
            int i = 0;
            //-----------------------------------------------------------------------------------------------------------------------------------
            
            //Поиск ячейки первого имени(в питстопе по ячейке "наименование") Для ПИТСТОП
            textBox_process.Text = "Ищем строку начала наименований в ПИТСТОП....";
            while (find_aim != "ТМЦ")
            {
                if (excelworksheet.Cells[current_row_Pitstop, current_column_Pitstop].Value != null)
                {
                    find_aim = excelworksheet.Cells[current_row_Pitstop, current_column_Pitstop].Value.ToString();
                }
                current_row_Pitstop++;

                if (current_row_Pitstop >= 20)
                {
                    current_column_Pitstop++;
                    current_row_Pitstop = 1;
                }
            }
            row_of_names_PitStop = current_row_Pitstop - 1;
            column_name_Pitstop = current_column_Pitstop;
            //Теперь знаем строку первой позиции
            begin_row_Pitstop = row_of_names_PitStop + 3;
            //готовим значения для прогонки
            current_row_Pitstop = begin_row_Pitstop;
            //поиск колонок цены и итого
            int column_Itogo = column_name_Pitstop;
            int column_price_Pitstop = column_name_Pitstop;
            int finding = column_name_Pitstop;
            int column_price_Olta = column_name_Olta;
            textBox_process.Text = "Ищем колонку ''итого по складам'' в ПИТСТОП....";
            while (find_aim != "Итого по складам")
            {
                if (excelworksheet.Cells[row_of_names_PitStop, finding].Value != null)
                {
                    if (excelworksheet.Cells[row_of_names_PitStop, finding].Value.ToString().Contains("цена"))
                    {
                        column_price_Pitstop = finding;
                    }
                    if (excelworksheet.Cells[row_of_names_PitStop, finding].Value.ToString().Contains("Итого по складам"))
                    {
                        column_Itogo = finding;
                        find_aim = "Итого по складам";
                    }
                    finding++;
                }
                else
                    finding++;
            }
            int column_of_Olta_in_Pitstop = finding + 7;
            excelworksheet.Cells[row_of_names_PitStop, column_of_Olta_in_Pitstop].Value = "Шинторг";
            textBox_process.Text = "Ищем конец файла в ПИТСТОП....";
            //Находим конец нашего файла(строку где null во всех столбцах 3x3)
            while (excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 1, column_name_Pitstop].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 2, column_name_Pitstop].Value != null ||
            excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop + 1].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 1, column_name_Pitstop + 1].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 2, column_name_Pitstop + 1].Value != null ||
            excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop + 2].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 1, column_name_Pitstop + 2].Value != null ||
            excelworksheet.Cells[current_row_Pitstop + 2, column_name_Pitstop + 2].Value != null)
            {
                current_row_Pitstop++;
                //textBox_Pitstop.Text = current_row_Pitstop.ToString();

            }
            end_of_rows_Pitstop = current_row_Pitstop - 1;
            current_row_Pitstop = begin_row_Pitstop;
            //------------------------------------------------------------------------------------------------------------------------
            //Поиск ячейки первого имени(в питстопе по ячейке "наименование") Для Шинторг
            textBox_process.Text = "Ищем начало наименование в ШИНТОРГ....";
            while (!find_aim.Contains("Номенклатура"))
            {
                if (excelworksheet2.Cells[current_row_Olta, current_column_Olta].Value != null)
                {
                    find_aim = excelworksheet2.Cells[current_row_Olta, current_column_Olta].Value.ToString();
                }
                current_row_Olta++;
                if (current_row_Olta >= 20)
                {
                    current_column_Olta++;
                    current_row_Olta = 1;
                }
            }
            //Запоминаем строку названий для столбцов
            row_of_names_Olta = current_row_Olta - 1;
            column_name_Olta = current_column_Olta;
            //Теперь знаем строку первой позиции
            begin_row_Olta = row_of_names_Olta + 1;
            //готовим значения для прогонки
            current_row_Olta = begin_row_Olta;
            textBox_process.Text = "Поиск колонки цен в НОРТЕК....";
            //теперь поиск колонки цен в Нортек+
            int find_nulls = 0;
            while (!find_aim.Contains("Цена"))
            {
                if (excelworksheet2.Cells[row_of_names_Olta, column_price_Olta].Value != null)
                {
                    if (excelworksheet2.Cells[row_of_names_Olta, column_price_Olta].Value.ToString().Contains("Цена") || excelworksheet2.Cells[row_of_names_Olta, column_price_Olta].Value.ToString().Contains("Розница"))
                    {
                        find_aim = "Цена";
                    }
                    column_price_Olta++;
                    find_nulls = 0;
                }
                else
                {
                    find_nulls++;
                    column_price_Olta++;
                }
                if (find_nulls >= 4)
                {
                    column_price_Olta = 1;
                    row_of_names_Olta++;
                    find_nulls = 0;
                }
            }
            column_price_Olta--;


            check_del = false;
            textBox_process.Text = "Проверяем наличие левых позиций в НОРТЕК....";
            //Очистим файл нортек от ненужных позиций
            while (excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta, column_name_Olta + 2].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta + 2].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta + 2].Value != null)
            {

                if (excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value != null)
                {
                    textBox_process.Text = current_row_Olta.ToString();

                    /*if (!excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString().Contains(" Автошина") && !excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString().Contains(" Автопокрышка")
                        && !excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString().Contains(" ОШЗ") && !excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString().Contains(" ЯШЗ"))
                    {
                        string lol3 = excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString();
                       
                        Excel.Range rg = (Excel.Range)excelworksheet2.Rows[current_row_Olta, Type.Missing];
                        rg.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        int lol2 = current_row_Olta;
                        check_del = true;
                    }*/
                    /* if (excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value != null)
                     {*/
                    string stroka = excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString();

                    stroka = stroka.ToLower();

                    foreach (string lol in useless)
                    {

                        if (stroka.Contains(lol))
                        {
                            Excel.Range rg = (Excel.Range)excelworksheet2.Rows[current_row_Olta, Type.Missing];
                            rg.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            int lol2 = current_row_Olta;
                            check_del = true;
                            stroka = excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString();
                            stroka = stroka.ToLower();

                        }
                    }

                    /*}*/
                    if (check_del == true)
                    {
                        textBox_process.Text = current_row_Olta + "Удаление... ";
                        check_del = false;
                    }
                    else
                        current_row_Olta++;
                }
                else
                    current_row_Olta++;
            }



            //------------------------------------------------------ОБЩИЙ ЦИКЛ---------------------------------------------------------------
            textBox_process.Text = "Начинаем сравнение...";
            current_row_Pitstop = begin_row_Pitstop;
            current_row_Olta = begin_row_Olta;

            while (current_row_Pitstop < end_of_rows_Pitstop)
            {
                // textBox_process.Text = current_row_Pitstop.ToString() ;
                if (excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString() == "1 Зимние шины" ||
                    excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString() == "2 Летние шины" ||
                    excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString() == "3 Легкогрузовые шины" ||
                    excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString() == "Итого:")
                {
                    current_row_Pitstop++;
                }
                if (current_row_Pitstop >= end_of_rows_Pitstop)
                    break;
                //декомпозируем строку на отдельные части в ПИТСТОПЕ
                parts_of_decomposition_Pitstop = decomposition_of_Pitstop(excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString());
                // textBox_process.Text = excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Value.ToString();
                while (check_simmiliar != true &&
                  (excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta].Value != null ||
               excelworksheet2.Cells[current_row_Olta, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta + 1].Value != null ||
               excelworksheet2.Cells[current_row_Olta, column_name_Olta + 2].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 1, column_name_Olta + 2].Value != null ||
               excelworksheet2.Cells[current_row_Olta + 2, column_name_Olta + 2].Value != null))
                {//вторая часть цикла по поиску наименований в Нортек
                    if (excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value != null)
                        current_name_Olta = excelworksheet2.Cells[current_row_Olta, column_name_Olta].Value.ToString();
                    current_name_Olta = current_name_Olta.ToLower();

                    if (current_name_Olta.Contains("зима") || current_name_Olta.Contains("грузовые") || current_name_Olta.Contains("сельхоз") || current_name_Olta.Contains("погрузчики"))
                    {
                        //приводим к одному размеру чтобы сравнить с ПИТСТОП
                        current_name_Olta = current_name_Olta.Replace(" автошина", "");
                        current_name_Olta = current_name_Olta.Replace(" автопокрышка", "");
                        current_name_Olta = current_name_Olta.Replace(" ошз", "");
                        current_name_Olta = current_name_Olta.Replace(" яшз", "");

                        string first = "", second = "";
                        i = 0;
                        while (current_name_Olta[i] != ' ')
                        {
                            first = first + current_name_Olta[i];
                            i++;
                        }
                        i++;
                        while (current_name_Olta[i] != ' ')
                        {
                            second = second + current_name_Olta[i];
                            i++;
                        }
                        parts_of_decomposition_Olta[0] = first + second;
                        parts_of_decomposition_Olta[0] = parts_of_decomposition_Olta[0].Replace("r", "/");
                        parts_of_decomposition_Olta[0] = parts_of_decomposition_Olta[0].Replace("/", " ");
                        if (parts_of_decomposition_Olta[0].Contains("c"))
                        {
                            parts_of_decomposition_Olta[0].Replace("c", "");
                            parts_of_decomposition_Pitstop[0].Replace("c", "");
                        }

                        //6 - индекс
                        parts_of_decomposition_Pitstop[6] = parts_of_decomposition_Pitstop[6].ToLower();
                        current_name_Olta = current_name_Olta.Replace(" " + parts_of_decomposition_Pitstop[6], "");
                        textBox_process.Text = current_row_Pitstop.ToString() + " " + current_row_Olta.ToString() + "Olta";
                        current_name_Olta = current_name_Olta.Replace("_", " ");

                        current_name_Olta = current_name_Olta.Replace("-", " ");


                        if (parts_of_decomposition_Olta[0] != parts_of_decomposition_Pitstop[0])
                        {
                            i = 0;
                            current_row_Olta++;
                        }
                        else

                        {
                            //current_name_Olta = current_name_Olta.Replace(" " + parts_of_decomposition_Pitstop[6], "");
                            i = 0;

                            /*if (current_name_Olta.Contains("yokohama"))
                            {
                                if (current_name_Olta.Contains("ig"))
                                    current_name_Olta = current_name_Olta.Replace("ig", "ig ");
                                else
                                    current_name_Olta = current_name_Olta.Replace("g0", "g 0");
                            }
                            if (current_name_Olta.Contains("bridgestone"))
                            {
                                if (current_name_Olta.Contains("ice cruiser 7000"))
                                    current_name_Olta = current_name_Olta.Replace("ice cruiser 7000", "ic7000");

                            }*/
                            if (current_name_Olta.Contains("viatti") && parts_of_decomposition_Pitstop[3].Contains("viatti"))
                            {
                                int k = 5;

                            }

                            if (current_name_Olta.Contains("nokian"))
                            {
                                if (current_name_Olta.Contains("rs2") && parts_of_decomposition_Pitstop[3].Contains("rs2 xl"))
                                {

                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace(" xl", "");
                                }
                                if (current_name_Olta.Contains("hakkapeliitta 8") && parts_of_decomposition_Pitstop[3].Contains("hkpl 8"))
                                {

                                    current_name_Olta = current_name_Olta.Replace("hakkapeliitta 8", "hkpl 8");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace(" xl", "");
                                }
                                if (current_name_Olta.Contains("hakkapeliitta 9") && parts_of_decomposition_Pitstop[3].Contains("hkpl 9"))
                                {

                                    current_name_Olta = current_name_Olta.Replace("hakkapeliitta 9", "hkpl 9");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace(" xl", "");
                                }
                                if (current_name_Olta.Contains("nordman 7") && parts_of_decomposition_Pitstop[3].Contains("nordman 7"))
                                {
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace(" xl", "");
                                }
                                if (current_name_Olta.Contains("rs2") && parts_of_decomposition_Pitstop[3].Contains("rs2"))
                                {
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("   xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                }
                                if (current_name_Olta.Contains("nordman 5") && parts_of_decomposition_Pitstop[3].Contains("nordman 5"))
                                {
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace("  xl", "");
                                    parts_of_decomposition_Pitstop[3] = parts_of_decomposition_Pitstop[3].Replace(" xl", "");
                                }
                            }

                            if (current_name_Olta.Contains("кама"))
                            {
                                if (current_name_Olta.Contains("euro") && current_name_Olta.Contains("519"))
                                {
                                    current_name_Olta = current_name_Olta.Replace("кама", "кама евро 519");

                                }



                            }
                            if (current_name_Olta.Contains("nitto"))
                            {
                                if (current_name_Olta.Contains("therma spike"))
                                {
                                    current_name_Olta = current_name_Olta.Replace("therma spike", "ntspk");

                                }
                            }
                            if (current_name_Olta.Contains("bridgestone"))
                            {
                                if (current_name_Olta.Contains("ice cruiser"))
                                {

                                    current_name_Olta = current_name_Olta.Replace("ice cruiser ", "ic");
                                }

                                if (current_name_Olta.Contains("spike 02") && current_name_Olta.Contains("suv") && current_name_Olta.Contains("xl"))
                                {
                                    current_name_Olta = current_name_Olta.Replace("suv", "");
                                    current_name_Olta = current_name_Olta.Replace("xl", "");
                                    current_name_Olta = current_name_Olta.Replace("spike 02", "spike 02 suv xl");
                                }
                            }
                            if (current_name_Olta.Contains("cordiant"))
                            {
                                if (current_name_Olta.Contains("snow cross 2") && parts_of_decomposition_Pitstop[3] != "snow cross 2")
                                {
                                    current_name_Olta = current_name_Olta.Replace("snow cross 2", "");
                                }
                            }
                            if (current_name_Olta.Contains("toyo"))
                            {
                                if (current_name_Olta.Contains("observe garit giz"))
                                {
                                    current_name_Olta = current_name_Olta.Replace("observe garit giz", "obgiz");
                                }
                            }

                            if (current_name_Olta.Contains(parts_of_decomposition_Pitstop[3]))
                            {
                                check_simmiliar = true;
                                //excelworksheet.Cells[current_row_Pitstop, column_name_Pitstop].Interior.Color = 12659;
                                if (excelworksheet.Cells[current_row_Pitstop, column_of_Olta_in_Pitstop].Value == null)
                                {
                                    if (excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                                        excelworksheet.Cells[current_row_Pitstop, column_of_Olta_in_Pitstop].Value = excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value;
                                    else
                                        excelworksheet.Cells[current_row_Pitstop, column_of_Olta_in_Pitstop].Value = "Ошибка";
                                }
                                else
                                    excelworksheet.Cells[current_row_Pitstop, column_of_Olta_in_Pitstop + 1].Value = excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value.ToString();
                                if (excelworksheet.Cells[current_row_Pitstop, column_price_Pitstop].Value != null && excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value != null)
                                {
                                    int pit = Convert.ToInt32(excelworksheet.Cells[current_row_Pitstop, column_price_Pitstop].Value.ToString());

                                    int olt = Convert.ToInt32(excelworksheet2.Cells[current_row_Olta, column_price_Olta].Value);
                                    if (pit > olt)
                                    {
                                        excelworksheet.Cells[current_row_Pitstop, column_of_Olta_in_Pitstop].Interior.Color = 16776960;
                                    }

                                }

                            }
                            else
                                current_row_Olta++;
                        }

                    }
                    else
                    {
                        current_row_Olta++;
                        textBox_process.Text = current_row_Pitstop.ToString() + " " + current_row_Olta.ToString() + "Olta";

                    }
                    parts_of_decomposition_Olta[0] = "";
                }
                parts_of_decomposition_Olta[0] = "";
                current_row_Pitstop++;
                check_simmiliar = false;

                current_row_Olta = begin_row_Olta;
            }
            excelappworkbook.Save();
            excelappworkbook.Close();
            excelappworkbook2.Save();
            excelappworkbook2.Close();


            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
            textBox_process.Text = "Работа завершена в нортек";

        }

        private void textBox_process_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
