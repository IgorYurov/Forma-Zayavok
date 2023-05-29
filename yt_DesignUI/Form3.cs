using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using yt_DesignUI.Components;
using yt_DesignUI.Controls;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;


namespace yt_DesignUI
{
    public partial class Form3 : Form
    {
        public static string connectString = "Provider=Microsoft.Ace.OLEDB.12.0;" + @"Data Source=|DataDirectory|\\BazaJiEst.accdb";
        static OleDbConnection myConnection = new OleDbConnection(connectString);
        StringBuilder sb;
        public Form3()
        {
            InitializeComponent();

            sb = new StringBuilder();

            //Animator.Start();

            comboBox2.MouseWheel += new MouseEventHandler(comboBox2_MouseWheel);
            comboBox3.MouseWheel += new MouseEventHandler(comboBox3_MouseWheel);
            comboBox4.MouseWheel += new MouseEventHandler(comboBox4_MouseWheel);
            comboBox5.MouseWheel += new MouseEventHandler(comboBox5_MouseWheel);
            comboBox6.MouseWheel += new MouseEventHandler(comboBox6_MouseWheel);
            comboBox7.MouseWheel += new MouseEventHandler(comboBox7_MouseWheel);
            comboBox8.MouseWheel += new MouseEventHandler(comboBox8_MouseWheel);
            comboBox9.MouseWheel += new MouseEventHandler(comboBox9_MouseWheel);
            comboBox10.MouseWheel += new MouseEventHandler(comboBox10_MouseWheel);
            comboBox11.MouseWheel += new MouseEventHandler(comboBox11_MouseWheel);
            comboBox12.MouseWheel += new MouseEventHandler(comboBox12_MouseWheel);

            comboBox3.Enabled = false;
            comboBox5.Enabled = false;
            comboBox7.Enabled = false;
            comboBox10.Visible = false;
            comboBox12.Visible = false;

            egoldsGoogleTextBox1.Enabled = false;
            egoldsGoogleTextBox3.Enabled = false;
            egoldsGoogleTextBox17.Visible = false;
            egoldsGoogleTextBox18.Visible = false;

            if (comboBox1.Items.Count == 0)
            {
                EgoldsFormStyle.fStyle selectedStyle = egoldsFormStyle1.FormStyle;
                comboBox1.DataSource = Enum.GetValues(typeof(EgoldsFormStyle.fStyle));
                comboBox1.SelectedItem = selectedStyle;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            egoldsFormStyle1.FormStyle = (EgoldsFormStyle.fStyle)comboBox1.SelectedItem;
        }
        private void comboBox2_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox3_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox4_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox5_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox6_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox7_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox8_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox9_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox10_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox11_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }
        private void comboBox12_MouseWheel(object sender, EventArgs e)
        {
            ((HandledMouseEventArgs)e).Handled = true;
        }

        private void yt_Button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Выберите путь сохранения!");

            Word.Document doc = null;
            try
            {
                Word.Application app = new Word.Application();
                string source = AppDomain.CurrentDomain.BaseDirectory + @"\\Poihali.dotx";
                doc = app.Documents.Add(source);
                doc.Activate();

                Word.Bookmarks wBookmarks = doc.Bookmarks;

                var str = string.Join(", ", listBox1.Items.Cast<DataItem>());
                doc.Bookmarks["TPTC"].Range.Text = str;
                doc.Bookmarks["Zayavitel"].Range.Text = textBox1.Text;

                if (egoldsGoogleTextBox3.Text != "")
                {
                    doc.Bookmarks["IP"].Range.Text = "";
                    doc.Bookmarks["ООО"].Range.Text = egoldsGoogleTextBox3.Text;
                }
                if (egoldsGoogleTextBox3.Text == "")
                {
                    doc.Bookmarks["IP"].Range.Text = "индивидуального предпринимателя";
                    doc.Bookmarks["ООО"].Range.Text = "";
                }
                doc.Bookmarks["Index0"].Range.Text = egoldsGoogleTextBox6.Text;

                doc.Bookmarks["StrZ"].Range.Text = comboBox6.Text;
                if (comboBox3.Enabled == true)
                {
                    doc.Bookmarks["RegZ"].Range.Text = comboBox3.Text + ",";
                }
                if (comboBox3.Enabled == false)
                {
                    doc.Bookmarks["RegZ"].Range.Text = "";
                }

                doc.Bookmarks["AdressZ"].Range.Text = egoldsGoogleTextBox7.Text;
                if (egoldsGoogleTextBox7.Text == egoldsGoogleTextBox8.Text)
                {
                    doc.Bookmarks["AdrOZ"].Range.Text = "";
                    doc.Bookmarks["OD"].Range.Text = " и адрес осуществления деятельности";
                    doc.Bookmarks["OD2"].Range.Text = "";
                }
                if (egoldsGoogleTextBox7.Text != egoldsGoogleTextBox8.Text)
                {
                    doc.Bookmarks["AdrOZ"].Range.Text = egoldsGoogleTextBox8.Text;
                    doc.Bookmarks["OD"].Range.Text = "";
                    doc.Bookmarks["OD2"].Range.Text = ", адрес осуществления деятельности:";
                }
                doc.Bookmarks["RegNom"].Range.Text = egoldsGoogleTextBox1.Text;
                doc.Bookmarks["Number"].Range.Text = egoldsGoogleTextBox4.Text;
                doc.Bookmarks["Email"].Range.Text = egoldsGoogleTextBox5.Text;

                if (comboBox7.Enabled == false)
                {
                    doc.Bookmarks["Gruppa"].Range.Text = comboBox12.Text + ":";
                    doc.Bookmarks["Product"].Range.Text = egoldsGoogleTextBox13.Text;
                }
                if (comboBox7.Enabled == true)
                {
                    doc.Bookmarks["Gruppa"].Range.Text = egoldsGoogleTextBox13.Text;
                    doc.Bookmarks["Product"].Range.Text = "";
                }
                if (comboBox9.Text == "Серийный выпуск")
                {
                    doc.Bookmarks["Vipusk"].Range.Text = comboBox9.Text;
                    doc.Bookmarks["One"].Range.Text = comboBox10.Text;
                    doc.Bookmarks["Two"].Range.Text = "";
                    doc.Bookmarks["Three"].Range.Text = "";
                    doc.Bookmarks["Four"].Range.Text = "";
                }
                if (comboBox9.Text == "Партия")
                {
                    doc.Bookmarks["Vipusk"].Range.Text = comboBox9.Text + ",";
                    doc.Bookmarks["One"].Range.Text = "";
                    doc.Bookmarks["Two"].Range.Text = egoldsGoogleTextBox18.Text;
                    doc.Bookmarks["Three"].Range.Text = "тонн,";
                    doc.Bookmarks["Four"].Range.Text = egoldsGoogleTextBox17.Text;
                }
                if (comboBox9.Text == "Единичное изделие")
                {
                    doc.Bookmarks["Vipusk"].Range.Text = comboBox9.Text;
                    doc.Bookmarks["One"].Range.Text = "";
                    doc.Bookmarks["Two"].Range.Text = "";
                    doc.Bookmarks["Three"].Range.Text = "";
                    doc.Bookmarks["Four"].Range.Text = egoldsGoogleTextBox17.Text;
                }
                doc.Bookmarks["Kod"].Range.Text = egoldsGoogleTextBox14.Text;
                doc.Bookmarks["Izgotovitel"].Range.Text = egoldsGoogleTextBox9.Text;
                doc.Bookmarks["Index"].Range.Text = egoldsGoogleTextBox10.Text;
                doc.Bookmarks["StranaIzgotovitelya"].Range.Text = comboBox4.Text;
                if (comboBox5.Enabled == true)
                {
                    doc.Bookmarks["RegI"].Range.Text = comboBox5.Text + ",";
                }
                if (comboBox5.Enabled == false)
                {
                    doc.Bookmarks["RegI"].Range.Text = "";
                }
                doc.Bookmarks["AdressI"].Range.Text = egoldsGoogleTextBox11.Text;
                if (egoldsGoogleTextBox11.Text == egoldsGoogleTextBox12.Text)
                {
                    doc.Bookmarks["AdrOI"].Range.Text = "";
                    doc.Bookmarks["OD3"].Range.Text = " и адрес осуществления деятельности";
                    doc.Bookmarks["OD4"].Range.Text = "";
                }
                if (egoldsGoogleTextBox11.Text != egoldsGoogleTextBox12.Text)
                {
                    doc.Bookmarks["AdrOI"].Range.Text = egoldsGoogleTextBox12.Text;
                    doc.Bookmarks["OD3"].Range.Text = "";
                    doc.Bookmarks["OD4"].Range.Text = ", адрес осуществления деятельности:";
                }
                doc.Bookmarks["Docs"].Range.Text = egoldsGoogleTextBox16.Text;
                doc.Bookmarks["TPTC2"].Range.Text = str;
                doc.Bookmarks["Days"].Range.Text = comboBox8.Text;
                doc.Bookmarks["TPTC3"].Range.Text = str;
                doc.Bookmarks["TPTC4"].Range.Text = str;
                doc.Bookmarks["Dopolnitelno"].Range.Text = egoldsGoogleTextBox15.Text;
                doc.Bookmarks["Name"].Range.Text = egoldsGoogleTextBox2.Text;

                myConnection.Open();

                string queryString = "INSERT INTO Tablica12(Zayavitel, RegNom, FioVim, Tel, Email, FioVrod, Strana, Region, IndexZ, YurAdrZ, AdrOsDeZ, Izgotovitel, IndexI, StranaI, RegionI, YurAdrI, AdrOsDeI, Naim, KodTnVed, DopInfa, DocVsoot, Shema, Meneger, Data) values('" + textBox1.Text + "','" + egoldsGoogleTextBox1.Text + "', '" + egoldsGoogleTextBox2.Text + "', '" + egoldsGoogleTextBox4.Text + "', '" + egoldsGoogleTextBox5.Text + "', '" + egoldsGoogleTextBox3.Text + "','" + comboBox6.Text + "', '" + comboBox3.Text + "', '" + egoldsGoogleTextBox6.Text + "', '" + egoldsGoogleTextBox7.Text + "','" + egoldsGoogleTextBox8.Text + "', '" + egoldsGoogleTextBox9.Text + "', '" + egoldsGoogleTextBox10.Text + "', '" + comboBox4.Text + "', '" + comboBox5.Text + "', '" + egoldsGoogleTextBox11.Text + "', '" + egoldsGoogleTextBox12.Text + "', '" + egoldsGoogleTextBox13.Text + "', '" + egoldsGoogleTextBox14.Text + "', '" + egoldsGoogleTextBox15.Text + "', '" + egoldsGoogleTextBox16.Text + "', '" + comboBox8.Text + "', '" + comboBox11.Text + "', '" + dateTimePicker1.Value + "')";
                OleDbDataAdapter DataAdapter = new OleDbDataAdapter("SELECT * FROM Tablica12", myConnection);
                DataSet dt = new DataSet();
                DataAdapter.Fill(dt);
                //this.tablica1TableAdapter.Fill(this.bazochkaDataSet.Tablica1);
                OleDbCommand command = new OleDbCommand(queryString, myConnection);
                command.ExecuteNonQuery();

                myConnection.Close();

                doc.Close();
                doc = null;
                //app = null;
                //doc.Quit();
                app.Quit();
                //return;
            }
            catch (Exception ex)
            {
                doc.Close();
                doc = null;
                MessageBox.Show("Во время выполнения произошла ошибка!");
            }
        }

        private void yt_Button6_Click(object sender, EventArgs e)
        {
            Hide();
            yt_DesignUI.Form2 f2 = new yt_DesignUI.Form2();
            f2.ShowDialog();
            Close();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string checkedState = checkedListBox1.SelectedItem.ToString();
            listBox1.Items.Add(new DataItem { TPTC = checkedListBox1.SelectedItem.ToString() });
        }

        private void yt_Button2_Click(object sender, EventArgs e)
        {
            foreach (object item in checkedListBox1.CheckedItems)
            {
                sb.Append(item.ToString());
            }
            if (sb.ToString() == "ТР ТС 004/2011 'О безопасности упаковки', утвержден Решением Комиссии Таможенного союза от 16.08.2011 года № 769")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 005/2011 'О безопасности упаковки', утвержден Решением Комиссии Таможенного союза от 16.08.2011 года № 769")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Упаковка для пищевой продукции");
                comboBox12.Items.Add("Укупорочные средства для пищевой продукции");
                comboBox12.Items.Add("Упаковка для парфюмерно-косметической продукции");
                comboBox12.Items.Add("Укупорочные средства для парфюмерно-косметической продукции");
                comboBox12.Items.Add("Упаковка для изделий детского ассортимента");
                comboBox12.Items.Add("Укупорочные средства для изделий детского ассортимента");
                comboBox12.Items.Add("Упаковка для бытовой продукции");
                comboBox12.Items.Add("Укупорочные средства для бытовой продукции");
                comboBox12.Items.Add("Упаковка для продукции промышленного назначения");
                comboBox12.Items.Add("Укупорочные средства для продукции промышленного назначения");
                comboBox12.Items.Add("Упаковка полимерная для продукции промышленного и бытового назначения");
            }
            if (sb.ToString() == "ТР ТС 007/2011 'О безопасности продукции предназначенной для детей и подростков', утвержден Решением Комиссии Таможенного союза от 23.09.2011 года № 797")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Изделия кожгалантерейные");
                comboBox12.Items.Add("Изделия 3-его слоя для детей старше 1 года и подростков");
                comboBox12.Items.Add("Изделия из меха для детей старше 1 года и подростков");
                comboBox12.Items.Add("Головные уборы 2-ого слоя для детей старше 1 года и подростков");
                comboBox12.Items.Add("Обувь валяная грубошерстная");
                comboBox12.Items.Add("Школьно-письменные принадлежности");
                comboBox12.Items.Add("Готовые штучные текстильные изделия для детей и подростков");
                comboBox12.Items.Add("Постельные принадлежности для детей и подростков");
                comboBox12.Items.Add("Обувь валяная грубошерстная для детей и подростков");
                comboBox12.Items.Add("Продукция издательская книжная, журнальная для детей и подростков");
                comboBox12.Items.Add("Соски молочные");
                comboBox12.Items.Add("Соски-пустышки");
                comboBox12.Items.Add("Изделия санитарно-гигиенические разового использования");
                comboBox12.Items.Add("Гигиенические ватные палочки");
                comboBox12.Items.Add("Посуда и столовые приборы для детей до 3 лет");
                comboBox12.Items.Add("Щетки зубные для детей и подростков");
                comboBox12.Items.Add("Массажеры для десен для детей и подростков");
                comboBox12.Items.Add("Изделия 1-го слоя бельевые для детей до 3 лет");
                comboBox12.Items.Add("Изделия чулочно-носочные трикотажные 1-го слоя для детей до 3 лет");
                comboBox12.Items.Add("Головные уборы (летние) 1-го слоя для детей до 3 лет");
            }
            if (sb.ToString() == "ТР ТС 009/2011 'О безопасности парфюмерно-косметической продукции', утвержден Решением Комиссии Таможенного союза от 23.09.2011 года № 799")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 010/2011 'О безопасности машин и оборудования', утвержден Решением Комиссии Таможенного союза от 18.10.2011 года № 823")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Автопогрузчики");
                comboBox12.Items.Add("Аппаратура для подготовки и очистки газов и жидкостей");
                comboBox12.Items.Add("Аппаратура массообменная криогенных систем и установок");
                comboBox12.Items.Add("Аппаратура теплообменная криогенных систем и установок");
                comboBox12.Items.Add("Аппараты водонагревательные и отопительные, работающие на жидком и твердом топливе");
                comboBox12.Items.Add("Арматура промышленная трубопроводная");
                comboBox12.Items.Add("Велосипеды (кроме детских)");
                comboBox12.Items.Add("Вентиляторы промышленные");
                comboBox12.Items.Add("Воздухонагреватели и воздухоохладители");
                comboBox12.Items.Add("Горелки газовые (кроме блочных), встраиваемые в оборудование, предназначенное для использования в технологических процессах на промышленных предприятиях");
                comboBox12.Items.Add("Горелки жидкотопливные, встраиваемые в оборудование, предназначенное для использования в технологических процессах на промышленных предприятиях");
                comboBox12.Items.Add("Горелки комбинированные, встраиваемые в оборудование, предназначенное для использования в технологических процессах на промышленных предприятиях");
                comboBox12.Items.Add("Дизель-генераторы");
                comboBox12.Items.Add("Дробилки");
                comboBox12.Items.Add("Инструмент абразивный");
                comboBox12.Items.Add("Инструмент из природных алмазов");
                comboBox12.Items.Add("Инструмент из синтетических алмазов");
                comboBox12.Items.Add("Инструмент из синтетических сверхтвердых материалов на основе нитрида бора (инструмент из эльбора)");
                comboBox12.Items.Add("Инструмент слесарно-монтажный с изолирующими рукоятками для работы в электроустановках напряжением до 1000 В");
                comboBox12.Items.Add("Компрессоры (воздушные и газовые приводные)");
                comboBox12.Items.Add("Конвейеры");
                comboBox12.Items.Add("Кондиционеры промышленные");
                comboBox12.Items.Add("Котлы отопительные, работающие на жидком и твердом топливе");
                comboBox12.Items.Add("Материалы абразивные");
                comboBox12.Items.Add("Машины для землеройных работ");
                comboBox12.Items.Add("Машины для мелиоративных работ");
                comboBox12.Items.Add("Машины для разработки и обслуживания карьеров");
                comboBox12.Items.Add("Машины дорожные");
                comboBox12.Items.Add("Машины кузнечно-прессовые");
                comboBox12.Items.Add("Машины тягодутьевые");
                comboBox12.Items.Add("Оборудование бумагоделательное");
                comboBox12.Items.Add("Оборудование буровое геолого-разведочное");
                comboBox12.Items.Add("Оборудование газоочистное");
                comboBox12.Items.Add("Оборудование для газопламенной обработки металлов и металлизации изделий");
                comboBox12.Items.Add("Оборудование для жидкого аммиака");
                comboBox12.Items.Add("Оборудование для коммунального хозяйства");
                comboBox12.Items.Add("Оборудование для переработки полимерных материалов");
                comboBox12.Items.Add("Оборудование для подготовки и очистки питьевой воды");
                comboBox12.Items.Add("Оборудование для приготовления строительных смесей");
                comboBox12.Items.Add("Оборудование для промышленности строительных материалов");
                comboBox12.Items.Add("Оборудование для сварки и газотермического напыления");
                comboBox12.Items.Add("Оборудование для химической чистки и крашения одежды и бытовых изделий");
                comboBox12.Items.Add("Оборудование насосное");
                comboBox12.Items.Add("Оборудование нефтегазоперерабатывающее");
                comboBox12.Items.Add("Оборудование нефтепромысловое");
                comboBox12.Items.Add("Оборудование полиграфическое");
                comboBox12.Items.Add("Оборудование прачечное промышленное");
                comboBox12.Items.Add("Оборудование пылеулавливающее");
                comboBox12.Items.Add("Оборудование строительное");
                comboBox12.Items.Add("Оборудование технологическое для выработки асбестовых нитей");
                comboBox12.Items.Add("Оборудование технологическое для выработки стекловолокна");
                comboBox12.Items.Add("Оборудование технологическое для выработки химических волокон");
                comboBox12.Items.Add("Оборудование технологическое для кабельной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для комбикормовой промышленности");
                comboBox12.Items.Add("Оборудование технологическое для легкой промышленности");
                comboBox12.Items.Add("Оборудование технологическое для лесозаготовки");
                comboBox12.Items.Add("Оборудование технологическое для литейного производства");
                comboBox12.Items.Add("Оборудование технологическое для мукомольно-крупяной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для мясомолочной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для общественного питания");
                comboBox12.Items.Add("Оборудование технологическое для пищеблоков");
                comboBox12.Items.Add("Оборудование технологическое для пищевой промышленности");
                comboBox12.Items.Add("Оборудование технологическое для рыбной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для стекольной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для текстильной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для торговли");
                comboBox12.Items.Add("Оборудование технологическое для торфяной промышленности");
                comboBox12.Items.Add("Оборудование технологическое для фарфоровой промышленности");
                comboBox12.Items.Add("Оборудование технологическое для фаянсовой промышленности");
                comboBox12.Items.Add("Оборудование технологическое для элеваторной промышленности");
                comboBox12.Items.Add("Оборудование технологическое и аппаратура для нанесения лакокрасочных покрытий на изделия машиностроения");
                comboBox12.Items.Add("Оборудование химическое");
                comboBox12.Items.Add("Оборудование целлюлозно-бумажное");
                comboBox12.Items.Add("Пилы дисковые с твердосплавными пластинами для обработки древесных материалов");
                comboBox12.Items.Add("Приспособления для грузоподъемных операций");
                comboBox12.Items.Add("Резцы");
                comboBox12.Items.Add("Станки деревообрабатывающие (небытовые)");
                comboBox12.Items.Add("Станки металлообрабатывающие");
                comboBox12.Items.Add("Тали электрические канатные");
                comboBox12.Items.Add("Тали электрические цепные");
                comboBox12.Items.Add("Тракторы промышленные");
                comboBox12.Items.Add("Транспорт производственный напольный безрельсовый");
                comboBox12.Items.Add("Турбины");
                comboBox12.Items.Add("Установки воздухоразделительные и редких газов");
                comboBox12.Items.Add("Установки газотурбинные");
                comboBox12.Items.Add("Установки холодильные");
                comboBox12.Items.Add("Фрезы");
                comboBox12.Items.Add("Фрезы насадные");
            }
            if (sb.ToString() == "ТР ТС 014/2011 'Безопасность автомобильных дорог', утвержден Решением Комиссии Таможенного союза от 18.10.2011 года № 827")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 015/2011 'О безопасности зерна', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 874")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Зерно злаковых культур");
                comboBox12.Items.Add("Масличные культуры");
                comboBox12.Items.Add("Зернобобовые культуры");
            }
            if (sb.ToString() == "ТР ТС 017/2011 'О безопасности продукции легкой промышленности', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 876")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Материалы декоративные");
                comboBox12.Items.Add("Материалы мебельные");
                comboBox12.Items.Add("Мех искусственный");
                comboBox12.Items.Add("Материалы обувные");
                comboBox12.Items.Add("Изделия верхние для взрослых");
                comboBox12.Items.Add("Изделия чулочно-носочные 2-го слоя для взрослых");
                comboBox12.Items.Add("Изделия перчаточные для взрослых");
                comboBox12.Items.Add("Изделия костюмные для взрослых");
                comboBox12.Items.Add("Изделия плательные для взрослых");
                comboBox12.Items.Add("Одежда домашняя для взрослых");
                comboBox12.Items.Add("Головные уборы 2-ого слоя для взрослых");
                comboBox12.Items.Add("Изделия ковровые");
                comboBox12.Items.Add("Изделия кожгалантерейные для взрослых");
                comboBox12.Items.Add("Изделия текстильно-галантерейные для взрослых");
                comboBox12.Items.Add("Обувь для взрослых");
                comboBox12.Items.Add("Изделия кожаные для взрослых");
                comboBox12.Items.Add("Изделия меховые для взрослых");
                comboBox12.Items.Add("Полотна трикотажные для взрослых");
                comboBox12.Items.Add("Материалы бельевые для взрослых");
                comboBox12.Items.Add("Материалы одежные для взрослых");
                comboBox12.Items.Add("Материалы полотенечные для взрослых");
                comboBox12.Items.Add("Белье столовое");
                comboBox12.Items.Add("Белье кухонное");
                comboBox12.Items.Add("Носовые платки");
                comboBox12.Items.Add("Полотенца");
                comboBox12.Items.Add("Простыни купальные для взрослых");
                comboBox12.Items.Add("Платочно-шарфовые изделия для взрослых");
            }
            if (sb.ToString() == "ТР ТС 020/2011 'Электромагнитная совместимость технических средств', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 879")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Продукты из мяса птицы");
                comboBox12.Items.Add("Продукты переработки яиц");
                comboBox12.Items.Add("Изделия хлебобулочные");
                comboBox12.Items.Add("Изделия мукомольно-крупяные");
                comboBox12.Items.Add("Сахар");
                comboBox12.Items.Add("Приправы");
                comboBox12.Items.Add("Пряности");
                comboBox12.Items.Add("Соль поваренная пищевая");
                comboBox12.Items.Add("Изделия кондитерские сахаристые");
                comboBox12.Items.Add("Изделия кондитерские мучные");
                comboBox12.Items.Add("Овощи свежие");
                comboBox12.Items.Add("Фрукты свежие");
                comboBox12.Items.Add("Ягоды свежие");
                comboBox12.Items.Add("Продукция плодоовощная");
                comboBox12.Items.Add("Грибы");
                comboBox12.Items.Add("Орехи");
                comboBox12.Items.Add("Напитки алкогольные");
                comboBox12.Items.Add("Напитки безалкогольные");
                comboBox12.Items.Add("Сухие концентраты для приготовления напитков");
                comboBox12.Items.Add("Крахмало-паточная продукция");
                comboBox12.Items.Add("Кофе");
                comboBox12.Items.Add("Цикорий");
                comboBox12.Items.Add("Чай");
                comboBox12.Items.Add("Продукты пищевые готовые");
                comboBox12.Items.Add("Полуфабрикаты");
                comboBox12.Items.Add("Какао-порошок");
                comboBox12.Items.Add("Пищевые добавки");
                comboBox12.Items.Add("Ароматизаторы пищевые");
                comboBox12.Items.Add("Вещества вкусоароматические");
                comboBox12.Items.Add("Технологические вспомогательные средства");
                comboBox12.Items.Add("Ферментные препараты");
            }
            if (sb.ToString() == "ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Продукты из мяса птицы");
                comboBox12.Items.Add("Продукты переработки яиц");
                comboBox12.Items.Add("Изделия хлебобулочные");
                comboBox12.Items.Add("Изделия мукомольно-крупяные");
                comboBox12.Items.Add("Сахар");
                comboBox12.Items.Add("Приправы");
                comboBox12.Items.Add("Пряности");
                comboBox12.Items.Add("Соль поваренная пищевая");
                comboBox12.Items.Add("Изделия кондитерские сахаристые");
                comboBox12.Items.Add("Изделия кондитерские мучные");
                comboBox12.Items.Add("Овощи свежие");
                comboBox12.Items.Add("Фрукты свежие");
                comboBox12.Items.Add("Ягоды свежие");
                comboBox12.Items.Add("Продукция плодоовощная");
                comboBox12.Items.Add("Грибы");
                comboBox12.Items.Add("Орехи");
                comboBox12.Items.Add("Напитки алкогольные");
                comboBox12.Items.Add("Напитки безалкогольные");
                comboBox12.Items.Add("Сухие концентраты для приготовления напитков");
                comboBox12.Items.Add("Крахмало-паточная продукция");
                comboBox12.Items.Add("Кофе");
                comboBox12.Items.Add("Цикорий");
                comboBox12.Items.Add("Чай");
                comboBox12.Items.Add("Продукты пищевые готовые");
                comboBox12.Items.Add("Полуфабрикаты");
                comboBox12.Items.Add("Какао-порошок");
                comboBox12.Items.Add("Пищевые добавки");
                comboBox12.Items.Add("Ароматизаторы пищевые");
                comboBox12.Items.Add("Вещества вкусоароматические");
                comboBox12.Items.Add("Технологические вспомогательные средства");
                comboBox12.Items.Add("Ферментные препараты");
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881ТР ТС 023/2011 'Технический регламент на соковую продукцию из фруктов и овощей', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 882")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Продукция соковая-сок");
                comboBox12.Items.Add("Продукция соковая-нектар");
                comboBox12.Items.Add("Продукция соковая-напиток сокосодержащий");
                comboBox12.Items.Add("Продукция соковая-морс");
                comboBox12.Items.Add("Продукция соковая-пюре");
                comboBox12.Items.Add("Продукция соковая-мякоть");
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881ТР ТС 024/2011 'Технический регламент на масложировую продукцию', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 883")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Масло растительное");
                comboBox12.Items.Add("Фракции масел растительных");
                comboBox12.Items.Add("Продукция масложировая-жиры рафинированные дезодорированные");
                comboBox12.Items.Add("Продукция масложировая-масла рафинированные дезодорированные");
                comboBox12.Items.Add("Продукция масложировая-маргарины");
                comboBox12.Items.Add("Продукция масложировая-спреды");
                comboBox12.Items.Add("Продукция масложировая-смеси топленые");
                comboBox12.Items.Add("Продукция масложировая-заменители молочного жира");
                comboBox12.Items.Add("Продукция масложировая-эквиваленты масла какао");
                comboBox12.Items.Add("Продукция масложировая-улучшители масла какао SOS-типа");
                comboBox12.Items.Add("Продукция масложировая-заменители масла какао");
                comboBox12.Items.Add("Продукция масложировая-соусы на основе растительных масел");
                comboBox12.Items.Add("Продукция масложировая-майонезы");
                comboBox12.Items.Add("Продукция масложировая-соусы майонезные");
                comboBox12.Items.Add("Продукция масложировая-кремы на растительных маслах");
                comboBox12.Items.Add("Продукция масложировая-глицерин дистиллированный");
                comboBox12.Items.Add("Жиры специального назначения");
            }
            if (sb.ToString() == "ТР ТС 025/2012 'О безопасности мебельной продукции', утвержден Решением Комиссии Таможенного союза от 15.06.2012 года № 32")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Мебель бытовая корпусная (кроме детской)");
                comboBox12.Items.Add("Мебель бытовая (кроме детской)-столы");
                comboBox12.Items.Add("Мебель бытовая (кроме детской)-кровати");
                comboBox12.Items.Add("Мебель бытовая для сидения и лежания (кроме детской)");
                comboBox12.Items.Add("Мебель бытовая (кроме детской)-матрацы");
                comboBox12.Items.Add("Наборы мебели бытовой (кроме детской)");
                comboBox12.Items.Add("Мебель для общественных и административных помещений корпусная (кроме мебели для дошкольных и учебных заведений)");
                comboBox12.Items.Add("Мебель для общественных и административных помещений (кроме мебели для дошкольных и учебных заведений)-столы");
                comboBox12.Items.Add("Мебель для общественных и административных помещений для сидения и лежания (кроме мебели для дошкольных и учебных заведений)");
                comboBox12.Items.Add("Наборы мебели для общественных и административных помещений (кроме мебели для дошкольных и учебных заведений)");
                comboBox12.Items.Add("Кресла для зрительных залов");
                comboBox12.Items.Add("Мебель для предприятий торговли");
            }
            if (sb.ToString() == "ТР ТС 029/2012 'Требования безопасности пищевых добавок, ароматизаторов и технологических вспомогательных средств', утвержден Решением Комиссии Таможенного союза от 20.07.2012 года № 58")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 030/2012 'О требованиях к смазочным материалам, маслам и специальным жидкостям', утвержден Решением Комиссии Таможенного союза от 20.07.2012 года № 59")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
                comboBox12.Items.Add("");
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881ТР ТС 033/2012 'О безопасности молока и молочной продукции', утвержден Решением Комиссии Таможенного союза от 20.07.2012 года № 67")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Продукты молочные");
                comboBox12.Items.Add("Продукты кисломолочные");
                comboBox12.Items.Add("Продукты молокосодержащие");
                comboBox12.Items.Add("Побочные продукты переработки молока");
                comboBox12.Items.Add("Функционально необходимые при производстве продуктов переработки молока компоненты");
            }
            if (sb.ToString() == "ТР ТС 021/2011 'О безопасности пищевой продукции', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 880ТР ТС 022/2011 'Пищевая продукция в части ее маркировки', утвержден Решением Комиссии Таможенного союза от 09.12.2011 года № 881ТР ТС 034/2012 'О безопасности мяса и мясной продукции', утвержден Решением Комиссии Таможенного союза от 09.10.2013 года № 68")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = true;
                comboBox12.Items.Add("Продукты мясные");
                comboBox12.Items.Add("Продукты мясосодержащие");
                comboBox12.Items.Add("Продукты из шпика");
                comboBox12.Items.Add("Желатин пищевой");
                comboBox12.Items.Add("Жир животный пищевой");
            }
            if (sb.ToString() == "ТР ТС 040/2016 'О безопасности рыбы и рыбной продукции', утвержден Решением Комиссии Таможенного союза от 18.10.2016 года № 162")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
            if (sb.ToString() == "ТР ТС 044/2017 'О безопасности упакованной питьевой воды, включая природную минеральную', утвержден Решением Комиссии Таможенного союза от 23.06.2017 года № 46")
            {
                comboBox12.Text = "";
                comboBox12.Items.Clear();
                comboBox12.Visible = false;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void yt_Button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            listBox1.Items.Clear();
            sb.Clear();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "ОГРН")
            {
                egoldsGoogleTextBox1.Text = "";
                egoldsGoogleTextBox1.Enabled = true;
                egoldsGoogleTextBox3.Text = "";
                egoldsGoogleTextBox3.Enabled = true;
            }
            if (comboBox2.Text == "ОГРНИП")
            {
                egoldsGoogleTextBox1.Text = "";
                egoldsGoogleTextBox1.Enabled = true;
                egoldsGoogleTextBox3.Text = "";
                egoldsGoogleTextBox3.Enabled = false;
            }
        }

        private void egoldsToggleSwitch1_CheckedChanged(object sender)
        {
            if (egoldsToggleSwitch1.Checked == true)
            {
                egoldsGoogleTextBox8.Focus();
                egoldsGoogleTextBox8.Text = egoldsGoogleTextBox7.Text;
            }
            if (egoldsToggleSwitch1.Checked == false)
            {
                egoldsToggleSwitch1.Focus();
                egoldsGoogleTextBox8.Text = "";
            }
        }

        private void egoldsToggleSwitch2_CheckedChanged(object sender)
        {
            if (egoldsToggleSwitch2.Checked == true)
            {
                egoldsGoogleTextBox12.Focus();
                egoldsGoogleTextBox12.Text = egoldsGoogleTextBox11.Text;
            }
            if (egoldsToggleSwitch2.Checked == false)
            {
                egoldsToggleSwitch2.Focus();
                egoldsGoogleTextBox12.Text = "";
            }
        }

        private void egoldsToggleSwitch3_CheckedChanged(object sender)
        {
            if (egoldsToggleSwitch3.Checked == true)
            {
                egoldsGoogleTextBox9.Focus();
                egoldsGoogleTextBox9.Text = textBox1.Text;
                egoldsGoogleTextBox10.Focus();
                egoldsGoogleTextBox10.Text = egoldsGoogleTextBox6.Text;
                egoldsGoogleTextBox11.Focus();
                egoldsGoogleTextBox11.Text = egoldsGoogleTextBox7.Text;
                egoldsGoogleTextBox12.Focus();
                egoldsGoogleTextBox12.Text = egoldsGoogleTextBox8.Text;
                comboBox4.Text = comboBox6.Text;
                comboBox5.Text = comboBox3.Text;
            }
            if (egoldsToggleSwitch3.Checked == false)
            {
                egoldsToggleSwitch3.Focus();
                egoldsGoogleTextBox9.Text = "";
                egoldsGoogleTextBox10.Text = "";
                egoldsGoogleTextBox11.Text = "";
                egoldsGoogleTextBox12.Text = "";
                comboBox4.Text = "";
                comboBox5.Text = "";
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox6.Text == "Россия")
            {
                comboBox3.Enabled = true;
            }
            else
            {
                comboBox3.Enabled = false;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.Text == "Россия")
            {
                comboBox5.Enabled = true;
            }
            else
            {
                comboBox5.Enabled = false;
            }
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox12.Visible == false)
            {
                comboBox7.Enabled = false;
            }
            if (comboBox12.Text == "Зерно злаковых культур")
            {
                comboBox7.Items.Clear();
                comboBox7.Enabled = true;
                comboBox7.Items.Add("Пшеница пищевая");
                comboBox7.Items.Add("Пшеница кормовая");
                comboBox7.Items.Add("Ячмень пищевой");
                comboBox7.Items.Add("Ячмень кормовой");
                comboBox7.Items.Add("Кукуруза пищевая");
                comboBox7.Items.Add("Кукуруза кормовая");
                comboBox7.Items.Add("Рожь пищевая");
                comboBox7.Items.Add("Рожь кормовая");
                comboBox7.Items.Add("Овес пищевой");
                comboBox7.Items.Add("Овес кормовой");
            }
            if (comboBox12.Text == "Масличные культуры")
            {
                comboBox7.Items.Clear();
                comboBox7.Enabled = true;
                comboBox7.Items.Add("Подсолнечник пищевой");
                comboBox7.Items.Add("Соя пищевая");
                comboBox7.Items.Add("Соя кормовая");
                comboBox7.Items.Add("Лен пищевой");
                comboBox7.Items.Add("Рапс пищевой");
                comboBox7.Items.Add("Рапс кормовой");
                comboBox7.Items.Add("Сафлор пищевой");
                comboBox7.Items.Add("Сафлор кормовой");
            }
            if (comboBox12.Text == "Зернобобовые культуры")
            {
                comboBox7.Items.Clear();
                comboBox7.Enabled = true;
                comboBox7.Items.Add("Горох пищевой");
                comboBox7.Items.Add("Горох кормовой");
                comboBox7.Items.Add("Нут пищевой");
                comboBox7.Items.Add("Нут кормовой");
            }
            if (comboBox12.Text != "Зерно злаковых культур" & comboBox12.Text != "Масличные культуры" & comboBox12.Text != "Зернобобовые культуры")
            {
                comboBox7.Items.Clear();
                comboBox7.Text = "";
                comboBox7.Enabled = false;
            }
            if (comboBox12.Text == "")
            {
                comboBox7.Items.Clear();
                comboBox7.Text = "";
                comboBox7.Enabled = false;
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (comboBox7.Enabled == false)
            //{
            //    egoldsGoogleTextBox13.Focus();
            //    egoldsGoogleTextBox13.Text = "";
            //    egoldsGoogleTextBox14.Focus();
            //    egoldsGoogleTextBox14.Text = "";
            //    egoldsGoogleTextBox15.Focus();
            //    egoldsGoogleTextBox15.Text = "";
            //    egoldsGoogleTextBox16.Focus();
            //    egoldsGoogleTextBox16.Text = "";
            //}
            if (comboBox7.Enabled == true & comboBox7.Text == "Пшеница пищевая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: пшеница на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1001 19 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 9353-2016 'Пшеница. Технические условия'";
            }
            if (comboBox7.Text == "Пшеница кормовая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: пшеница на кормовые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1001 99 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 54078-2010 пшеница кормовая. Технические условия";
            }
            if (comboBox7.Text == "Ячмень пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: ячмень на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1003 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 28672-2019 Ячмень. Технические условия";
            }
            if (comboBox7.Text == "Ячмень кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: ячмень на кормовые цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1003 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 53900-2010 Ячмень кормовой. Технические условия";
            }
            if (comboBox7.Text == "Кукуруза пищевая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: кукуруза на продовольственные цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1005 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 13634-90 Кукуруза. Требования при заготовках и поставках";
            }
            if (comboBox7.Text == "Кукуруза кормовая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: кукуруза на кормовые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1005 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 53903-2010 Кукуруза кормовая. Технические условия ";
            }
            if (comboBox7.Text == "Рожь пищевая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: рожь на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1002 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 16990-2017. Рожь. Технические условия";
            }
            if (comboBox7.Text == "Рожь кормовая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: рожь на кормовые цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1002 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен.";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 54079-2010 Рожь кормовая. Технические условия";
            }
            if (comboBox7.Text == "Овес пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: овес на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1004 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 28673-2019. Овес. Технические условия";
            }
            if (comboBox7.Text == "Овес кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зерно злаковых культур: овес на кормовые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1004 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 53901-2010 Овес кормовой. Технические условия";
            }
            if (comboBox7.Text == "Подсолнечник пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: подсолнечник на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1206 00 990 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 22391-2015 Подсолнечник. Технические условия";
            }
            if (comboBox7.Text == "Соя пищевая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: соя на продовольственные цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1201 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 17109-88 Соя. Требования при заготовках и поставках";
            }
            if (comboBox7.Text == "Соя кормовая")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: соя на кормовые цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1201 90 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 17109-88 Соя. Требования при заготовках и поставках";
            }
            if (comboBox7.Text == "Лен пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: лен масличный на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1204 009 00 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 10582-76 Семена льна масличного. Промышленное сырье. Технические условия";
            }
            if (comboBox7.Text == "Рапс пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: рапс для промышленной переработки на пищевые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1205 90 000 9";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 10583-76 «Рапс для промышленной переработки. Технические условия»";
            }
            if (comboBox7.Text == "Рапс кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: Рапс для промышленной переработки на кормовые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1205 90 000 9";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 10583-76 «Рапс для промышленной переработки. Технические условия»";
            }
            if (comboBox7.Text == "Сафлор пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: сафлор на продовольственные цели. Урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1207 60 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 12096-76 «Сафлор для переработки. Технические условия»";
            }
            if (comboBox7.Text == "Сафлор кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Масличные культуры: сафлор для переработки. Урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "1207 60 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 12096-76 «Сафлор для переработки. Технические условия»";
            }
            if (comboBox7.Text == "Горох пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зернобобовые культуры: горох на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "0713 10 900 9";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 28674-2019 «Горох. Технические условия»";
            }
            if (comboBox7.Text == "Горох кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зернобобовые культуры: горох на кормовые цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "0713 10 900 1";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ Р 54630-2011 «Горох кормовой. Технические условия»";
            }
            if (comboBox7.Text == "Нут пищевой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зернобобовые культуры: нут на продовольственные цели, урожай 2021 года";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "0713 20 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 8758-76 Нут. Требования при заготовках и поставках";
            }
            if (comboBox7.Text == "Нут кормовой")
            {
                egoldsGoogleTextBox13.Focus();
                egoldsGoogleTextBox13.Text = "Зернобобовые культуры: нут на кормовые цели, урожай 2021 года ";
                egoldsGoogleTextBox14.Focus();
                egoldsGoogleTextBox14.Text = "0713 20 000 0";
                egoldsGoogleTextBox15.Focus();
                egoldsGoogleTextBox15.Text = "Хранить в чистых, сухих, без постороннего запаха, не зараженных вредителями транспортных средствах и зернохранилищах в соответствии с правилами перевозок, действующими на транспорте данного вида. При соблюдении условий хранения срок годности продукции не ограничен";
                egoldsGoogleTextBox16.Focus();
                egoldsGoogleTextBox16.Text = "ГОСТ 8758-76 Нут. Требования при заготовках и поставках";
            }
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox9.Text == "Серийный выпуск")
            {
                comboBox10.Visible = true;
                egoldsGoogleTextBox17.Visible = false;
                egoldsGoogleTextBox18.Visible = false;
            }
            if (comboBox9.Text == "Партия")
            {
                comboBox10.Visible = false;
                egoldsGoogleTextBox17.Visible = true;
                egoldsGoogleTextBox18.Visible = true;
            }
            if (comboBox9.Text == "Единичное изделие")
            {
                comboBox10.Visible = false;
                egoldsGoogleTextBox17.Visible = true;
                egoldsGoogleTextBox18.Visible = false;
            }
        }

        private void egoldsGoogleTextBox15_TextChanged(object sender, EventArgs e)
        {
            //panel1.Focus();
        }

        private void yt_Button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Выберите путь сохранения!");

            Word.Document doc = null;
            try
            {
                Word.Application app = new Word.Application();
                string source = AppDomain.CurrentDomain.BaseDirectory + @"\\Act.dotx";
                doc = app.Documents.Add(source);
                doc.Activate();

                Word.Bookmarks wBookmarks = doc.Bookmarks;

                doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                doc.Bookmarks["Obrazec"].Range.Text = egoldsGoogleTextBox13.Text;
                doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " т";
                doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text;
                doc.Bookmarks["Podpis"].Range.Text = egoldsGoogleTextBox2.Text;
                if (comboBox12.Text == "Масличные культуры")
                {
                    doc.Bookmarks["Gost"].Range.Text = "ГОСТ 10852-86";
                }
                if (comboBox12.Text == "Зерно злаковых культур" || comboBox12.Text == "Зернобобовые культуры")
                {
                    doc.Bookmarks["Gost"].Range.Text = "ГОСТ 13586.3 - 2015";
                }
                doc.Close();
                doc = null;
                app.Quit();
            }
            catch (Exception ed)
            {
                doc.Close();
                doc = null;
                MessageBox.Show("Во время выполнения произошла ошибка!");
            }
        }

        private void yt_Button5_Click(object sender, EventArgs e)
        {
            if (comboBox7.Text == "Пшеница пищевая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActPshenicaPish.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Пшеница кормовая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActPshenicaKorm.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Рапс кормовой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActRapsKorm.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Рапс пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActRapsPish.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Рожь кормовая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActRoschKorm.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Рожь пищевая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActRoschPish.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Соя кормовая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActSoyaKorm.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Соя пищевая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActSoyaPishevaya.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Горох пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActGorohPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Горох кормовой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActGorohKormovoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Кукуруза кормовая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActKukuruzaKormovaya.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Кукуруза пищевая")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActKukuruzaPishevaya.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Лен пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActLenPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Нут кормовой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActNutKormovoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Нут пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActNutPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Овес кормовой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActOvesKormovoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Овес пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActOvesPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Подсолнечник пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActPodsolnechnikPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Ячмень кормовой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActYachmenKormovoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }

            if (comboBox7.Text == "Ячмень пищевой")
            {
                MessageBox.Show("Выберите путь сохранения!");

                Word.Document doc = null;
                try
                {
                    Word.Application app = new Word.Application();
                    string source = AppDomain.CurrentDomain.BaseDirectory + @"\\ActYachmenPishevoy.dotx";
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    Word.Bookmarks wBookmarks = doc.Bookmarks;

                    doc.Bookmarks["Zakazchik"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Tonn"].Range.Text = egoldsGoogleTextBox18.Text + " тонн";
                    doc.Bookmarks["Izgotov"].Range.Text = egoldsGoogleTextBox9.Text + " " + egoldsGoogleTextBox10.Text + ", " + comboBox4.Text + ", " + comboBox5.Text + ", " + egoldsGoogleTextBox11.Text + "\n" + comboBox2.Text + " " + egoldsGoogleTextBox1.Text;
                    doc.Bookmarks["Otobran"].Range.Text = textBox1.Text + " " + egoldsGoogleTextBox6.Text + ", " + comboBox6.Text + ", " + comboBox3.Text + ", " + egoldsGoogleTextBox7.Text;
                    doc.Bookmarks["Otvetstv"].Range.Text = egoldsGoogleTextBox2.Text;

                    doc.Close();
                    doc = null;
                    app.Quit();
                }
                catch (Exception ed)
                {
                    doc.Close();
                    doc = null;
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                }
            }
        }

        private void egoldsGoogleTextBox15_MouseEnter(object sender, EventArgs e)
        {
            //panel1.Focus();
        }

        private void egoldsGoogleTextBox15_MouseClick(object sender, MouseEventArgs e)
        {
            egoldsGoogleTextBox15.Clear();
        }
    }
}
