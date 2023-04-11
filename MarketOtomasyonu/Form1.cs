using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Collections;
using Microsoft.VisualBasic;

namespace MarketOtomasyonu
{
    public partial class Form1 : Form
    {
        public static OleDbConnection Con;

        private Form2 Form2 = new Form2();

        private ArrayList ticket_products = new ArrayList();
        private ArrayList ticket_counts = new ArrayList();
        private ArrayList ticket_prices = new ArrayList();

        private ArrayList products = new ArrayList();
        private ArrayList types = new ArrayList();

        private OleDbDataReader product;

        private string BackupLoc;
        private string product_img_loc;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Size = new Size(640, 358);
            this.MaximumSize = this.Size;
            this.MinimumSize = this.Size;
            groupBox3.Location = new Point(241, 12);
            groupBox4.Location = new Point(241, 12);
            groupBox5.Location = new Point(241, 12);

            this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            try
            {
                Con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Data\Database.accdb");
                Con.Open();
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message, "Market Otomasyonu - Veritabanı Bağlantı Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }

            Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\Data\Backups");
            Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\Data\Images");

            Backup();
            GetTypes();
            GetProducts();
        }

        private void Backup()
        {
            bool BackupEnabled = false;

            string Query = "select * from settings";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                switch (result["id"].ToString())
                {
                    case "auto_daily_backup":
                        if (result["val"].ToString() == "1")
                        {
                            BackupEnabled = true;
                        }
                        else
                        {
                            BackupEnabled = false;
                        }

                        break;

                    case "backup_loc":
                        BackupLoc = result["val"].ToString();

                        if (!Directory.Exists(BackupLoc))
                        {
                            BackupLoc = Directory.GetCurrentDirectory() + @"\Data\Backups";
                        }

                        break;
                }
            }

            if (!BackupEnabled)
            {
                checkBox1.Checked = false;
            }
            else
            {
                string BackupFileLoc = BackupLoc + @"\" + DateTime.Now.ToString("yyyy-MM-dd") + ".accdb";

                if (!File.Exists(BackupFileLoc))
                {
                    File.Copy(@"Data\Database.accdb", BackupFileLoc);
                }
            }

            textBox2.Text = BackupLoc;
        }

        private void GetTypes()
        {
            types = new ArrayList();

            button13.Enabled = false;
            button15.Enabled = false;

            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            string Query = "select * from types order by id desc";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                types.Add(Convert.ToInt32(result["id"]));
                comboBox1.Items.Add(result["type_name"].ToString());
                comboBox2.Items.Add(result["type_name"].ToString());
            }
        }

        private void GetProducts()
        {
            listBox3.ClearSelected();
            listBox3.Items.Clear();

            products = new ArrayList();

            string Query = "select * from products order by id desc";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                products.Add(Convert.ToInt32(result["id"]));
                listBox3.Items.Add(result["product_name"].ToString());
            }
        }

        private void DeleteTypeImages(int type_id)
        {
            string img_loc;

            string Query = "select id from products where type_id=@type_id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@type_id", type_id);
            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                img_loc = @"Data\Images\" + result["id"].ToString();
                if (File.Exists(img_loc))
                {
                    File.Delete(img_loc);
                }
            }
        }

        private string GetTypeName(int type_id)
        {
            string Query = "select type_name from types where id=@id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@id", type_id);
            OleDbDataReader result = cmd.ExecuteReader();

            if (result.Read())
            {
                return result["type_name"].ToString();
            }

            return "";
        }

        private string GetProductName(int product_id)
        {
            string Query = "select product_name from products where id=@id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@id", product_id);
            OleDbDataReader result = cmd.ExecuteReader();

            result.Read();
            return result["product_name"].ToString();
        }

        private int GetProductDailySale(int product_id, int ticket_id)
        {
            string Query = "select count(*) as total from sales where product_id=@product_id and ticket_id=@ticket_id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@product_id", product_id);
            cmd.Parameters.AddWithValue("@ticket_id", ticket_id);
            OleDbDataReader result = cmd.ExecuteReader();

            result.Read();
            return Convert.ToInt32(result["total"]);
        }

        private string Daily_Total()
        {
            DateTime date = new DateTime(monthCalendar1.SelectionRange.Start.Year, monthCalendar1.SelectionRange.Start.Month, monthCalendar1.SelectionRange.Start.Day);

            string Query = "select sum(price) as daily_total from sales where sale_date=@sale_date";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@sale_date", date);
            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                string value = result["daily_total"].ToString();

                if (!String.IsNullOrEmpty(value))
                {
                    return value;
                }
            }

            return "0,00";
        }

        private string Monthly_Total()
        {
            DateTime start_date = new DateTime(monthCalendar1.SelectionRange.Start.Year, monthCalendar1.SelectionRange.Start.Month, 1);
            DateTime end_date = new DateTime(monthCalendar1.SelectionRange.Start.Year, monthCalendar1.SelectionRange.Start.Month, 1).AddMonths(1);

            string Query = "select sum(price) as monthly_total from sales where sale_date>=@sale_date_start and sale_date<@sale_date_end";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@sale_date_start", start_date);
            cmd.Parameters.AddWithValue("@sale_date_end", end_date);
            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                string value = result["monthly_total"].ToString();

                if (!String.IsNullOrEmpty(value))
                {
                    return value;
                }
            }

            return "0,00";
        }

        private int GetNextBarcode()
        {
            string Query = "select max(barcode) as biggest from products";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            OleDbDataReader result = cmd.ExecuteReader();

            result.Read();

            if (!String.IsNullOrEmpty(result["biggest"].ToString()))
            {
                return Convert.ToInt32(result["biggest"]) + 1;
            }

            return 0;
        }

        private int StockLoss(int product_id)
        {
            int loss = 0;

            for (int i=0; i<listBox1.Items.Count; i++)
            {
                if(Convert.ToInt32(ticket_products[i]) == product_id)
                {
                    loss += Convert.ToInt32(ticket_counts[i]);
                }
            }

            return loss;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;

            groupBox1.Visible = true;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox5.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = true;
            button4.Enabled = true;

            groupBox1.Visible = false;
            groupBox3.Visible = true;
            groupBox4.Visible = false;
            groupBox5.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = false;
            button4.Enabled = true;

            groupBox1.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = true;
            groupBox5.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = false;

            groupBox1.Visible = false;
            groupBox3.Visible = false;
            groupBox4.Visible = false;
            groupBox5.Visible = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            int barcode;

            if (int.TryParse(textBox1.Text, out barcode))
            {
                string Query = "select * from products where barcode=@barcode";
                OleDbCommand cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@barcode", barcode);
                product = cmd.ExecuteReader();

                if (product.Read())
                {
                    pictureBox1.ImageLocation = @"Data\Images\" + product["id"].ToString();

                    label7.Text = product["barcode"].ToString();
                    label6.Text = product["product_name"].ToString();
                    label5.Text = product["price"].ToString() + " TL";

                    int type_id;

                    if (int.TryParse(product["type_id"].ToString(), out type_id))
                    {
                        label8.Text = GetTypeName(Convert.ToInt32(type_id));
                    }

                    numericUpDown1.Maximum = Convert.ToInt32(product["stock"]) - StockLoss(Convert.ToInt32(product["id"]));
                    if(numericUpDown1.Maximum > 0)
                    {
                        numericUpDown1.Value = 1;
                    }

                    return;
                }
            }

            pictureBox1.ImageLocation = null;

            label7.Text = null;
            label8.Text = null;
            label6.Text = null;
            label5.Text = null;

            numericUpDown1.Maximum = 0;
            numericUpDown1.Value = 0;
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown1.Value == 0)
            {
                button8.Enabled = false;
            }
            else
            {
                button8.Enabled = true;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal amount = 0;

            if (listBox1.SelectedIndex == -1)
            {
                button5.Enabled = false;
                button6.Enabled = false;
                button7.Enabled = false;
            }
            else
            {
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = true;
            }

            for(int i = 0; i < ticket_prices.Count; i++)
            {
                amount += Convert.ToDecimal(ticket_prices[i]) * Convert.ToInt32(ticket_counts[i]);
            }

            label24.Text = amount.ToString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            string Query;

            if (checkBox1.Checked)
            {
                Query = "update settings set val='1' where id='auto_daily_backup'";
            }
            else
            {
                Query = "update settings set val='0' where id='auto_daily_backup'";
            }
            

            OleDbCommand cmd = new OleDbCommand(Query, Con);
            cmd.ExecuteNonQuery();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = BackupLoc;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string Query = "update settings set val=@backup_loc where id='backup_loc'";
                OleDbCommand cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@backup_loc", folderBrowserDialog1.SelectedPath);
                cmd.ExecuteNonQuery();

                textBox2.Text = BackupLoc = folderBrowserDialog1.SelectedPath;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if(MessageBox.Show("Geri yüklenecek olan yedek geçerli veritabanınızın üzerine yazılacaktır. Devam etmek istediğinize emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                openFileDialog1.InitialDirectory = BackupLoc;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Con.Close();

                    File.Copy(openFileDialog1.FileName, @"Data\Database.accdb", true);

                    Application.Restart();
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-ffff") + ".accdb";
            saveFileDialog1.InitialDirectory = BackupLoc;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                File.Copy(@"Data\Database.accdb", saveFileDialog1.FileName);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ticket_products.Add(Convert.ToInt32(product["id"]));
            ticket_prices.Add(Convert.ToDecimal(product["price"]));
            ticket_counts.Add(Convert.ToInt32(numericUpDown1.Value));

            listBox1.Items.Add(numericUpDown1.Value + " x " + product["product_name"]);

            textBox1.Text = "";

            listBox1.SelectedIndex = listBox1.Items.Count - 1;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            label26.Text = Daily_Total() + " TL";
            label27.Text = Monthly_Total() + " TL";
        }

        private void button9_Click(object sender, EventArgs e)
        { 
            Form2.ShowDialog();

            if (!String.IsNullOrEmpty(Form2.sel_barcode))
            {
                textBox1.Text = Form2.sel_barcode;
            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            label26.Text = Daily_Total() + " TL";
            label27.Text = Monthly_Total() + " TL";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            button13.Enabled = true;
            button15.Enabled = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Cins ve bu cinsteki ürünler veritabanından kalıcı olarak silinecektir. Devam etmek istediğinize emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                int selindex = comboBox1.SelectedIndex;
                int type_id = Convert.ToInt32(types[selindex]);

                DeleteTypeImages(type_id);

                string Query = "delete from types where id=@type_id";
                OleDbCommand cmd = new OleDbCommand(Query, Con);
                cmd.Parameters.AddWithValue("@type_id", type_id);
                cmd.ExecuteNonQuery();

                Query = "delete from products where type_id=@type_id";
                cmd = new OleDbCommand(Query, Con);
                cmd.Parameters.AddWithValue("@type_id", type_id);
                cmd.ExecuteNonQuery();

                GetTypes();
                GetProducts();
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            int selindex = comboBox1.SelectedIndex;
            int type_id = Convert.ToInt32(types[selindex]);
            string type_name = Interaction.InputBox(" ", " ", comboBox1.Text);
            if (!String.IsNullOrEmpty(type_name))
            {
                string Query = "update types set type_name=@type_name where id=@type_id";
                OleDbCommand cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@type_name", type_name);
                cmd.Parameters.AddWithValue("@type_id", type_id);
                cmd.ExecuteNonQuery();

                GetTypes();
                comboBox1.SelectedIndex = selindex;
                GetProducts();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string Query = "insert into types (type_name) values ('Yeni Cins')";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.ExecuteNonQuery();

            GetTypes();
            comboBox1.SelectedIndex = 0;
            GetProducts();
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            checkBox2.CheckedChanged -= new EventHandler(checkBox2_CheckedChanged);
            numericUpDown3.ValueChanged -= new EventHandler(numericUpDown3_ValueChanged);

            if (listBox3.SelectedItems.Count == 0)
            {
                numericUpDown3.Value = 0;
                comboBox2.SelectedIndex = -1;
                textBox5.Text = "";
                numericUpDown4.Value = 0;
                numericUpDown2.Value = 0;
                checkBox2.Checked = false;

                numericUpDown3.Enabled = false;
                comboBox2.Enabled = false;
                textBox5.Enabled = false;
                numericUpDown4.Enabled = false;
                numericUpDown2.Enabled = false;
                checkBox2.Enabled = false;

                button17.Enabled = false;
                button16.Enabled = false;
            }
            else
            {
                int selindex = listBox3.SelectedIndex;
                int product_id = Convert.ToInt32(products[selindex]);

                string Query = "select * from products where id=@product_id";
                OleDbCommand cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@product_id", product_id);
                OleDbDataReader result = cmd.ExecuteReader();
                result.Read();

                numericUpDown3.Value = Convert.ToInt32(result["barcode"]);
                textBox5.Text = result["product_name"].ToString();
                numericUpDown4.Value = Convert.ToDecimal(result["price"]);
                numericUpDown2.Value = Convert.ToInt32(result["stock"]);

                int type_id;

                comboBox2.SelectedIndex = -1;

                if (int.TryParse(result["type_id"].ToString(), out type_id))
                {
                    for (int i = 0; i < types.Count; i++)
                    {
                        if (type_id == Convert.ToInt32(types[i]))
                        {
                            comboBox2.SelectedIndex = i;
                        }
                    }
                }

                if (File.Exists(product_img_loc = @"Data\Images\" + result["id"].ToString()))
                {
                    checkBox2.Checked = true;
                }
                else
                {
                    checkBox2.Checked = false;
                }

                numericUpDown3.Enabled = true;
                comboBox2.Enabled = true;
                textBox5.Enabled = true;
                numericUpDown4.Enabled = true;
                numericUpDown2.Enabled = true;
                checkBox2.Enabled = true;

                button17.Enabled = true;
                button16.Enabled = true;
            }

            checkBox2.CheckedChanged += new EventHandler(checkBox2_CheckedChanged);
            numericUpDown3.ValueChanged += new EventHandler(numericUpDown3_ValueChanged);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Seçili ürün silinecek. Devam etmek istediğinize emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                if (File.Exists(product_img_loc))
                {
                    File.Delete(product_img_loc);
                }

                int selindex = listBox3.SelectedIndex;
                int product_id = Convert.ToInt32(products[selindex]);

                string Query = "delete from products where id=@product_id";
                OleDbCommand cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@product_id", product_id);
                cmd.ExecuteNonQuery();

                GetProducts();
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string Query = "insert into products (barcode, product_name, stock, price) values (@barcode, 'Yeni Ürün', '0', '0')";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@barcode", GetNextBarcode());
            cmd.ExecuteNonQuery();

            GetProducts();
            listBox3.SelectedIndex = 0;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    File.Copy(openFileDialog2.FileName, product_img_loc);
                }
                else
                {
                    checkBox2.CheckedChanged -= new EventHandler(checkBox2_CheckedChanged);
                    checkBox2.Checked = false;
                    checkBox2.CheckedChanged += new EventHandler(checkBox2_CheckedChanged);
                }
            }
            else
            {
                File.Delete(product_img_loc);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(textBox5.Text) || String.IsNullOrWhiteSpace(textBox5.Text))
            {
                MessageBox.Show("Gerekli alanları doldurunuz.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox5.Focus();
                return;
            }

            int selindex = listBox3.SelectedIndex;
            int product_id = Convert.ToInt32(products[selindex]);

            string Query = "update products set barcode=@barcode, product_name=@product_name, stock=@stock, price=@price, type_id=@type_id where id=@id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@barcode", Convert.ToInt32(numericUpDown3.Value));
            cmd.Parameters.AddWithValue("@product_name", textBox5.Text);
            cmd.Parameters.AddWithValue("@stock", Convert.ToInt32(numericUpDown2.Value));
            cmd.Parameters.AddWithValue("@price", Convert.ToDouble(numericUpDown4.Value));

            if(comboBox2.SelectedIndex == -1)
            {
                cmd.Parameters.AddWithValue("@type_id", DBNull.Value);
            }
            else
            {
                int type_id = Convert.ToInt32(types[comboBox2.SelectedIndex]);
                cmd.Parameters.AddWithValue("@type_id", type_id);
            }

            cmd.Parameters.AddWithValue("@id", product_id);

            cmd.ExecuteNonQuery();

            MessageBox.Show("Ürün kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            GetProducts();
            listBox3.SelectedIndex = selindex;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            int selindex = listBox3.SelectedIndex;
            int product_id = Convert.ToInt32(products[selindex]);

            string Query = "select count(*) as total from products where barcode=@barcode and id<>@product_id";
            OleDbCommand cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@barcode", numericUpDown3.Value);
            cmd.Parameters.AddWithValue("@product_id", product_id);
            OleDbDataReader result = cmd.ExecuteReader();
            result.Read();

            if (Convert.ToInt32(result["total"]) > 0)
            {
                MessageBox.Show("Bu barkod kullanılıyor.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
                numericUpDown3.ValueChanged -= new EventHandler(numericUpDown3_ValueChanged);
                numericUpDown3.Value = GetNextBarcode();
                numericUpDown3.ValueChanged += new EventHandler(numericUpDown3_ValueChanged);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ticket_products.Remove(ticket_products[listBox1.SelectedIndex]);
            ticket_prices.Remove(ticket_prices[listBox1.SelectedIndex]);
            ticket_counts.Remove(ticket_counts[listBox1.SelectedIndex]);

            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string Query;
            OleDbCommand cmd;
            int product_id, count;

            for (int i=0; i<listBox1.Items.Count; i++)
            {
                product_id = Convert.ToInt32(ticket_products[i]);
                count = Convert.ToInt32(ticket_counts[i]);

                Query = "update products set stock=(stock - @count) where id=@product_id";
                cmd = new OleDbCommand(Query, Con);

                cmd.Parameters.AddWithValue("@count", count);
                cmd.Parameters.AddWithValue("@product_id", product_id);
                cmd.ExecuteNonQuery();
            }

            Query = "insert into sales (price) values (@price)";
            cmd = new OleDbCommand(Query, Con);

            cmd.Parameters.AddWithValue("@price", Convert.ToDecimal(label24.Text));
            cmd.ExecuteNonQuery();

            button6.PerformClick();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ticket_products = new ArrayList();
            ticket_prices = new ArrayList();
            ticket_counts = new ArrayList();

            listBox1.SelectedIndex = -1;
            listBox1.Items.Clear();
        }
    }
}
