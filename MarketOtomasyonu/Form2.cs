using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

namespace MarketOtomasyonu
{
    public partial class Form2 : Form
    {
        private ArrayList type_ids = new ArrayList();
        private ArrayList product_ids = new ArrayList();

        public string sel_barcode { get; set; }

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            listBox2.ClearSelected();
            listBox2.Items.Clear();

            listBox1.Items.Clear();

            this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            ListTypes();
            listBox1.SelectedIndex = 0;
        }

        private void ListTypes()
        {
            type_ids.Add(0);
            listBox1.Items.Add("« Tümü »");

            string Query = "select * from types";
            OleDbCommand cmd = new OleDbCommand(Query, Form1.Con);

            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                type_ids.Add(Convert.ToInt32(result["id"]));
                listBox1.Items.Add(result["type_name"].ToString());
            }
        }

        private void GetBarcode()
        {
            if (!String.IsNullOrEmpty(textBox1.Text))
            {
                this.sel_barcode = textBox1.Text;
                this.Close();
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Query;
            int type_id = Convert.ToInt32(type_ids[listBox1.SelectedIndex]);

            product_ids = new ArrayList();

            listBox2.ClearSelected();
            listBox2.Items.Clear();

            if (type_id == 0)
            {
                Query = "select * from products";
            }
            else
            {
                Query = "select * from products where type_id=@type_id";
            }
            
            OleDbCommand cmd = new OleDbCommand(Query, Form1.Con);

            cmd.Parameters.AddWithValue("@type_id", type_id);
            OleDbDataReader result = cmd.ExecuteReader();

            while (result.Read())
            {
                product_ids.Add(Convert.ToInt32(result["id"]));

                if (Convert.ToInt32(result["stock"]) == 0)
                {
                    listBox2.Items.Add("(Stokta Yok) " + result["product_name"].ToString());
                }
                else
                {
                    listBox2.Items.Add(result["product_name"].ToString());
                }
                
                
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedItems.Count == 0)
            {
                textBox1.Text = null;
                button1.Enabled = false;
            }
            else
            {
                textBox1.Text = null;
                button1.Enabled = true;

                int product_id = Convert.ToInt32(product_ids[listBox2.SelectedIndex]);

                string Query = "select barcode from products where id=@id";
                OleDbCommand cmd = new OleDbCommand(Query, Form1.Con);

                cmd.Parameters.AddWithValue("@id", product_id);
                OleDbDataReader result = cmd.ExecuteReader();

                result.Read();
                textBox1.Text = result["barcode"].ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetBarcode();
        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            GetBarcode();
        }
    }
}
