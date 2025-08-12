using Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;
using DataTable = System.Data.DataTable;
using word = Microsoft.Office.Interop.Word;

namespace It_formulaire
{
    public partial class recupe : UserControl
    {
        private static recupe a;
        public static recupe Instance
        {
            get
            {
                if (a == null)
                {
                    a = new recupe();
                }
                return a;
            }
        }
        MySqlConnection con = new MySqlConnection("datasource=10.10.81.112;port=3306;username=root;password=;database=recupe");
        public recupe()
        {
            InitializeComponent();
            con.Open();
            string req = "select DISTINCT * from user";


            MySqlCommand Hol = new MySqlCommand(req, con);
            MySqlDataReader read = Hol.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(read);
            dataGridView1.DataSource = table;

            con.Close();
        }

        private void recupe_Load(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            string req = "insert into user (Nom,date) values ('" + comboBox1.Text + "','" + textBox1.Text + "')";

            MySqlCommand Holl = new MySqlCommand(req, con);

            MySqlDataReader reader2 = Holl.ExecuteReader();

            comboBox1.Text = "";
            textBox1.Text = "";

            con.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[rowIndex];
            comboBox1.Text = row.Cells[0].Value.ToString();
            textBox1.Text = row.Cells[1].Value.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            label1.Text = DateTime.Today.ToString("dd/MM/yyyy");

            if (comboBox1.Text == "SAHNINE Alaa eddine")
            {
                label9.Text = "Administrateur système Junior";
                label10.Text = "0770208731";
                label11.Text = "A.SAHNINE@groupe-metidji.com";
                label12.Text = "SAHNINE";
                label13.Text = "Alaa eddine";
            }
            if (comboBox1.Text == "MAIZA Bachir")
            {
                label9.Text = "Administrateur réseau Junior";
                label10.Text = "0770208738";
                label11.Text = "B.Maiza@groupe-metidji.com";
                label12.Text = "MAIZA";
                label13.Text = "Bachir";
            }
            if (comboBox1.Text == "HAKIKI Mohammed Oussama")
            {
                label9.Text = "Administrateur Sécurité Réseau Junior";
                label10.Text = "0770208734";
                label11.Text = "M.Hakiki@groupe-metidji.com";
                label12.Text = "HAKIKI";
                label13.Text = "Mohammed Oussama";
            }
            if (comboBox1.Text == "OULD KEBAILI Aimene")
            {
                label9.Text = "IT HELPDESK";
                label10.Text = "770208871";
                label11.Text = "A.OULDKEBAILI@groupe-metidji.com";
                label12.Text = "OULD KEBAILI";
                label13.Text = "Aimene";
            }

            if (comboBox1.Text == "CHABANE CHAOUCH Leila")
            {
                label9.Text = "Analyste fonctionnel";
                label10.Text = "0770209093";
                label11.Text = "L.chabane@groupe-metidji.com";
                label12.Text = "CHABANE CHAOUCH";
                label13.Text = "Leila";
            }
            if (comboBox1.Text == "ZAOUIDI Nadir")
            {
                label9.Text = "Analyste fonctionnel";
                label10.Text = "0770215132";
                label11.Text = "N.ZAOUIDI@groupe-metidji.com";
                label12.Text = "ZAOUIDI";
                label13.Text = "Nadir";
            }
            if (comboBox1.Text == "SENOUCI Youcef")
            {
                label9.Text = "IT Engineer";
                label10.Text = "0770208914";
                label11.Text = "Y.SENOUCI@groupe-metidji.com";
                label12.Text = "SENOUCI";
                label13.Text = "Youcef";
            }





            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath + @"\word\Modèle Demande de congé et récupération HOLDING 2020.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark Nom = doc.Bookmarks["Nom"];
            word.Bookmark Prenom = doc.Bookmarks["Prenom"];
            word.Bookmark Fonction = doc.Bookmarks["Fonction"];
            word.Bookmark RecupConge = doc.Bookmarks["RecupConge"];
            word.Bookmark DateTravail = doc.Bookmarks["DateTravail"];
            word.Bookmark NmbrJours = doc.Bookmarks["NmbrJours"];
            word.Bookmark DateD = doc.Bookmarks["DateD"];
            word.Bookmark DateR = doc.Bookmarks["DateR"];
            word.Bookmark Telephone = doc.Bookmarks["Telephone"];
            word.Bookmark Adresse = doc.Bookmarks["Adresse"];
            word.Bookmark Date = doc.Bookmarks["Date"];




            word.Range rNom = Nom.Range;
            word.Range rPrenom = Prenom.Range;
            word.Range rFonction = Fonction.Range;
            word.Range rRecupConge = RecupConge.Range;
            word.Range rDateTravail = DateTravail.Range;
            word.Range rNmbrJours = NmbrJours.Range;
            word.Range rDateD = DateD.Range;
            word.Range rDateR = DateR.Range;
            word.Range rTelephone = Telephone.Range;
            word.Range rAdresse = Adresse.Range;
            word.Range rDate = Date.Range;



            rNom.Text = label12.Text.ToString();
            rPrenom.Text = label13.Text.ToString();
            rFonction.Text = label9.Text.ToString();
            rRecupConge.Text = "Récupération";
            rDateTravail.Text = textBox1.Text.ToString();
            rNmbrJours.Text = "01";
            //rDateD.Text = textBox1.Text.ToString();
            //rDateR.Text = textBox5.Text.ToString();
            rTelephone.Text = label10.Text.ToString();
            rAdresse.Text = label11.Text.ToString();
            rDate.Text = label1.Text.ToString();



            doc.Bookmarks.Add("Nom", rNom);
            doc.Bookmarks.Add("Prenom", rPrenom);
            doc.Bookmarks.Add("Fonction", rFonction);
            doc.Bookmarks.Add("RecupConge", rRecupConge);
            doc.Bookmarks.Add("DateTravail", rDateTravail);
            doc.Bookmarks.Add("NmbrJours", rNmbrJours);
            doc.Bookmarks.Add("DateD", rDateD);
            doc.Bookmarks.Add("DateR", rDateR);
            doc.Bookmarks.Add("Telephone", rTelephone);
            doc.Bookmarks.Add("Adresse", rAdresse);
            doc.Bookmarks.Add("Date", rDate);



            app.Documents.Open(Application.StartupPath + @"\word\Modèle Demande de congé et récupération HOLDING 2020.docx");

            con.Open();
            string req = "delete from user where Nom ='" + comboBox1.Text.ToString() + "'and date = '" + textBox1.Text.ToString() + "'";


            MySqlCommand Hol = new MySqlCommand(req, con);
            MySqlDataReader read = Hol.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(read);
            dataGridView1.DataSource = table;

            con.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            con.Open();
            string req = "select DISTINCT * from user";


            MySqlCommand Hol = new MySqlCommand(req, con);
            MySqlDataReader read = Hol.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(read);
            dataGridView1.DataSource = table;

            con.Close();
        }
    }
}
