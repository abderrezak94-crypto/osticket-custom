using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;
using word = Microsoft.Office.Interop.Word;

namespace It_formulaire
{
    public partial class AutorisationAbsence : UserControl
    {
        private static AutorisationAbsence a;
        public static AutorisationAbsence Instance
        {
            get
            {
                if (a == null)
                {
                    a = new AutorisationAbsence();
                }
                return a;
            }
        }
        public AutorisationAbsence()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = DateTime.Today.ToString("dd/MM/yyyy");

            if (comboBox1.Text == "SAHNINE Alaa eddine")
            {
                label8.Text = "Administrateur système Junior";
            }
            if (comboBox1.Text == "MAIZA Bachir")
            {
                label8.Text = "Administrateur réseau Junior";
            }
            if (comboBox1.Text == "BOUKERROUCHA Yacine Oussama")
            {
                label8.Text = "IT Support";
            }
            if (comboBox1.Text == "REGUIBA Taki eddine")
            {
                label8.Text = "IT Support";
            }

            if (comboBox1.Text == "CHABANE CHAOUCH Leila")
            {
                label8.Text = "Analyste fonctionnel";
            }
            if (comboBox1.Text == "ZAOUIDI Nadir")
            {
                label8.Text = "Analyste fonctionnel";
            }
            if (comboBox1.Text == "SENOUCI Youcef")
            {
                label8.Text = "IT Engineer";
            }





            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath + @"\word\Modèle Autorisation d'absence 2020.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark matricule = doc.Bookmarks["Duree"];
            word.Bookmark nom = doc.Bookmarks["Date"];
            word.Bookmark prenom = doc.Bookmarks["DateD"];
            word.Bookmark ddn = doc.Bookmarks["DateR"];
            word.Bookmark lieu = doc.Bookmarks["Motif"];
            word.Bookmark telephone = doc.Bookmarks["NomPrenom"];
            word.Bookmark nnom = doc.Bookmarks["Fonction"];
            /*word.Bookmark ddate = doc.Bookmarks["Ddate"];
            word.Bookmark Agissant = doc.Bookmarks["Agissant"];*/


            word.Range rmatricule = matricule.Range;
            word.Range rnom = nom.Range;
            word.Range rprenom = prenom.Range;
            word.Range rddn = ddn.Range;
            word.Range rlieu = lieu.Range;
            word.Range rtelephone = telephone.Range;
            word.Range rnnom = nnom.Range;
            /*word.Range rddate = ddate.Range;
            word.Range rAgissant = Agissant.Range;*/


            rmatricule.Text = utilisateur.Text.ToString();
            rprenom.Text = textBox1.Text.ToString();
            rnom.Text = textBox2.Text.ToString();
            rddn.Text = textBox5.Text.ToString();
            rlieu.Text = textBox4.Text.ToString();
            rtelephone.Text = comboBox1.Text.ToString();
            rnnom.Text = label8.Text.ToString();
            /*rddate.Text = textBox2.Text.ToString();
            rAgissant.Text = textBox6.Text.ToString();*/




            doc.Bookmarks.Add("Duree", rmatricule);
            doc.Bookmarks.Add("Date", rnom);
            doc.Bookmarks.Add("DateD", rprenom);
            doc.Bookmarks.Add("DateR", rddn);
            doc.Bookmarks.Add("Motif", rlieu);
            /*doc.Bookmarks.Add("Nserie", rtelephone);
            doc.Bookmarks.Add("Nnom", rnnom);
            doc.Bookmarks.Add("Ddate", rddate);
            doc.Bookmarks.Add("Agissant", rAgissant);*/






            app.Documents.Open(Application.StartupPath + @"\word\Modèle Autorisation d'absence 2020.docx");
        }
    }
}
