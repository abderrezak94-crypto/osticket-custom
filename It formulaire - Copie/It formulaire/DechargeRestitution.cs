using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using word = Microsoft.Office.Interop.Word;

namespace It_formulaire
{
    public partial class DechargeRestitution : UserControl
    {
        private static DechargeRestitution a;
        public static DechargeRestitution Instance
        {
            get
            {
                if (a == null)
                {
                    a = new DechargeRestitution();
                }
                return a;
            }
        }
        public DechargeRestitution()
        {
            InitializeComponent();
        }

        private void DechargeRestitution_Load(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {


            textBox2.Text = DateTime.Today.ToString("dd/MM/yyyy");

            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath + @"\word\Modèle DECHARGE DE RESTITUTION 2020.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark matricule = doc.Bookmarks["Nom"];
            word.Bookmark nom = doc.Bookmarks["Filiale"];
            word.Bookmark prenom = doc.Bookmarks["Date"];
            word.Bookmark ddn = doc.Bookmarks["Designation"];
            word.Bookmark lieu = doc.Bookmarks["Quantite"];
            word.Bookmark telephone = doc.Bookmarks["Nserie"];
            word.Bookmark nnom = doc.Bookmarks["Nnom"];
            word.Bookmark ddate = doc.Bookmarks["Ddate"];
            word.Bookmark Agissant = doc.Bookmarks["Agissant"];


            word.Range rmatricule = matricule.Range;
            word.Range rnom = nom.Range;
            word.Range rprenom = prenom.Range;
            word.Range rddn = ddn.Range;
            word.Range rlieu = lieu.Range;
            word.Range rtelephone = telephone.Range;
            word.Range rnnom = nnom.Range;
            word.Range rddate = ddate.Range;
            word.Range rAgissant = Agissant.Range;


            rmatricule.Text = utilisateur.Text.ToString();
            rnom.Text = textBox1.Text.ToString();
            rprenom.Text = textBox2.Text.ToString();
            rddn.Text = textBox5.Text.ToString();
            rlieu.Text = textBox4.Text.ToString();
            rtelephone.Text = textBox3.Text.ToString();
            rnnom.Text = utilisateur.Text.ToString();
            rddate.Text = textBox2.Text.ToString();
            rAgissant.Text = textBox6.Text.ToString();




            doc.Bookmarks.Add("Nom", rmatricule);
            doc.Bookmarks.Add("Filiale", rnom);
            doc.Bookmarks.Add("Date", rprenom);
            doc.Bookmarks.Add("Designation", rddn);
            doc.Bookmarks.Add("Quantite", rlieu);
            doc.Bookmarks.Add("Nserie", rtelephone);
            doc.Bookmarks.Add("Nnom", rnnom);
            doc.Bookmarks.Add("Ddate", rddate);
            doc.Bookmarks.Add("Agissant", rAgissant);






            app.Documents.Open(Application.StartupPath + @"\word\Modèle DECHARGE DE RESTITUTION 2020.docx");
        }
    }
}
