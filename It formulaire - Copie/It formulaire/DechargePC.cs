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
    public partial class DechargePC : UserControl

    {
        private static DechargePC a;
        public static DechargePC Instance
        {
            get
            {
                if (a == null)
                {
                    a = new DechargePC();
                }
                return a;
            }
        }
        public DechargePC()
        {
            InitializeComponent();
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = DateTime.Today.ToString("dd/MM/yyyy");

            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath + @"\word\Modèle DECHARGE PC.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark matricule = doc.Bookmarks["nom_et_prenom"];
            word.Bookmark nom = doc.Bookmarks["fonction"];
            word.Bookmark prenom = doc.Bookmarks["Filiale"];
            word.Bookmark ddn = doc.Bookmarks["Designation"];
            word.Bookmark lieu = doc.Bookmarks["Declare_Le"];
            word.Bookmark telephone = doc.Bookmarks["n_serie"];
            word.Bookmark nnom = doc.Bookmarks["Recep"];
            word.Bookmark ddate = doc.Bookmarks["date_doc"];
            word.Bookmark CPU = doc.Bookmarks["CPU"];
            word.Bookmark RAM = doc.Bookmarks["RAM"];
            word.Bookmark Storage = doc.Bookmarks["disque"];



            word.Range rmatricule = matricule.Range;
            word.Range rnom = nom.Range;
            word.Range rprenom = prenom.Range;
            word.Range rddn = ddn.Range;
            word.Range rlieu = lieu.Range;
            word.Range rtelephone = telephone.Range;
            word.Range rnnom = nnom.Range;
            word.Range rddate = ddate.Range;
            word.Range rCPU = CPU.Range;
            word.Range rRAM = RAM.Range;
            word.Range rStorage = Storage.Range;


            rmatricule.Text = utilisateur.Text.ToString();
            rnom.Text = textBox1.Text.ToString();
            rprenom.Text = textBox5.Text.ToString();
            rddn.Text = textBox3.Text.ToString();
            rlieu.Text = textBox2.Text.ToString();
            rtelephone.Text = textBox4.Text.ToString();
            rnnom.Text = utilisateur.Text.ToString();
            rddate.Text = textBox2.Text.ToString();
            rCPU.Text = textBox6.Text.ToString();
            rRAM.Text = textBox8.Text.ToString();
            rStorage.Text = textBox7.Text.ToString();




            doc.Bookmarks.Add("nom_et_prenom", rmatricule);
            doc.Bookmarks.Add("Filiale", rnom);
            doc.Bookmarks.Add("Date", rprenom);
            doc.Bookmarks.Add("Designation", rddn);
            doc.Bookmarks.Add("Quantite", rlieu);
            doc.Bookmarks.Add("Nserie", rtelephone);
            doc.Bookmarks.Add("Nnom", rnnom);
            doc.Bookmarks.Add("Ddate", rddate);






            app.Documents.Open(Application.StartupPath + @"\word\Modèle DECHARGE PC.docx");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
