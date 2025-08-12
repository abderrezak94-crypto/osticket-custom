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
    public partial class DechargeRestitutionPCc : UserControl
    {
        private static DechargeRestitutionPCc a;
        public static DechargeRestitutionPCc Instance
        {
            get
            {
                if (a == null)
                {
                    a = new DechargeRestitutionPCc();
                }
                return a;
            }
        }
        public DechargeRestitutionPCc()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = DateTime.Today.ToString("dd/MM/yyyy");

            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath+@"\word\Modèle DECHARGE DE RESTITUTION PC.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark matricule = doc.Bookmarks["Nom"];
            word.Bookmark nom = doc.Bookmarks["Agissantqualite"];
            word.Bookmark prenom = doc.Bookmarks["Filiale"];
            word.Bookmark ddn = doc.Bookmarks["Designation"];
            word.Bookmark lieu = doc.Bookmarks["Date"];
            word.Bookmark telephone = doc.Bookmarks["serialnumber"];
            word.Bookmark nnom = doc.Bookmarks["Reception"];
            word.Bookmark ddate = doc.Bookmarks["RemisLe"];
            word.Bookmark CPU = doc.Bookmarks["CPU"];
            word.Bookmark RAM = doc.Bookmarks["RAM"];
            word.Bookmark Storage = doc.Bookmarks["Stockage"];



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






            app.Documents.Open(Application.StartupPath + @"\word\Modèle DECHARGE DE RESTITUTION PC.docx");
        }
    }
}
