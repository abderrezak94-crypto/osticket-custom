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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;
using word = Microsoft.Office.Interop.Word;

namespace It_formulaire
{
    public partial class OrdreMission : UserControl
    {
        private static OrdreMission a;
        public static OrdreMission Instance
        {
            get
            {
                if (a == null)
                {
                    a = new OrdreMission();
                }
                return a;
            }
        }
        public OrdreMission()
        {
            InitializeComponent();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = DateTime.Today.ToString("dd/MM/yyyy");

            if (comboBox1.Text == "SAHNINE Alaa eddine")
            {
                label9.Text = "Administrateur système Junior";
                //label10.Text = "0770208731";
                //label11.Text = "A.SAHNINE@groupe-metidji.com";
                //label12.Text = "SAHNINE";
                //label13.Text = "Alaa eddine";
            }
            if (comboBox1.Text == "MAIZA Bachir")
            {
                label9.Text = "Administrateur réseau Junior";
                //label10.Text = "0770208738";
                //label11.Text = "B.Maiza@groupe-metidji.com";
                //label12.Text = "MAIZA";
                //label13.Text = "Bachir";
            }
            if (comboBox1.Text == "BOUKERROUCHA Yacine Oussama")
            {
                label9.Text = "IT Support";
                //label10.Text = "0770208734";
                //label11.Text = "M.Hakiki@groupe-metidji.com";
                //label12.Text = "HAKIKI";
                //label13.Text = "Mohammed Oussama";
            }
            if (comboBox1.Text == "REGUIBA Taki eddine")
            {
                label9.Text = "IT Support";
                //label10.Text = "770208871";
                //label11.Text = "A.OULDKEBAILI@groupe-metidji.com";
                //label12.Text = "OULD KEBAILI";
                //label13.Text = "Aimene";
            }

            if (comboBox1.Text == "CHABANE CHAOUCH Leila")
            {
                label9.Text = "Analyste fonctionnel";
                //label10.Text = "0770209093";
                //label11.Text = "L.chabane@groupe-metidji.com";
                //label12.Text = "CHABANE CHAOUCH";
                //label13.Text = "Leila";
            }
            if (comboBox1.Text == "ZAOUIDI Nadir")
            {
                label9.Text = "Analyste fonctionnel";
                //label10.Text = "0770215132";
                //label11.Text = "N.ZAOUIDI@groupe-metidji.com";
                //label12.Text = "ZAOUIDI";
                //label13.Text = "Nadir";
            }
            if (comboBox1.Text == "SENOUCI Youcef")
            {
                label9.Text = "IT Engineer";
                //label10.Text = "0770208914";
                //label11.Text = "Y.SENOUCI@groupe-metidji.com";
                //label12.Text = "SENOUCI";
                //label13.Text = "Youcef";
            }





            word.Application app = new word.Application();
            word.Document doc = app.Documents.Open(Application.StartupPath + @"\word\Modèle Ordre de mission 2020.docx");
            app.ActiveWindow.View.ReadingLayout = false;

            word.Bookmark Nom = doc.Bookmarks["NomPrenom"];
            //word.Bookmark Prenom = doc.Bookmarks["Prenom"];
            word.Bookmark Fonction = doc.Bookmarks["Fonction"];
            //word.Bookmark RecupConge = doc.Bookmarks["RecupConge"];
            word.Bookmark DateTravail = doc.Bookmarks["Destination"];
            word.Bookmark NmbrJours = doc.Bookmarks["ObjetMission"];
            word.Bookmark DateD = doc.Bookmarks["DateD"];
            word.Bookmark DateR = doc.Bookmarks["DateR"];
            //word.Bookmark Telephone = doc.Bookmarks["Telephone"];
            //word.Bookmark Adresse = doc.Bookmarks["Adresse"];
            word.Bookmark Date = doc.Bookmarks["Date"];




            word.Range rNom = Nom.Range;
            //word.Range rPrenom = Prenom.Range;
            word.Range rFonction = Fonction.Range;
            //word.Range rRecupConge = RecupConge.Range;
            word.Range rDateTravail = DateTravail.Range;
            word.Range rNmbrJours = NmbrJours.Range;
            word.Range rDateD = DateD.Range;
            word.Range rDateR = DateR.Range;
            //word.Range rTelephone = Telephone.Range;
            //word.Range rAdresse = Adresse.Range;
            word.Range rDate = Date.Range;



            rNom.Text = comboBox1.Text.ToString();
            //rPrenom.Text = label13.Text.ToString();
            rFonction.Text = label9.Text.ToString();
            //rRecupConge.Text = comboBox2.Text.ToString();
            rDateTravail.Text = textBox4.Text.ToString();
            rNmbrJours.Text = utilisateur.Text.ToString();
            rDateD.Text = textBox1.Text.ToString();
            rDateR.Text = textBox5.Text.ToString();
            //rTelephone.Text = label10.Text.ToString();
            //rAdresse.Text = label11.Text.ToString();
            rDate.Text = textBox2.Text.ToString();



            doc.Bookmarks.Add("Nom", rNom);
            //doc.Bookmarks.Add("Prenom", rPrenom);
            doc.Bookmarks.Add("Fonction", rFonction);
            //doc.Bookmarks.Add("RecupConge", rRecupConge);
            doc.Bookmarks.Add("DateTravail", rDateTravail);
            doc.Bookmarks.Add("NmbrJours", rNmbrJours);
            doc.Bookmarks.Add("DateD", rDateD);
            doc.Bookmarks.Add("DateR", rDateR);
            //doc.Bookmarks.Add("Telephone", rTelephone);
            //doc.Bookmarks.Add("Adresse", rAdresse);
            doc.Bookmarks.Add("Date", rDate);



            app.Documents.Open(Application.StartupPath + @"\word\Modèle Ordre de mission 2020.docx");
        }
    }
}
