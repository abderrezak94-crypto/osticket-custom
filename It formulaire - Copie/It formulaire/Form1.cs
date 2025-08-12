using Microsoft.VisualBasic.ApplicationServices;

namespace It_formulaire
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(DemandeCongeRecuperation.Instance))
            {
                panel1.Controls.Add(DemandeCongeRecuperation.Instance);
                DemandeCongeRecuperation.Instance.Dock = DockStyle.Fill;
                DemandeCongeRecuperation.Instance.BringToFront();
            }
            else
            {
                DemandeCongeRecuperation.Instance.BringToFront();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(Decharge.Instance))
            {
                panel1.Controls.Add(Decharge.Instance);
                Decharge.Instance.Dock = DockStyle.Fill;
                Decharge.Instance.BringToFront();
            }
            else
            {
                Decharge.Instance.BringToFront();
            }
        }

        private void Holding_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(AutorisationAbsence.Instance))
            {
                panel1.Controls.Add(AutorisationAbsence.Instance);
                AutorisationAbsence.Instance.Dock = DockStyle.Fill;
                AutorisationAbsence.Instance.BringToFront();
            }
            else
            {
                AutorisationAbsence.Instance.BringToFront();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(DechargeRestitution.Instance))
            {
                panel1.Controls.Add(DechargeRestitution.Instance);
                DechargeRestitution.Instance.Dock = DockStyle.Fill;
                DechargeRestitution.Instance.BringToFront();
            }
            else
            {
                DechargeRestitution.Instance.BringToFront();
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(OrdreMission.Instance))
            {
                panel1.Controls.Add(OrdreMission.Instance);
                OrdreMission.Instance.Dock = DockStyle.Fill;
                OrdreMission.Instance.BringToFront();
            }
            else
            {
                OrdreMission.Instance.BringToFront();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(DechargePC.Instance))
            {
                panel1.Controls.Add(DechargePC.Instance);
                DechargePC.Instance.Dock = DockStyle.Fill;
                DechargePC.Instance.BringToFront();
            }
            else
            {
                DechargePC.Instance.BringToFront();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(DechargeRestitutionPCc.Instance))
            {
                panel1.Controls.Add(DechargeRestitutionPCc.Instance);
                DechargeRestitutionPCc.Instance.Dock = DockStyle.Fill;
                DechargeRestitutionPCc.Instance.BringToFront();
            }
            else
            {
                DechargeRestitutionPCc.Instance.BringToFront();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!panel1.Controls.Contains(recupe.Instance))
            {
                panel1.Controls.Add(recupe.Instance);
                recupe.Instance.Dock = DockStyle.Fill;
                recupe.Instance.BringToFront();
            }
            else
            {
                recupe.Instance.BringToFront();
            }
        }
    }
}