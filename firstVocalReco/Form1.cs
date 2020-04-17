using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Speech.Recognition;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace firstVocalReco
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
         

        public Form1()
        {
            InitializeComponent();

        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            // On tue TOUT les process excel existant sur le poste
            KillSpecificExcelFileProcess(""); // appel une void stocker un eu plus lmoin


           //Ouvrir le fichier excel avec la grammaire et les actions
           // ou lance l'interop Excel
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
           //on ouvre le fichier dont on a besoin
           Workbook MonClasseur = xlApp.Workbooks.Open(@"C:\\BddRecoVocal.xlsx");
           Worksheet MaFeuil = (Worksheet)MonClasseur.Worksheets["Contact"];
            //nombre de boucle à faire pour créer grammaire


            Choices commands = new Choices();
                        //boucle création liste grammaire

             string NbVal = MaFeuil.Cells[1,1].value;
            // int NbPhrase = Convert.ToInt32(NbVal) + 2;
            int NbPhrase = 5;
            int compteur;
            for (compteur = 2; compteur < NbPhrase; compteur++)
            {
                string Phrase = MaFeuil.Cells[compteur,1].value;
                commands.Add(new string[] { Phrase });
            }

            
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
            
        }

        // Action quadn on ferme la frome
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Etes-vous certain de vouloir quitter ?", "Quitter", MessageBoxButtons.YesNo) == DialogResult.No)
                e.Cancel = true;
            //on tue tout les processus excel pour vider la mémoire, bon ok Excel c'st nul pour ca met c génial pour les calculs auto
            KillSpecificExcelFileProcess("");
        }



        void recEngine_SpeechRecognized(object  sender, SpeechRecognizedEventArgs e)
        {

            richTextBox1.Text += "\n";
            richTextBox1.Text += e.Result.Text;


            // On envoi le texte dans excel, on calcule la feuille et on voit ce que ca donne

            // Action en fcontion de la phrase et c'est la que excel rentre en jeu
            switch (e.Result.Text)
            {
                case "Bonjour John":
                   MessageBox.Show("Que puis je faire pour toi?");

                    // On crée et charge une nouvelle grammaire
                    Choices commands = new Choices();
                    commands.Add(new string[] { "Ajouter un truc", "comment vas tu", "ajouter à la liste", "Ouvrir un fichier", "je suis une grosse saucisse" });
                    GrammarBuilder gBuilder = new GrammarBuilder();
                    gBuilder.Append(commands);
                    Grammar grammar = new Grammar(gBuilder);

                    // et on relance
                    recEngine.LoadGrammarAsync(grammar);
                    recEngine.SpeechRecognized += recEngine_SpeechRecognized;
                    
                    break;
                case "print my name":
                   richTextBox1.Text += "\n Guihlem";
                   break;
            }
        }



        // gestion bouton démarrage et arret
        private void btnEnable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            btnDisable.Enabled = true;
        }        

        private void btnDisable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btnDisable.Enabled = false;
        }

        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
               // if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }


    }
}
