﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Diagnostics;

namespace ComarchXL
{
    /// <summary>
    /// Interaction logic for ComarchImport.xaml
    /// </summary>
    public partial class ComarchImport : Window
    {
        public ComarchImport()
        {
            InitializeComponent();
        }

        private void cmdExit_Click(object sender, RoutedEventArgs e)
        {

        }

        int SessionID = -1;


        private void cmdLogin_Click(object sender, RoutedEventArgs e)
        {

            ComarchTools.zaloguj(ref SessionID);
            Sesja.Content = SessionID.ToString();
        }

        private void cmdLogout_Click(object sender, RoutedEventArgs e)
        {
            if (SessionID > 0)
                ComarchTools.wyloguj(ref SessionID);
            Sesja.Content = SessionID.ToString();
        }


        private void cmdDodajTowar_Click(object sender, RoutedEventArgs e)
        {
            string kod = "PPP002";
            string nazwa = "PPP002";
            string jm = "szt";
            int TowarID = -1;

            ComarchTools.nowyProduct(SessionID, ref TowarID, kod, nazwa, jm, 1, "IFS");
        }

        private void cmdDodajBOM_Click(object sender, RoutedEventArgs e)
        {

            int IDReceptury = -1;
            string kodReceptury = "recepturaX";
            string towarKod = "PPP002";
            string towarIlosc = "1";
            ComarchTools.nowaReceptura(SessionID, ref IDReceptury, kodReceptury, towarKod, towarIlosc);

            string skladnikKod = "SAP001";
            string skladnikIlosc = "1";
            ComarchTools.nowySkladnik(ref IDReceptury, skladnikKod, skladnikIlosc);

            string skladnikKod2 = "SAP002";
            string skladnikIlosc2 = "1";
            ComarchTools.nowySkladnik(ref IDReceptury, skladnikKod2, skladnikIlosc2);

            ComarchTools.zamkniecieReceptury(IDReceptury);
        }

        private void cmdImportPozycjeGlowne_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujPozycjeGlowne(SessionID);
        }

        private void cmdImportBieznikowB_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "BIEŻNIK BOK");
        }

        private void ImportTest_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importTest(SessionID);
        } 

        private void cmdImportDrutowek_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "DRUTOWKA");
        }

        private void cmdImportApexow_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "APEX");
        }

        private void cmdImportKap_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID,"KAPA OPONY");
        }

        private void cmdImportKordyCiete_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "KORD GUMOWANY CIĘTY");
        }

        private void cmdImportOponySurowe_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "OPONA SUROWA");
        }

        private void cmdImportOponyWulkanizowane_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "OPONA WULKANIZOWANA");
        }



        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = -2;     // -10...10

            // Synchronous
            // synthesizer.Speak("pocałuj się w dupe komar h");

            // Asynchronous
            //synthesizer.SpeakAsync("zaloguj się najpierw");
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {

        }

        private void Window_Closed(object sender, EventArgs e)
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = -2;     // -10...10    
                                       //synthesizer.SpeakAsync("spadam stąd");

            ComarchTools.wyloguj(ref SessionID);
        }

        // commands
        private void NewCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = (SessionID > 0);
        }

        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            MessageBox.Show("The New command was invoked");
        }

        private void ExitCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = (SessionID > 0);
        }

        private void ExitCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            MessageBox.Show("The Exit command was invoked");
            Application.Current.Shutdown();

        }

        private void cmdImportTkaninaOchronna_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "TKANINA OCHRONNA");
        }

        private void cmdImportApexyDrutowki_Click(object sender, RoutedEventArgs e)
        {
            ComarchTools.importujBomy(SessionID, "APEX+DRUT");
        }

        private void cmdKillProc_Click(object sender, RoutedEventArgs e)
        {
           

            Process[] _proceses = null;
            _proceses = Process.GetProcessesByName("calc");
            foreach (Process proces in _proceses)
            {
                proces.Kill();
            }
        }

        private void cmdCalcc_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("calc");
        }

    }

  

}
