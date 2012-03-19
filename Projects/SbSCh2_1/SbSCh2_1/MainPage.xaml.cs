using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace SbSCh2_1
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            button1.Content = "0";
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            int n = Convert.ToInt16(button1.Content.ToString());
            n++;
            button1.Content = n;
        }

        private void addIt(object sender, RoutedEventArgs e)
        {
            int n = Convert.ToInt16(button1.Content.ToString());
            n+=2;
            button1.Content = n;
        }

        private void button1_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
