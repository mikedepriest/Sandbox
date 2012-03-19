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

namespace SbSCh2_3
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void rectangle1_MouseMove(object sender, MouseEventArgs e)
        {
            textBox1.Text = e.GetPosition(this).X.ToString();
            textBox2.Text = e.GetPosition(this).Y.ToString();
        }

        private void rectangle1_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            rectangle1.Fill = new SolidColorBrush(Colors.Green);
            e.Handled = true;

        }

        private void rectangle1_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            rectangle1.Fill = new SolidColorBrush(Colors.Yellow);
            e.Handled = true;
        }
    }
}
