﻿using System;
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

namespace SbSCh2_4
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void checkBox1_Checked(object sender, RoutedEventArgs e)
        {
            textBlock1.Text = checkBox1.IsChecked.ToString();

        }

        private void checkBox1_Unchecked(object sender, RoutedEventArgs e)
        {
            textBlock1.Text = checkBox1.IsChecked.ToString();
        }

        private void checkBox1_Click(object sender, RoutedEventArgs e)
        {
            textBlock1.Text = checkBox1.IsChecked.ToString();
        }
    }
}
