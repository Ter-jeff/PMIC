﻿using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MyWpf.Controls
{
    public partial class TextBoxButton
    {
        public TextBoxButton()
        {
            ImageSource = new BitmapImage(new Uri("pack://application:,,,/MyWpf;component/Resource/shell32_3191.ico"));
            InitializeComponent();
        }

        public event EventHandler Click;
        public event EventHandler TextChanged;

        public ImageSource ImageSource
        {
            get { return (ImageSource)GetValue(ImageSourceProperty); }
            set { SetValue(ImageSourceProperty, value); }
        }

        // Using a DependencyProperty as the backing store for ImageSource.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ImageSourceProperty =
            DependencyProperty.Register("ImageSource", typeof(ImageSource), typeof(TextBoxButton), new PropertyMetadata(null));
        public Orientation Orientation
        {
            get { return (Orientation)GetValue(OrientationProperty); }
            set { SetValue(OrientationProperty, value); }
        }

        public static readonly DependencyProperty OrientationProperty =
              DependencyProperty.Register("Orientation", typeof(Orientation), typeof(TextBoxButton), new PropertyMetadata(null));

        public string Header
        {
            get { return (string)GetValue(HeaderProperty); }
            set { SetValue(HeaderProperty, value); }
        }

        public static readonly DependencyProperty HeaderProperty =
              DependencyProperty.Register("Header", typeof(string), typeof(TextBoxButton), new PropertyMetadata(null));

        public string Text
        {
            get { return (string)GetValue(TextProperty); }
            set { SetValue(TextProperty, value); }
        }

        public static readonly DependencyProperty TextProperty =
              DependencyProperty.Register("Text", typeof(string), typeof(TextBoxButton), new PropertyMetadata(null));

        private void TextBoxButtonTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextChanged != null)
                TextChanged(this, EventArgs.Empty);
        }

        private void TextBoxButtonButton_Click(object sender, RoutedEventArgs e)
        {
            if (Click != null)
                Click(this, EventArgs.Empty);
        }
    }
}