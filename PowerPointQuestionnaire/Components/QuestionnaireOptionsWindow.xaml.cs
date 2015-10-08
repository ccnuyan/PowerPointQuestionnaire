using System.Windows;
using System.Windows.Controls;
using PowerPointQuestionnaire.Interfaces;
using PowerPointQuestionnaire.Model;
using PowerPointQuestionnaire.Services;

namespace PowerPointQuestionnaire.Components
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class QuestionnaireOptionsWindow : Window
    {

        public QuestionnaireOptionsWindow()
        {
            InitializeComponent();
            ChoicesComboBox.Items.Add(2);
            ChoicesComboBox.Items.Add(3);
            ChoicesComboBox.Items.Add(4);
            ChoicesComboBox.Items.Add(5);
            ChoicesComboBox.Items.Add(6);
            ChoicesComboBox.Items.Add(7);
            ChoicesComboBox.Items.Add(8);
            ChoicesComboBox.SelectedIndex = 2;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            //TODO validation 

            if (ChoicesComboBox.SelectedItem != null)
            {
                DialogResult = true;
                Close();
            }
            else
            {
                ErrorTextBox.Text = "请至少提供选项个数";
            }
        }

        public void Initialize(QuestionnaireModel questionnaire)
        {
            ChoicesComboBox.SelectedItem = questionnaire.choices;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
