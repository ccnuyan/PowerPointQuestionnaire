using System;
using System.Diagnostics;
using System.Net;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Tools.Ribbon;
using PowerPointQuestionnaire.Components;
using PowerPointQuestionnaire.Controls;
using PowerPointQuestionnaire.Interfaces;
using PowerPointQuestionnaire.Model;
using PowerPointQuestionnaire.Services;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointQuestionnaire
{
    public partial class QuestionnaireRibbon
    {

#if DEBUG
        const string baseUrl = "http://localhost.:3000/";
#else
        const string baseUrl= "http://www.iccnu.net/";
#endif

        private LoginWindow _loginWindow;

        private const string ChoiceString = "ABCDEFGH";

        private PowerPoint.Slide _selectedSlide;

        private readonly IQuestionnaireUtil _questionnaireUtil = new QuestionnaireUtil();

        private void loginButton_Click(object sender, RibbonControlEventArgs e)
        {
            TryLogin();
        }

        private void QuestionnaireRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            RefreshSetAndAdd();

            AppWapper.App.SlideSelectionChanged += App_SlideSelectionChanged;
            AppWapper.App.SlideShowNextSlide += App_SlideShowNextSlide;
            AppWapper.App.PresentationOpen += App_PresentationOpen;


            _loginWindow = new LoginWindow();
            _loginWindow.Closed += LoginWindowClosed;
        }

        private void App_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            var index = Wn.View.CurrentShowPosition;

            var slide = Wn.Presentation.Slides[index];

            GoToSlidePage(slide);
        }

        private void LoginWindowClosed(object sender, EventArgs e)
        {
            _loginWindow = new LoginWindow();
            _loginWindow.Closed += LoginWindowClosed;
        }

        private void App_PresentationOpen(PowerPoint.Presentation Pres)
        {
            RefreshSetAndAdd();
        }

        private void TryLogin()
        {
            if (_loginWindow.ShowDialog() == true)
            {
                slideOperationGroup.Visible = true;
                loginButton.Visible = false;
                usernameLabel.Label = @"欢迎您，
" + AuthService.Me.displayName;
            }
            else
            {
                slideOperationGroup.Visible = false;
                loginButton.Visible = true;
                usernameLabel.Label = "您尚未登陆";
            }

            RefreshSetAndAdd();
        }


        public void RefreshSetAndAdd()
        {
            var slide = _selectedSlide;
            questionnairePageButton.Visible = false;

            if (AuthService.Me == null || slide == null)
            {
                setSlideButton.Visible = false;
                buttonCancel.Visible = false;
                return;
            }

            if (_questionnaireUtil.Check(slide))
            {
                var questionnaire = _questionnaireUtil.Deserialize(slide);

                if (questionnaire.user != AuthService.Me._id.ToString())
                {
                    setSlideButton.Visible = false;
                    buttonCancel.Visible = false;

                    errorLabel.Label = "该页的问卷不属于你";
                }
                else
                {
                    questionnairePageButton.Visible = true;

                    setSlideButton.Visible = true;
                    buttonCancel.Visible = true;

                    setSlideButton.Label = "重新设置问卷";

                    errorLabel.Label = " ";
                }
            }
            else
            {
                setSlideButton.Visible = true;
                buttonCancel.Visible = false;

                setSlideButton.Label = "将当前选中页设置为问卷";
            }
        }

        private PowerPoint.Slide AddSlideByQuestionnaire(QuestionnaireModel questionnaire)
        {
            var index = Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1;

            var slide = AppWapper.App.ActivePresentation.Slides.Add(index, PowerPoint.PpSlideLayout.ppLayoutBlank);

            var textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 600, 50);//向当前PPT添加文本框
            textbox.TextFrame.TextRange.Text = "调查:";
            textbox.TextFrame.TextRange.Font.Size = 48;//设置文本字体大小

            for (var i = 0; i < questionnaire.choices; i++)
            {
                var choiceTextbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 120 + 40 * i, 400, 50);//向当前PPT添加文本框
                choiceTextbox.TextFrame.TextRange.Text = ChoiceString[i] + ":";
                choiceTextbox.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小
            }

            _selectedSlide = slide;

            return _selectedSlide;
        }

        private async void CreateQuestionnaireSlideRecord(PowerPoint.Slide slide, QuestionnaireModel questionnaire)
        {
            try
            {
                var created = await _questionnaireUtil.CreateAsync(questionnaire);
                _questionnaireUtil.Mark(slide, created);
            }
            catch (WebException)
            {
                slide.Delete();
                MessageBox.Show("问卷添加失败了, 这可能是个网络错误.");
            }
        }

        private void App_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            if (SldRange.Count != 1)
            {
                _selectedSlide = null;
            }
            else
            {
                _selectedSlide = SldRange[1];
            }
            RefreshSetAndAdd();
        }

        private void addNewSlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            var window = new QuestionnaireOptionsWindow();
            if (window.ShowDialog() == true)
            {
                var questionnaire = new QuestionnaireModel
                {
                    choices = (int)window.ChoicesComboBox.SelectionBoxItem
                };

                var slide = AddSlideByQuestionnaire(questionnaire);

                CreateQuestionnaireSlideRecord(slide, questionnaire);
            }
        }

        private void setSlideButton_Click(object sender, RibbonControlEventArgs e)
        {
            var slide = _selectedSlide;

            var window = new QuestionnaireOptionsWindow();
            if (_questionnaireUtil.Check(slide))
            {
                window.Initialize(_questionnaireUtil.Deserialize(slide));
            }

            if (window.ShowDialog() == true)
            {
                var choices = (int) window.ChoicesComboBox.SelectionBoxItem;
                if (!_questionnaireUtil.Check(slide))
                {
                    //CreateQuestionnaireSlideRecord
                    var questionnaire = new QuestionnaireModel()
                    {
                        choices = choices
                    };

                    try
                    {
                        CreateQuestionnaireSlideRecord(slide, questionnaire);
                    }
                    catch
                    {
                        _questionnaireUtil.Unmark(slide);
                        MessageBox.Show("问卷设置失败了,这可能是个网络错误");
                    }
                    finally
                    {
                        RefreshSetAndAdd();
                    }
                }
                else
                {
                    //UpdateSelectedSlide
                    var questionnaire = _questionnaireUtil.Deserialize(slide);
                    questionnaire.choices = choices;

                    _questionnaireUtil.Mark(slide, questionnaire);
                }

            }
        }

        private async void buttonCancel_Click(object sender, RibbonControlEventArgs e)
        {
            var questionnaire = _questionnaireUtil.Deserialize(_selectedSlide);
            try
            {
                await _questionnaireUtil.DeleteAsync(questionnaire.id);
            }
            catch
            {
                Console.WriteLine("delete failed");
            }
            finally
            {
                _questionnaireUtil.Unmark(_selectedSlide);
                RefreshSetAndAdd();
            }
        }

        private void homePageButton_Click(object sender, RibbonControlEventArgs e)
        {
            Process.Start(baseUrl);
        }

        private void questionnairePageButton_Click(object sender, RibbonControlEventArgs e)
        {
            var slide = _selectedSlide;

            GoToSlidePage(slide);
        }

        private async void GoToSlidePage(PowerPoint.Slide slide)
        {
            if (!_questionnaireUtil.Check(slide))
            {
                //MessageBox.Show("无效的问卷.");
                return;
            }

            var questionnaire = _questionnaireUtil.Deserialize(slide);

            if (questionnaire.user != AuthService.Me._id.ToString())
            {
                //MessageBox.Show("这个问卷不是你创建的.");
                return;
            }

            try
            {
                var tempFile = await _questionnaireUtil.UploadAsync(slide);
                questionnaire.tempFileId = tempFile._id;
                var updated = await _questionnaireUtil.UpdateAsync(questionnaire);

                Process.Start(baseUrl + "questionnaires/" + updated.id);
            }
            catch (Exception e)
            {
                MessageBox.Show("网络错误: "+e.Message);
            }
        }
    }
}
