namespace PowerPointQuestionnaire
{
    partial class QuestionnaireRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public QuestionnaireRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab = this.Factory.CreateRibbonTab();
            this.authGroup = this.Factory.CreateRibbonGroup();
            this.usernameLabel = this.Factory.CreateRibbonLabel();
            this.loginButton = this.Factory.CreateRibbonButton();
            this.slideOperationGroup = this.Factory.CreateRibbonGroup();
            this.addNewQuestionnaireSlideButton = this.Factory.CreateRibbonButton();
            this.setSlideButton = this.Factory.CreateRibbonButton();
            this.buttonCancel = this.Factory.CreateRibbonButton();
            this.errorLabel = this.Factory.CreateRibbonLabel();
            this.tab.SuspendLayout();
            this.authGroup.SuspendLayout();
            this.slideOperationGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab
            // 
            this.tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab.Groups.Add(this.authGroup);
            this.tab.Groups.Add(this.slideOperationGroup);
            this.tab.Label = "PPT问卷";
            this.tab.Name = "tab";
            // 
            // authGroup
            // 
            this.authGroup.Items.Add(this.usernameLabel);
            this.authGroup.Items.Add(this.loginButton);
            this.authGroup.Label = "认证/登陆";
            this.authGroup.Name = "authGroup";
            // 
            // usernameLabel
            // 
            this.usernameLabel.Label = "您还没有登录";
            this.usernameLabel.Name = "usernameLabel";
            // 
            // loginButton
            // 
            this.loginButton.Label = "登陆";
            this.loginButton.Name = "loginButton";
            this.loginButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loginButton_Click);
            // 
            // slideOperationGroup
            // 
            this.slideOperationGroup.Items.Add(this.addNewQuestionnaireSlideButton);
            this.slideOperationGroup.Items.Add(this.setSlideButton);
            this.slideOperationGroup.Items.Add(this.buttonCancel);
            this.slideOperationGroup.Items.Add(this.errorLabel);
            this.slideOperationGroup.Label = "设置";
            this.slideOperationGroup.Name = "slideOperationGroup";
            this.slideOperationGroup.Visible = false;
            // 
            // addNewQuestionnaireSlideButton
            // 
            this.addNewQuestionnaireSlideButton.Label = "添加新问卷页";
            this.addNewQuestionnaireSlideButton.Name = "addNewQuestionnaireSlideButton";
            this.addNewQuestionnaireSlideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addNewQuestionnaireSlideButton_Click);
            // 
            // setSlideButton
            // 
            this.setSlideButton.Label = "设置当前页为问卷";
            this.setSlideButton.Name = "setSlideButton";
            this.setSlideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.setSlideButton_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Label = "取消问卷标记";
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCancel_Click);
            // 
            // errorLabel
            // 
            this.errorLabel.Label = " ";
            this.errorLabel.Name = "errorLabel";
            // 
            // QuestionnaireRibbon
            // 
            this.Name = "QuestionnaireRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.QuestionnaireRibbon_Load);
            this.tab.ResumeLayout(false);
            this.tab.PerformLayout();
            this.authGroup.ResumeLayout(false);
            this.authGroup.PerformLayout();
            this.slideOperationGroup.ResumeLayout(false);
            this.slideOperationGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup authGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginButton;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup slideOperationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton setSlideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addNewQuestionnaireSlideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCancel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel usernameLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel errorLabel;
    }

    partial class ThisRibbonCollection
    {
        internal QuestionnaireRibbon QuestionnaireRibbon
        {
            get { return this.GetRibbon<QuestionnaireRibbon>(); }
        }
    }
}
