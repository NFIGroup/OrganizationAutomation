using System.AddIn;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
////////////////////////////////////////////////////////////////////////////////
//
// File: WorkspaceAddIn.cs
//
// Comments:
//
// Notes: 
//
// Pre-Conditions: 
//
////////////////////////////////////////////////////////////////////////////////
namespace OrganizationAutomation
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        /// <summary>
        /// The current workspace record context.
        /// </summary>
        private IRecordContext _recordContext;
        private Label label1;
        public IOrganization _OrgRecord;

        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext)
        {
            if (!inDesignMode)
            {
                _recordContext = RecordContext;
            }
            else
            {
                InitializeComponent();
            }
        }

        /// <summary>
        /// Method called by the Add-In framework to initialize it in design mode.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Decimal Conversion";
            this.ResumeLayout(false);

        }

        #region IAddInControl Members

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            _OrgRecord = _recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Organization) as IOrganization;
            if (ActionName.Contains("DecimalConvert"))
            {
                string fieldName = ActionName.Split('@')[1];//Get the field name, which was changed
                string fieldVal = GetFieldValue("CO", fieldName, _OrgRecord);//Get value
                string convertedFieldValue = Decimalconversion(fieldVal, fieldName);//Convert it 2 decimal
                if (convertedFieldValue != "")
                    SetFieldValue("CO", fieldName, convertedFieldValue, _OrgRecord);//set the value
                _recordContext.RefreshWorkspace();//refresh the workspace
            }
        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }
        /// <summary>
        /// Method which is called to get value of a field.
        /// </summary>
        /// <param name="fieldName">The name of the custom field.</param>
        /// <returns>Value of the field</returns>
        public static string GetFieldValue(string packageName, string fieldName, IOrganization _OrgRecord)
        {
            string value = "";
            IList<ICustomAttribute> orgCustomAttributes = _OrgRecord.CustomAttributes;
            foreach (ICustomAttribute val in orgCustomAttributes)
            {
                if (val.PackageName == packageName)//if package name matches
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)//if field matches
                    {
                        if (val.GenericField.DataValue.Value != null)
                        {
                            value = val.GenericField.DataValue.Value.ToString();
                            break;
                        }
                    }
                }
            }
            return value;
        }
        /// <summary>
        /// Method which is use to set value to a field using record Context 
        /// </summary>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public void SetFieldValue(string packageName, string fieldName, string value, IOrganization _OrgRecord)
        {
            IList<ICustomAttribute> orgCustomAttributes = _OrgRecord.CustomAttributes;

            foreach (ICustomAttribute val in orgCustomAttributes)
            {
                if (val.PackageName == packageName)
                {
                    if (val.GenericField.Name == packageName + "$" + fieldName)
                    {
                        switch (val.GenericField.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:
                                if (value == "1" || value.ToLower() == "true")
                                {
                                    val.GenericField.DataValue.Value = true;
                                }
                                else if (value == "0" || value.ToLower() == "false")
                                {
                                    val.GenericField.DataValue.Value = false;
                                }
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.INTEGER:
                                if (value.Trim() == "" || value.Trim() == null)
                                {
                                    val.GenericField.DataValue.Value = null;
                                }
                                else
                                {
                                    val.GenericField.DataValue.Value = Convert.ToInt32(value);
                                }
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                val.GenericField.DataValue.Value = value;
                                break;
                            case RightNow.AddIns.Common.DataTypeEnum.DATETIME:
                                val.GenericField.DataValue.Value = Convert.ToDateTime(value);
                                break;
                        }
                    }
                }
            }
            return;
        }
        /// <summary>
        /// Function to convert string value to decimal upto 2 digit
        /// </summary>
        /// <param name="val"></param>
        /// <param name="fieldName"></param>
        /// <returns>Decimal Value upto 2 digit</returns>
        public string Decimalconversion(string val, string fieldName)
        {
            Decimal output = 0;
            try
            {
                if (Regex.Matches(val, @"[a-zA-Z]").Count > 0)
                {
                    MessageBox.Show("Value in field " + fieldName + " " + "should be Numeric", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else
                {
                    if (!(val.Contains(".")))
                    {
                        val = val + ".00";
                    }
                    if ((val.Contains(".")))
                    {
                        int length = val.Substring(val.IndexOf(".")).Length;
                        if (length == 2)
                        {
                            val = val + "0";
                        }
                    }
                    output = Math.Round(Convert.ToDecimal(val), 2);
                    return output.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return "";
        }
        #endregion
    }

    [AddIn("Workspace Factory AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members

        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Organization Automation"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "Converts numbers to 2 decimals"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            return true;
        }

        #endregion
    }
}