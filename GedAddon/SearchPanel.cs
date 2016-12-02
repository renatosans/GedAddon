using System;
using System.Globalization;
using System.Collections.Generic;
using SAPbouiCOM;
using SystemIntegration;
using Microsoft.SharePoint.Client;


namespace GedAddon
{
    public class SearchPanel
    {
        private SAPbouiCOM.Form ownerForm; // Form a que pertence o painel de busca

        private int ownerTab; // Aba a que pertence o painel de busca

        private Dictionary<String, String[]> controlDictionary;

        private SAPbouiCOM.Item searchPanel;

        public int Width
        {
            get { return searchPanel.Width; }
        }

        public int Height
        {
            get { return searchPanel.Height; }
        }


        public SearchPanel(SAPbouiCOM.Form ownerForm, int ownerTab, int top, int left, int width, int height)
        {
            this.ownerForm = ownerForm;
            this.ownerTab = ownerTab;
            this.controlDictionary = new Dictionary<String, String[]>();

            searchPanel = ownerForm.Items.Add("pnlSearch", BoFormItemTypes.it_RECTANGLE);
            searchPanel.FromPane = ownerTab;
            searchPanel.ToPane = ownerTab;
            searchPanel.Left = left;
            searchPanel.Top = top;
            searchPanel.BackColor = 7;
            searchPanel.Width = width;
            searchPanel.Height = (int)(height * 0.2); // 20% de ClientHeight
            searchPanel.Visible = false;
        }

        /// <summary>
        /// Renderiza os controles a partir dos campos de busca
        /// </summary>
        public void CreateControls(List<DocLibField> searchFields)
        {
            if ((searchFields == null) || (searchFields.Count < 1)) return;

            int fieldId = 1;
            int fieldPos = 0;
            foreach (DocLibField searchField in searchFields)
            {
                SAPbouiCOM.Item titleItem = ownerForm.Items.Add("title" + fieldId.ToString(), BoFormItemTypes.it_STATIC);
                titleItem.FromPane = ownerTab;
                titleItem.ToPane = ownerTab;
                titleItem.Left = searchPanel.Left + 25 + fieldPos * 125;
                titleItem.Top = searchPanel.Top + 25;
                titleItem.Width = 100;
                titleItem.Height = 20;
                titleItem.Visible = false;
                SAPbouiCOM.StaticText title = (SAPbouiCOM.StaticText)titleItem.Specific;
                title.Caption = searchField.DisplayName;

                if (searchField.FieldType == FieldType.Text)
                {
                    String uniqueId = fieldId.ToString() + "_1";
                    ownerForm.DataSources.UserDataSources.Add("bind" + uniqueId, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                    controlDictionary.Add(searchField.InternalName, new String[] { uniqueId });

                    SAPbouiCOM.Item fieldItem = ownerForm.Items.Add("field" + uniqueId, BoFormItemTypes.it_EDIT);
                    fieldItem.FromPane = ownerTab;
                    fieldItem.ToPane = ownerTab;
                    fieldItem.Left = searchPanel.Left + 25 + fieldPos * 125;
                    fieldItem.Top = searchPanel.Top + 50;
                    fieldItem.Width = 100;
                    fieldItem.Height = 20;
                    fieldItem.Visible = false;
                    SAPbouiCOM.EditText editField = (SAPbouiCOM.EditText)fieldItem.Specific;
                    editField.DataBind.SetBound(true, "", "bind" + uniqueId);
                    fieldPos++;
                }

                if (searchField.FieldType == FieldType.Choice)
                {
                    String uniqueId = fieldId.ToString() + "_1";
                    ownerForm.DataSources.UserDataSources.Add("bind" + uniqueId, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
                    controlDictionary.Add(searchField.InternalName, new String[] { uniqueId });

                    SAPbouiCOM.Item fieldItem = ownerForm.Items.Add("field" + uniqueId, BoFormItemTypes.it_COMBO_BOX);
                    fieldItem.FromPane = ownerTab;
                    fieldItem.ToPane = ownerTab;
                    fieldItem.Left = searchPanel.Left + 25 + fieldPos * 125;
                    fieldItem.Top = searchPanel.Top + 50;
                    fieldItem.Width = 100;
                    fieldItem.Height = 20;
                    fieldItem.Visible = false;
                    SAPbouiCOM.ComboBox comboField = (SAPbouiCOM.ComboBox)fieldItem.Specific;
                    foreach (String value in searchField.FieldValues)
                        comboField.ValidValues.Add(value, null);
                    comboField.DataBind.SetBound(true, "", "bind" + uniqueId);
                    fieldPos++;
                }

                if (searchField.FieldType == FieldType.DateTime)
                {
                    String uniqueId1 = fieldId.ToString() + "_1";
                    String uniqueId2 = fieldId.ToString() + "_2";
                    ownerForm.DataSources.UserDataSources.Add("bind" + uniqueId1, SAPbouiCOM.BoDataType.dt_DATE, 254);
                    ownerForm.DataSources.UserDataSources.Add("bind" + uniqueId2, SAPbouiCOM.BoDataType.dt_DATE, 254);
                    controlDictionary.Add(searchField.InternalName, new String[] { uniqueId1, uniqueId2 });

                    SAPbouiCOM.Item fromItem = ownerForm.Items.Add("field" + uniqueId1, BoFormItemTypes.it_EDIT);
                    fromItem.FromPane = ownerTab;
                    fromItem.ToPane = ownerTab;
                    fromItem.Left = searchPanel.Left + 25 + fieldPos * 125;
                    fromItem.Top = searchPanel.Top + 50;
                    fromItem.Width = 100;
                    fromItem.Height = 20;
                    fromItem.Visible = false;
                    SAPbouiCOM.EditText fromField = (SAPbouiCOM.EditText)fromItem.Specific;
                    fromField.DataBind.SetBound(true, "", "bind" + uniqueId1);
                    fieldPos++;

                    SAPbouiCOM.Item toItem = ownerForm.Items.Add("field" + uniqueId2, BoFormItemTypes.it_EDIT);
                    toItem.FromPane = ownerTab;
                    toItem.ToPane = ownerTab;
                    toItem.Left = searchPanel.Left + 10 + fieldPos * 125;
                    toItem.Top = searchPanel.Top + 50;
                    toItem.Width = 100;
                    toItem.Height = 20;
                    toItem.Visible = false;
                    SAPbouiCOM.EditText toField = (SAPbouiCOM.EditText)toItem.Specific;
                    toField.DataBind.SetBound(true, "", "bind" + uniqueId2);
                    fieldPos++;
                }

                fieldId++;
            } // end foreach
        }

        /// <summary>
        /// Obtem as chaves de busca a partir do que o usuário preencheu no formulário/painel
        /// </summary>
        public Dictionary<String, String[]> GetSearchKeys()
        {
            Dictionary<String, String[]> searchKeys = new Dictionary<String, String[]>();

            SAPbouiCOM.Items formItems = ownerForm.Items;
            SAPbouiCOM.UserDataSources dataSources = ownerForm.DataSources.UserDataSources;
            SAPbouiCOM.Item cardNameItem = ownerForm.Items.Item("5");
            SAPbouiCOM.EditText cardNameSpecific = (SAPbouiCOM.EditText)cardNameItem.Specific;
            searchKeys.Add("Cliente", new String[] { cardNameSpecific.Value });

            foreach (String key in controlDictionary.Keys)
            {
                String fieldName = key;
                String[] controls = controlDictionary[key];
                List<String> fieldData = new List<String>();
                foreach (String uniqueId in controls)
                {
                    UserDataSource dataSource = dataSources.Item("bind" + uniqueId);
                    Item formItem = formItems.Item("field" + uniqueId);
                    if ((dataSource != null) && (formItem != null) && (dataSource.DataType == BoDataType.dt_DATE))
                    {
                        SAPbouiCOM.EditText editField = (SAPbouiCOM.EditText)formItem.Specific;
                        DateTime value; Boolean parsed = DateTime.TryParseExact(editField.Value, "yyyyMMdd", null, DateTimeStyles.None, out value);
                        String strValue = null; if (parsed) strValue = value.ToString("yyyy-MM-ddTHH:mm:ss");
                        fieldData.Add(strValue); // caso não consiga fazer o parse da data insere null no lugar
                    }
                    if ((dataSource != null) && (formItem != null) && (dataSource.DataType == BoDataType.dt_SHORT_TEXT))
                    {
                        if (formItem.Type == BoFormItemTypes.it_COMBO_BOX)
                        {
                            SAPbouiCOM.ComboBox comboField = (SAPbouiCOM.ComboBox)formItem.Specific;
                            String strValue = null; if (comboField.Selected != null) strValue = comboField.Selected.Value;
                            fieldData.Add(strValue); // caso nenhum item esteja selecionado insere null no lugar
                        }
                        if (formItem.Type == BoFormItemTypes.it_EDIT)
                        {
                            SAPbouiCOM.EditText editField = (SAPbouiCOM.EditText)formItem.Specific;
                            fieldData.Add(editField.Value);
                        }
                    }
                }

                Boolean dataFilled = true;
                foreach (String dataItem in fieldData)
                    if (String.IsNullOrEmpty(dataItem)) dataFilled = false;
                if (dataFilled) searchKeys.Add(fieldName, fieldData.ToArray());
            }

            return searchKeys;
        }
    }

}
