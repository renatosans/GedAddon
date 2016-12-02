using System;
using Microsoft.SharePoint.Client;


namespace SystemIntegration
{
    public class DocLibField
    {
        private String displayName;

        private FieldType fieldType;

        private String[] fieldValues;

        private String internalName;

        public String DisplayName
        {
            get { return displayName; }
        }

        public FieldType FieldType
        {
            get { return fieldType; }
        }

        public String[] FieldValues
        {
            get { return fieldValues; }
        }

        public String InternalName
        {
            get { return internalName; }
        }


        public DocLibField(String displayName, FieldType fieldType, String[] fieldValues, String internalName)
        {
            this.displayName = displayName;
            this.fieldType = fieldType;
            this.fieldValues = fieldValues;
            this.internalName = internalName;
        }
    }

}
