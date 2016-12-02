using System;
using SAPbouiCOM;


namespace GedAddon
{
    public class SboFlatButton
    {
        private SAPbouiCOM.Form ownerForm; // Form a que pertence o botão

        private int ownerTab; // Aba a que pertence o botão

        private String name;

        private String imgPressed;

        private String imgReleased;

        private String imgDisabled;

        private String currentImage;

        private Boolean enabled;

        private Boolean pressed;


        public Boolean Enabled
        {
            get { return enabled; }
        }

        public Boolean Pressed
        {
            get { return pressed; }
        }

        /// <summary>
        /// Chamar os métodos SetButtonImages e SetState após o construtor nesta ordem
        /// </summary>
        public SboFlatButton(SAPbouiCOM.Form ownerForm, int ownerTab, String name, int top, int left)
        {
            this.ownerForm = ownerForm;
            this.ownerTab = ownerTab;
            this.name = name;
            this.enabled = false;
            this.pressed = false;

            SAPbouiCOM.Item imageItem = ownerForm.Items.Add(name, BoFormItemTypes.it_PICTURE);
            imageItem.FromPane = ownerTab;
            imageItem.ToPane = ownerTab;
            imageItem.Left = left;
            imageItem.Top = top;
            imageItem.Width = 35;
            imageItem.Height = 35;
            imageItem.Visible = false;
        }

        private void SetImage()
        {
            String lastError;
            String buttonImage = imgReleased;
            if (pressed) buttonImage = imgPressed;
            if (!enabled) buttonImage = imgDisabled;

            // Caso a imagem já esteja setada não é necessário continuar
            if (currentImage == buttonImage) return;

            try
            {
                SAPbouiCOM.Item imageItem = ownerForm.Items.Item(name);
                SAPbouiCOM.PictureBox pictBox = (SAPbouiCOM.PictureBox)imageItem.Specific;
                currentImage = buttonImage;
                pictBox.Picture = currentImage;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        public void SetButtonImages(String imgPressed, String imgReleased, String imgDisabled)
        {
            this.imgPressed = imgPressed;
            this.imgReleased = imgReleased;
            this.imgDisabled = imgDisabled;
        }

        public void SetState(Boolean enabled)
        {
            this.enabled = enabled;
            SetImage();
        }

        public void Press()
        {
            // Aborta caso o botão esteja desabilitado
            if (!enabled) return;

            // Aborta caso já esteja pressionado
            if (pressed) return;

            pressed = true;
            SetImage();
        }

        public void Release()
        {
            // Aborta caso o botão esteja desabilitado
            if (!enabled) return;

            // Aborta caso já esteja "despressionado"
            if (!pressed) return;

            pressed = false;
            SetImage();
        }
    }

}
