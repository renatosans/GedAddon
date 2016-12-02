using System;
using System.Collections.Generic;
using SAPbouiCOM;
using DocMageFramework.Parsing;
using DocMageFramework.FileUtils;
using DocMageFramework.DataManipulation;


namespace GedAddon
{
    public class AddonController
    {
        private GuiController guiController;

        private String lastError;

        private Boolean isAttached;

        public Boolean IsAttached // Acoplamento a interface do SAP B1 bem sucedida
        {
            get { return isAttached; }
        }


        public AddonController()
        {
            guiController = new GuiController();
            guiController.InitializeGui(AttachToSBO());
        }

        private SAPbouiCOM.Application AttachToSBO()
        {
            SAPbouiCOM.Application sboApplication = GetSBOApplication();
            if (sboApplication == null)
            {
                isAttached = false;
                return null;
            }

            sboApplication.AppEvent += new _IApplicationEvents_AppEventEventHandler(SboApplicationEvent);
            sboApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SboMenuEvent);
            sboApplication.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SboFormDataEvent);
            sboApplication.ItemEvent += new _IApplicationEvents_ItemEventEventHandler(SboItemEvent);

            // Adiciona um filtro aos eventos para escolher apenas os tipos que interessam a aplicação
            EventFilters eventFilters = new EventFilters();
            EventFilter newFilter = eventFilters.Add(BoEventTypes.et_ALL_EVENTS);
            newFilter.AddEx("134");
            newFilter.AddEx("customForm");
            newFilter.AddEx("openDialog");
            sboApplication.SetFilter(eventFilters);

            isAttached = true;
            return sboApplication;
        }

        private void SboApplicationEvent(BoAppEventTypes eventType)
        {
            // Monitora o evento de shut down da aplicação
            if (eventType == BoAppEventTypes.aet_ShutDown)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void SboMenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.MenuEvent(ref pVal, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SboFormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.FormDataEvent(ref businessObjectInfo, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void SboItemEvent(String formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            try
            {
                guiController.ItemEvent(formUID, ref pVal, out bubbleEvent);
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        // Busca os databases de uma empresa no BD
        private List<String> GetCompanyDatabases(String companyName)
        {
            // Estados de um database: 0 = ONLINE    1 = RESTORING    2 = RECOVERING    3 = RECOVERY_PENDING
            //                         4 = SUSPECT   5 = EMERGENCY    6 = OFFLINE
            String xmlLocation = FileResource.MapDesktopResource(@"Xml\DataAccess.xml");
            DataAccess dataAccess = DataAccess.Instance;
            dataAccess.MountConnection(xmlLocation, "master");
            dataAccess.OpenConnection();
            DBQuery dbQuery = new DBQuery(dataAccess.GetConnection());
            dbQuery.Query = "SELECT name FROM sys.databases WHERE state = 0";
            dbQuery.Execute(true);
            List<Object> databaseList = dbQuery.ExtractFromResultset(new String[] { "name" });
            List<String> databaseNames = new List<String>();
            foreach (Object[] database in databaseList)
            {
                String databaseName = (String)database[0];
                databaseNames.Add(databaseName);
            }
            List<String> sboDatabases = new List<String>();
            foreach (String databaseName in databaseNames)
            {
                dbQuery.Query = "USE [" + databaseName + "]";
                dbQuery.Execute(false);
                dbQuery.Query = "DECLARE @objectId INT" + Environment.NewLine +
                                "SELECT @objectId = object_id  FROM sys.tables WHERE name = 'CINF'" + Environment.NewLine +
                                "SELECT @objectId AS objectId";
                dbQuery.Execute(true);
                int? objectId = dbQuery.ExtractFromResultset();
                if (objectId != null) sboDatabases.Add(databaseName);
            }
            List<String> companyDatabases = new List<String>();
            foreach (String sboDatabase in sboDatabases)
            {
                dbQuery.Query = "USE [" + sboDatabase + "]";
                dbQuery.Execute(false);
                dbQuery.Query = "SELECT CompnyName FROM CINF";
                dbQuery.Execute(true);
                List<Object> resultSet = dbQuery.ExtractFromResultset(new String[] { "CompnyName" });
                Object[] rowValues = (Object[])resultSet[0];
                String sboCompany = (String)rowValues[0];
                if (sboCompany == companyName) companyDatabases.Add(sboDatabase);
            }
            dataAccess.CloseConnection();

            return companyDatabases;
        }

        // Obtem o objeto de aplicação associado ao SAP Business One
        private SAPbouiCOM.Application GetSBOApplication()
        {
            String[] args = ArgumentParser.ParseCommandLine(Environment.CommandLine);
            String connectionString = args[1];

            SboGuiApi sboGuiApi = new SboGuiApi();
            try
            {
                sboGuiApi.Connect(connectionString);
            }
            catch
            {
                return null;
            }

            return sboGuiApi.GetApplication(-1);
        }
    }

}
