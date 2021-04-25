using System;
using LSEXT;
using LSSERVICEPROVIDERLib;

namespace MockNautilus
{
    public class MockExtensionParameters : LSEXT.LSExtensionParameters
    {
        public string Product { get; set; }
        public string Version { get; set; }
        public string Server { get; set; }
        public string ComputerName { get; set; }
        public string HKeyCurrentUser { get; set; }
        public string HKeyLocalMachine { get; set; }
        public string OperatorName { get; set; }
        public string RoleName { get; set; }
        public decimal OperatorId { get; set; }
        public decimal RoleId { get; set; }
        public decimal WorkstationId { get; set; }
        public decimal SessionId { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string PasswordRole { get; set; }
        public string ServerInfo { get; set; }

        public decimal? EntityId { get; set; }
        public ADODB.Recordset EntityRecords { get; set; }

        public dynamic this[string key]
        {
            get { return GetKey(key); }
        }

        dynamic GetKey(string key)
        {
            if (key == "PRODUCT") return Product;
            if (key == "VERSION") return Version;
            if (key == "SERVER") return Server;
            if (key == "COMPUTER_NAME") return ComputerName;
            if (key == "HKEY_CURRENT_USER") return HKeyCurrentUser;
            if (key == "HKEY_LOCAL_MACHINE") return HKeyLocalMachine;
            if (key == "OPERATOR_NAME") return OperatorName;
            if (key == "ROLE_NAME") return RoleName;
            if (key == "OPERATOR_ID") return OperatorId;
            if (key == "ROLE_ID") return RoleId;
            if (key == "WORKSTATION_ID") return WorkstationId;
            if (key == "SESSION_ID") return SessionId;
            if (key == "USERNAME") return Username;
            if (key == "PASSWORD") return Password;
            if (key == "PASSWORD_ROLE") return PasswordRole;
            if (key == "SERVER_INFO") return ServerInfo;
            if (key == "ENTITY_ID") return EntityId;
            if (key == "RECORDS") return EntityRecords;
            if (key == "SERVICE_PROVIDER") return new MockNautilusServiceProvider();
            throw new ArgumentException();
        }

        void LSEXT.IExtensionParametersEx.Add(string Key, object Value)
        {
            throw new NotImplementedException();
        }

        int LSEXT.IExtensionParametersEx.Count
        {
            get { throw new NotImplementedException(); }
        }

        System.Collections.IEnumerator LSEXT.IExtensionParametersEx.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        LSEXT.LSExtensionParameter LSEXT.IExtensionParametersEx.Parameter(string Key)
        {
            throw new NotImplementedException();
        }

        void LSEXT.IExtensionParametersEx.Remove(string Key)
        {
            throw new NotImplementedException();
        }

        void LSEXT.IExtensionParametersEx.RemoveAll()
        {
            throw new NotImplementedException();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        dynamic IExtensionParametersEx.this[string Key] { get { return GetKey(Key); } set => throw new NotImplementedException(); }
    }


    public class MockNautilusServiceProvider : LSSERVICEPROVIDERLib.NautilusServiceProvider
    {
        public dynamic QueryServiceProvider(string serviceProviderName)
        {
            if (serviceProviderName == "ProcessXML")
                return new MockNautilusProcessXML();
            if (serviceProviderName == "Config")
                return new MockNautilusConfig();
            if (serviceProviderName == "Schema")
                return new MockNautilusSchema();
            if (serviceProviderName == "PopupMenu")
                return new MockNautilusPopupMenu();

            throw new NotImplementedException();
        }
    }


    public class MockNautilusProcessXML : LSSERVICEPROVIDERLib.NautilusProcessXML
    {
        public string ProcessXML(object xmlDoc)
        {
            return string.Empty;
        }

        public string ProcessXMLFile(string filename)
        {
            return string.Empty;
        }

        public string ProcessXMLWithResponse(object xmlDoc, object xmlResponse)
        {
            return string.Empty;
        }

        public string ProcessXMLFileWithResponse(string filename, object xmlResponse)
        {
            return string.Empty;
        }

        public void SetImportOption(ProcessXMLOption importOption, ProcessXMLSetting importSetting)
        {
        }

        public void ResetImportOptions()
        {
        }
    }


    public class MockNautilusConfig : NautilusConfig
    {
        public bool GetBooleanItem(string itemName)
        {
            throw new NotImplementedException();
        }

        public string GetStringItem(string itemName)
        {
            if (itemName == "System\\Extensions\\PES\\Default Export Directory")
                return Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            if (itemName == "System\\Extensions\\PES\\PES Extension Groups")
                return "PES,TTPE,TTTP";

            if (itemName == "System\\Extensions\\PES\\Workflow for Source Sample")
                return "PES Source";

            throw new NotImplementedException();
        }

        public double GetDoubleItem(string itemName)
        {
            throw new NotImplementedException();
        }

        public bool CheckItemExists(string itemName)
        {
            throw new NotImplementedException();
        }

        public bool CheckSecurity(string itemName)
        {
            throw new NotImplementedException();
        }
    }


    public class MockNautilusSchema : NautilusSchema
    {
        public double EntityIdFromName(string entityName)
        {
            throw new NotImplementedException();
        }

        public string EntityNameFromId(double EntityID)
        {
            throw new NotImplementedException();
        }

        public void EntityIconNames(double EntityID, out string selectedIcon, out string unselectedIcon)
        {
            throw new NotImplementedException();
        }

        public double GetEntityTableId(double EntityID)
        {
            throw new NotImplementedException();
        }

        public double GetEntityKeyFieldId(double EntityID)
        {
            throw new NotImplementedException();
        }

        public double GetEntityNameFieldId(double EntityID)
        {
            throw new NotImplementedException();
        }

        public double GetEntityFilterId(double EntityID)
        {
            throw new NotImplementedException();
        }

        public string GetTableDatabaseName(double tableId)
        {
            throw new NotImplementedException();
        }

        public string GetFieldDatabaseName(double fieldId)
        {
            throw new NotImplementedException();
        }

        public double GetTableExtraTableId(double tableId)
        {
            throw new NotImplementedException();
        }

        public double GetTableKeyFieldId(double tableId)
        {
            throw new NotImplementedException();
        }

        public double GetFieldTableId(double fieldId)
        {
            throw new NotImplementedException();
        }

        public bool CheckFieldInTable(string tableName, string fieldName)
        {
            throw new NotImplementedException();
        }

        public double GetTableIdFromDatabaseName(string tableName)
        {
            throw new NotImplementedException();
        }

        public bool GetTableAllowVersioning(double tableId)
        {
            throw new NotImplementedException();
        }

        public string GetEntitySecurityItemPrefix(double EntityID)
        {
            throw new NotImplementedException();
        }
    }


    public class MockNautilusPopupMenu : NautilusPopupMenu
    {
        public void PopupMenu(double EntityID, string recordIds)
        {
            throw new NotImplementedException();
        }

        public void PopupMenuAtPosition(double EntityID, string recordIds, int screenX, int screenY)
        {
            throw new NotImplementedException();
        }

        public void PopupMenuSelectDefault(double EntityID, string recordIds)
        {
            throw new NotImplementedException();
        }

        public void PopupMenuSelectByName(double EntityID, string recordIds, string itemName)
        {
            throw new NotImplementedException();
        }
    }
}