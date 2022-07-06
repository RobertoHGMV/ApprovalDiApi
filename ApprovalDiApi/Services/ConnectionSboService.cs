using ApprovalDiApi.Models;
using SAPbobsCOM;

namespace ApprovalDiApi.Services
{
    public class ConnectionSboService
    {
        public Company Company { get; private set; }

        public ConnectionSboService()
        {
            Company = new Company();
        }

        public Company ConnectSbo(ConnectionSbo conSbo)
        {
            Company.Server = conSbo.Server;
            Company.CompanyDB = conSbo.CompanyDB;
            Company.DbUserName = conSbo.DbUserName;
            Company.DbPassword = conSbo.DbPassword;
            Company.DbServerType = conSbo.DbServerType;
            Company.UserName = conSbo.UserName;
            Company.Password = conSbo.Password;

            Company.UseTrusted = false;
            Company.language = BoSuppLangs.ln_Portuguese_Br;
            Company.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;

            if (Company.Connect() != 0)
                throw new Exception("Erro ao conectar no sap:" +
                           $"[{Company.GetLastErrorCode()}] - [{Company.GetLastErrorDescription()}]");

            return Company;
        }
    }
}
