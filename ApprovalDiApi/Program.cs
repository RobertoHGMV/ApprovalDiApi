using ApprovalDiApi.Models;
using ApprovalDiApi.Services;
using SAPbobsCOM;
using System;
using System.Collections.Generic;

namespace ApprovalDiApi
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Teste de aprovação de documento com aprovação pelo SAP");

                var company = GetCompany();

                AddDocument(company);

                PrintSuccess();
            }
            catch (Exception ex)
            {
                PrintMessageError(ex.Message);
            }
        }

        private static Company GetCompany()
        {
            var connSboService = new ConnectionSboService();
            var connSbo = GetParams();
            return connSboService.ConnectSbo(connSbo);
        }

        private static void AddDocument(Company company)
        {
            try
            {
                var wtmCode = 4;
                var appService = new ApprovationSboService(company);
                var docService = new DocumentService();

                var document = docService.CreateDocument(company);

                company.StartTransaction();

                appService.ActiveAndCheckApproval(document, new List<int> { wtmCode });

                docService.Add(company, document);

                appService.DisableApprovalTemplates();

                if (company.InTransaction)
                    company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception)
            {
                if (company.InTransaction)
                    company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                throw;
            }
        }

        private static ConnectionSbo GetParams()
        {
            return new ConnectionSbo(
                @"DESKTOP-JHISO7L\B1",
                "SBO_ENGESOFTWARE_TESTE",
                "sa",
                "sap@123",
                BoDataServerTypes.dst_MSSQL2017,
                "manager",
                "sapbone");
        }

        private static void PrintSuccess()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine("Operação realizada com sucesso");
            Console.ReadKey();
        }

        private static void PrintMessageError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(message);
            Console.ResetColor();
            Console.ReadKey();
        }
    }
}
