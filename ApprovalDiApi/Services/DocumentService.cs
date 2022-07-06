using SAPbobsCOM;
using System;

namespace ApprovalDiApi.Services
{
    public class DocumentService
    {
        public Documents CreateDocument(Company company)
        {
            var document = company.GetBusinessObject(BoObjectTypes.oPurchaseOrders) as Documents;

            var currentDate = DateTime.Now.Date;
            //document.BPL_IDAssignedToInvoice = 1;
            document.CardCode = "COL000001";
            document.CardName = "FORNECEDOR";
            document.TaxDate = currentDate;
            document.DocDueDate = currentDate;
            document.DocType = BoDocumentTypes.dDocument_Items;

            document.Lines.ItemCode = "DOP000001";
            document.Lines.ItemDescription = "ISS RETIDO A RECOLHER";
            document.Lines.Quantity = 1;
            document.Lines.UnitPrice = 1;

            return document;
        }

        public void Add(Company company, Documents document)
        {
            if (document.Add() != 0)
                throw new Exception(GetLastError(company));
        }

        public string GetLastError(Company company)
        {
            var message = "Erro ao adicionar documento no SAP.\n";
            return $"{message}[{company.GetLastErrorCode()}]-[{company.GetLastErrorDescription()}]";
        }
    }
}
